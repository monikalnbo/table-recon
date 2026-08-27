/**
 * pi Agent · Excel（内置 excel-mcp：查询 + 修改）
 *
 * 模型自定义：baseUrl + apiKey + model 全部可配（任何 OpenAI 兼容端点）。
 * Excel 能力：直接 spawn 本仓库 excel-mcp/server.py（标准 MCP stdio 协议），
 * 把 13 个工具动态注册进 pi 会话 —— 不依赖 pi-mcp-adapter，克隆即用。
 * 同时挂 table-recon MCP（若存在），对账语义走 compare_tables。
 *
 * 配置（优先级：环境变量 > agent/config.json > 内置默认）：
 *   EXCEL_AGENT_BASE_URL / EXCEL_AGENT_API_KEY / EXCEL_AGENT_MODEL / EXCEL_AGENT_API
 *
 * 用法：
 *   cp agent/config.example.json agent/config.json   # 填 baseUrl / apiKey / model
 *   npm install
 *   node excel-agent.mjs "查询 测试_B方.xlsx 里实称>700 的单号并汇报"
 */

import { createAgentSession, DefaultResourceLoader, ModelRuntime, getAgentDir } from "@earendil-works/pi-coding-agent";
import { Type } from "@sinclair/typebox";
import { spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/* ---------------- 模型配置 ---------------- */
const def = { provider: "custom-llm", api: "openai-completions", reasoning: false, contextWindow: 131072, maxTokens: 8192 };
let cfg = { ...def };
const cfgFile = path.join(root, "agent", "config.json");
if (fs.existsSync(cfgFile)) cfg = { ...cfg, ...JSON.parse(fs.readFileSync(cfgFile, "utf8")) };
for (const [env, k] of [["EXCEL_AGENT_BASE_URL", "baseUrl"], ["EXCEL_AGENT_API_KEY", "apiKey"], ["EXCEL_AGENT_MODEL", "model"], ["EXCEL_AGENT_API", "api"]]) {
	if (process.env[env]) cfg[k] = process.env[env];
}
if (!cfg.baseUrl || !cfg.apiKey || !cfg.model) {
	console.error("缺少模型配置，二选一：");
	console.error("  1) cp agent/config.example.json agent/config.json 并填 baseUrl/apiKey/model");
	console.error("  2) export EXCEL_AGENT_BASE_URL=... EXCEL_AGENT_API_KEY=... EXCEL_AGENT_MODEL=...");
	process.exit(1);
}
const providerId = cfg.provider || "custom-llm";

/* ---------------- 极简 MCP stdio 客户端（子进程 + 行 JSON-RPC） ---------------- */
class McpClient {
	constructor(name, command, args, env) {
		this.name = name;
		this.nextId = 1;
		this.pending = new Map();
		this.child = spawn(command, args, { stdio: ["pipe", "pipe", "inherit"], env: { ...process.env, ...(env || {}) }, cwd: root });
		this.buf = "";
		this.child.stdout.on("data", (d) => this._feed(d));
		this.child.on("exit", (code) => {
			if (code !== 0 && !this.disposed) console.error(`[mcp:${name}] 进程退出 code=${code}`);
			for (const { reject } of this.pending.values()) reject(new Error(`MCP ${name} 进程已退出`));
			this.pending.clear();
		});
	}
	_feed(data) {
		this.buf += data;
		let idx;
		while ((idx = this.buf.indexOf("\n")) >= 0) {
			const line = this.buf.slice(0, idx).trim();
			this.buf = this.buf.slice(idx + 1);
			if (!line) continue;
			try {
				const msg = JSON.parse(line);
				if (msg.id !== undefined && this.pending.has(msg.id)) {
					const { resolve, reject } = this.pending.get(msg.id);
					this.pending.delete(msg.id);
					msg.error ? reject(new Error(msg.error.message)) : resolve(msg.result);
				}
			} catch { /* 非 JSON 行忽略 */ }
		}
	}
	async request(method, params) {
		const id = this.nextId++;
		return new Promise((resolve, reject) => {
			this.pending.set(id, { resolve, reject });
			this.child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n");
			setTimeout(() => {
				if (this.pending.has(id)) {
					this.pending.delete(id);
					reject(new Error(`MCP ${this.name} ${method} 超时`));
				}
			}, 60_000).unref();
		});
	}
	async start() {
		await this.request("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "pi-excel-agent", version: "1.0.0" } });
		this.child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n");
		const { tools } = await this.request("tools/list", {});
		return tools;
	}
	async call(name, args) {
		const r = await this.request("tools/call", { name, arguments: args });
		const text = (r.content || []).map((c) => c.text || "").join("\n");
		if (r.isError) throw new Error(text || `工具 ${name} 报错`);
		return text;
	}
	dispose() {
		this.disposed = true;
		this.child.kill();
	}
}

const py = process.env.PYTHON || "python3";
const servers = [
	{ prefix: "excel", client: new McpClient("excel", py, [path.join(root, "excel-mcp/server.py")]) },
	{ prefix: "recon", client: new McpClient("table-recon", py, [path.join(root, "03-mcp-server/server.py")], { TABLE_RECON_HOME: root }) },
];

const toolNames = [];
for (const s of servers) {
	let tools = [];
	try {
		tools = await s.client.start();
	} catch (e) {
		console.error(`[mcp:${s.prefix}] 启动失败，跳过: ${e.message}`);
		continue;
	}
	s.tools = tools;
	for (const t of tools) toolNames.push(`${s.prefix}_${t.name}`);
	console.log(`▶ MCP ${s.prefix}: ${tools.length} 工具就绪 (${tools.map((t) => t.name).slice(0, 4).join(", ")}…)`);
}

/* ---------------- 生成 models.json 注册自定义 provider ---------------- */
const modelsJson = {
	providers: {
		[providerId]: {
			name: `Custom (${new URL(cfg.baseUrl).host})`,
			baseUrl: cfg.baseUrl,
			apiKey: cfg.apiKey, // 字面量或 "$ENV_VAR"（pi 自动展开）
			api: cfg.api || "openai-completions",
			models: [{
				id: cfg.model, name: cfg.model, reasoning: !!cfg.reasoning, input: ["text"],
				contextWindow: cfg.contextWindow || 131072, maxTokens: cfg.maxTokens || 8192,
				cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
			}],
		},
	},
};
const modelsPath = path.join(os.tmpdir(), `pi-excel-models-${process.pid}.json`);
fs.writeFileSync(modelsPath, JSON.stringify(modelsJson, null, 2));
const modelRuntime = await ModelRuntime.create({ modelsPath });
const model = modelRuntime.getModel(providerId, cfg.model);
if (!model) {
	console.error(`模型注册失败: ${providerId}/${cfg.model}`);
	process.exit(1);
}

/* ---------------- 把 MCP 工具注册进 pi（inline extension） ---------------- */
const mcpBridge = {
	name: "mcp-bridge",
	factory: (pi) => {
		for (const s of servers) {
			for (const t of s.tools || []) {
				const fullName = `${s.prefix}_${t.name}`;
				pi.registerTool({
					name: fullName,
					label: fullName,
					description: `[${s.prefix}] ${t.description}`,
					parameters: Type.Unsafe(t.inputSchema || { type: "object", properties: {} }),
					execute: async (_id, args) => {
						try {
							const text = await s.client.call(t.name, args);
							return { content: [{ type: "text", text }], details: {} };
						} catch (e) {
							return { content: [{ type: "text", text: `错误: ${e.message}` }], isError: true, details: {} };
						}
					},
				});
			}
		}
	},
};

const resourceLoader = new DefaultResourceLoader({ cwd: root, agentDir: getAgentDir(), extensionFactories: [mcpBridge] });
await resourceLoader.reload();

/* ---------------- 任务与会话 ---------------- */
const task =
	process.argv[2] ??
	[
		"请先读 .pi/skills/table-recon/SKILL.md，然后完成：",
		`1) 用 excel_inspect_workbook 查看 ${path.join(root, "测试_A方.xlsx")} 和 ${path.join(root, "测试_B方.xlsx")}；`,
		"2) 用 excel_query_rows 查 B 方实称重量大于 700 的单号；",
		"3) 把 A 方复制到 /tmp/演示_A方.xlsx，用 excel_update_rows 把 DD001 的商品改成「苹果(测试)」，用 excel_format_range 把表头加粗；",
		"4) 中文汇报每一步结果。",
	].join("\n");

console.log(`▶ 模型  ${providerId}/${cfg.model}`);
console.log(`▶ 端点  ${cfg.baseUrl}`);
console.log(`▶ cwd   ${root}`);
console.log(`▶ 任务  ${task}\n`);

const { session } = await createAgentSession({
	cwd: root,
	resourceLoader,
	model,
	modelRuntime,
	tools: ["read", "bash", ...toolNames],
});

try {
	session.subscribe((event) => {
		if (event.type === "message_update" && event.assistantMessageEvent.type === "text_delta") {
			process.stdout.write(event.assistantMessageEvent.delta);
		}
		if (event.type === "tool_execution_start") console.log(`\n[工具] ${event.toolName}`);
	});
	await session.prompt(task);
	console.log("\n\n✔ 完成");
} finally {
	session.dispose();
	for (const s of servers) s.client.dispose();
	fs.rmSync(modelsPath, { force: true });
}

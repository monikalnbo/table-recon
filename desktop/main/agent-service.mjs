/**
 * desktop/main/agent-service.mjs —— pi Agent 服务（Electron 主进程）
 * =============================================================================
 * 深度嵌入要点：
 *   - 工具 = 进程内直调 canonical 双挂载（excel-mcp-js / recon-js 的 createTools）
 *     不起 MCP 子进程、不依赖 pi-mcp-adapter（Spike2 实证：adapter 只在 pi CLI 的
 *     session_start 事件初始化，SDK 会话不发该事件 → 不可用）
 *   - UI 与 AI 同一套工具实现（renderer 的核对按钮走 recon-js 同一代码路径）
 * 会话生命周期：
 *   - 单会话常驻；配置签名变化（baseUrl/key/model/api/provider）→ dispose 重建
 *   - 事件回调存实例字段 _events（每轮 run 换新回调，会话复用不失效——B 类回归已覆盖）
 * 配置落盘：userData/agent-config.json（apiKey 明文 + 支持 $ENV 引用，由 pi 展开）
 */
import fs from "node:fs";
import path from "node:path";
import { createAgentSession, ModelRuntime } from "@earendil-works/pi-coding-agent"; // pi SDK（Electron ≥35 / Node22）
import { Type } from "@sinclair/typebox"; // 动态工具 schema 用 Type.Unsafe 包原始 JSON Schema
import { createTools as excelTools } from "excel-mcp-js/tools.mjs"; // 13 个 Excel 工具（canonical）
import { createTools as reconTools } from "recon-js"; // 核对工具（canonical）

/**
 * 挂载器：给 canonical 工具加前缀（excel_* / recon_*）并包装成 pi 工具形状。
 * parameters 用 Type.Unsafe 直包原始 JSON Schema（canonical 即契约，无需翻译）。
 * 错误转 isError 内容回传（AI 可自纠），不抛穿（防会话崩）。
 */
const mount = (prefix, tools) =>
	tools().map((t) => ({ // tools 是工厂（每次调用产新闭包，无状态）
		name: `${prefix}_${t.name}`, // excel_query_rows / recon_compare_tables …
		description: `[${prefix}] ${t.description}`, // 前缀帮模型分组
		parameters: Type.Unsafe(t.inputSchema || { type: "object", properties: {} }), // 原始 schema 直包
		execute: async (_id, args) => {
			try {
				return { content: [{ type: "text", text: await t.execute(args) }], details: {} }; // 成功：文本载荷
			} catch (e) {
				return { content: [{ type: "text", text: `错误: ${e.message}` }], isError: true, details: {} }; // 失败：isError 标记
			}
		},
	}));

// 桌面版全量工具（模块级常量：与 server.mjs 生态同形）
export const AGENT_TOOLS = [...mount("excel", excelTools), ...mount("recon", reconTools)];

export class AgentService {
	/** userDataDir：Electron app.getPath('userData')（portable 模式已重定向到 exe 旁） */
	constructor(userDataDir) {
		this.userDataDir = userDataDir; // 配置根目录
		this.configFile = path.join(userDataDir, "agent-config.json"); // 模型配置
		this.modelsFile = path.join(userDataDir, "models.json"); // pi 的自定义 provider 注册
		this.session = null; // 常驻会话
		this.busy = false; // 串行锁（同时只跑一个任务）
	}

	/** 读配置；无文件 → 全默认骨架（baseUrl/key/model 为空 = 未配置态） */
	getConfig() {
		try {
			return JSON.parse(fs.readFileSync(this.configFile, "utf8"));
		} catch {
			return { baseUrl: "", apiKey: "", model: "", api: "openai-completions", provider: "custom-llm", reasoning: false, contextWindow: 131072, maxTokens: 8192 };
		}
	}

	/** 合并写配置（浅合并：UI 只提交改动字段也安全） */
	setConfig(cfg) {
		const cur = this.getConfig(); // 现值
		const next = { ...cur, ...cfg }; // 浅合并
		fs.mkdirSync(this.userDataDir, { recursive: true }); // 首次写入建目录
		fs.writeFileSync(this.configFile, JSON.stringify(next, null, 2));
		return next;
	}

	/**
	 * 生成 pi 的 models.json（自定义 provider 注册）。
	 * apiKey 原样写入：字面量或 "$ENV_VAR"（pi 运行时展开 → key 可不落明文）。
	 * 返回 providerId。
	 */
	writeModels(cfg) {
		const pid = cfg.provider || "custom-llm"; // provider 标识
		const models = { // pi models.json 格式（见 pi docs/models.md）
			providers: {
				[pid]: {
					name: `Custom (${safeHost(cfg.baseUrl)})`, // 显示名（host 部分即可辨识）
					baseUrl: cfg.baseUrl, // OpenAI 兼容端点
					apiKey: cfg.apiKey, // 字面量或 $ENV 引用
					api: cfg.api || "openai-completions", // 协议（绝大多数用这个）
					models: [{ // 单模型注册（桌面场景够用）
						id: cfg.model, name: cfg.model,
						reasoning: !!cfg.reasoning, // 是否推理模型
						input: ["text"], // 输入模态
						contextWindow: cfg.contextWindow || 131072, // 上下文窗口
						maxTokens: cfg.maxTokens || 8192, // 单回复上限
						cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }, // 自带端点不计费
					}],
				},
			},
		};
		fs.writeFileSync(this.modelsFile, JSON.stringify(models, null, 2));
		return pid;
	}

	/**
	 * 启动（或按需重建）会话。
	 * events 先存实例字段（_events）——subscribe 闭包每轮读最新值（修复：旧版闭包
	 * 捕获首轮回调，第二轮起 UI 无输出的 bug）。
	 */
	async ensureSession(events) {
		this._events = events; // ① 先落回调（无论是否复用会话）
		const cfg = this.getConfig();
		if (!cfg.baseUrl || !cfg.apiKey || !cfg.model) throw new Error("模型未配置：请先在 AI 设置里填 baseUrl / apiKey / model"); // 未配置前置拦截
		if (this.session) { // ② 已有会话：签名相同直接复用
			if (this._sessionCfgSig === JSON.stringify([cfg.baseUrl, cfg.apiKey, cfg.model, cfg.api, cfg.provider])) return this.session;
			this.dispose(); // 配置变了 → 销毁重建
		}
		this._sessionCfgSig = JSON.stringify([cfg.baseUrl, cfg.apiKey, cfg.model, cfg.api, cfg.provider]); // 记新签名
		const pid = this.writeModels(cfg); // ③ 注册 provider
		const modelRuntime = await ModelRuntime.create({ modelsPath: this.modelsFile }); // ④ 运行时挂载
		const model = modelRuntime.getModel(pid, cfg.model); // 取模型对象
		if (!model) throw new Error(`模型注册失败: ${pid}/${cfg.model}（检查 baseUrl/model 是否匹配）`);

		// ⑤ 工具注册扩展（inline factory：pi SDK 官方路径，实测可用）
		const toolsExt = {
			name: "mcp-tools", // 扩展显示名
			factory: (pi) => { // pi: ExtensionAPI
				for (const t of AGENT_TOOLS) pi.registerTool(t); // 全量注册 15 工具
			},
		};
		// ⑥ 隔离资源加载：agentDir 指到 userData 子目录——不扫用户全局 ~/.pi
		//    （分发版不该吃宿主的 pi 配置；想复用登录态是后续选项）
		fs.mkdirSync(path.join(this.userDataDir, "pi-agent"), { recursive: true }); // 目录必须先存在
		const { DefaultResourceLoader, getAgentDir } = await import("@earendil-works/pi-coding-agent"); // 同包再取（保持单例解析）
		const loader = new DefaultResourceLoader({
			cwd: this.userDataDir, // 会话 cwd（报告等相对路径落点）
			agentDir: path.join(this.userDataDir, "pi-agent"), // 隔离的 agent 目录
			extensionFactories: [toolsExt], // 挂工具扩展
		});
		await loader.reload(); // 触发 factory 执行（注册工具）

		// ⑦ 建会话：白名单工具 = read/bash + 15 个业务工具（不引 write/edit/find 等）
		const { session } = await createAgentSession({
			cwd: this.userDataDir,
			resourceLoader: loader, // 上面的加载器
			model, // 自定义模型
			modelRuntime,
			tools: ["read", "bash", ...AGENT_TOOLS.map((t) => t.name)], // 工具白名单
		});
		// ⑧ 订阅事件：闭包读 this._events（每轮最新回调）→ 会话复用时事件不失效
		session.subscribe((e) => {
			const ev = this._events || {}; // 动态取当前回调
			if (e.type === "message_update" && e.assistantMessageEvent.type === "text_delta") ev.onDelta?.(e.assistantMessageEvent.delta); // 流式文本
			if (e.type === "tool_execution_start") ev.onToolStart?.(e.toolName); // 工具开始
			if (e.type === "tool_execution_end") { // 工具结束（截 300 字摘要防刷屏）
				const txt = (e.result?.content || []).map((c) => c.text || "").join(" ").slice(0, 300);
				ev.onToolEnd?.(e.toolName, !!e.isError, txt);
			}
		});
		this.session = session; // 常驻
		return session;
	}

	/**
	 * 跑一个任务（串行）。context = UI 注入的当前状态（已加载文件/列/规则）
	 * → buildContext 拼成前言，模型无需用户重复给路径。
	 */
	async run(task, context = {}, events = {}) {
		if (this.busy) throw new Error("已有任务在跑，请先停止"); // 串行锁（并发跑会串流）
		this.busy = true;
		try {
			const session = await this.ensureSession(events); // 就绪/复用会话
			const ctx = buildContext(context); // UI 状态 → 前言
			await session.prompt(ctx ? `${ctx}\n\n任务：${task}` : task); // 组装并执行
			events.onDone?.(); // 正常完成
		} catch (e) {
			events.onError?.(e.message); // 错误回传（不抛穿，main.mjs 还有兜底 catch）
		} finally {
			this.busy = false; // 解锁（必经路径）
		}
	}

	/** 中止当前任务（会话保留，下一任务可继续） */
	async stop() {
		try { await this.session?.abort(); } catch { /* ignore */ }
	}

	/** 销毁会话（配置变更/退出时） */
	dispose() {
		try { this.session?.dispose(); } catch { /* ignore */ }
		this.session = null;
		this._sessionCfgSig = null; // 清签名（下次必重建）
	}
}

/** UI 状态 → 任务前言：已加载文件、列名、关联列、规则 → 模型直接可用 */
function buildContext(ctx) {
	const parts = [];
	if (ctx.fileA) parts.push(`用户已在界面加载 A 方表格: ${ctx.fileA}（列: ${ctx.headersA?.join(" / ") || "未知"}，关联列: ${ctx.keyA || "未选"}）`); // A 方事实
	if (ctx.fileB) parts.push(`B 方表格: ${ctx.fileB}（列: ${ctx.headersB?.join(" / ") || "未知"}，关联列: ${ctx.keyB || "未选"}）`); // B 方事实
	if (ctx.rules?.length) parts.push(`界面已配置规则: ${JSON.stringify(ctx.rules)}`); // 规则事实
	if (parts.length) parts.push("可直接对这些文件用 excel_* / recon_* 工具操作，不必让用户重新提供路径。"); // 行动指引
	return parts.join("\n");
}

/** baseUrl → host（显示名用；非法 URL 兜底 "custom"） */
function safeHost(u) {
	try { return new URL(u).host; } catch { return "custom"; }
}

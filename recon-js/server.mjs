#!/usr/bin/env node
/**
 * recon-js/server.mjs —— stdio MCP server（薄壳）
 * =============================================================================
 * 职责：把 index.createTools() 暴露为按行 JSON-RPC 的 stdio MCP 服务
 *      （协议与 03-mcp-server/server.py 完全一致，pi-mcp-adapter 实测兼容）。
 * 注册示例（pi 的 mcp.json）：
 *   "table-recon": { "command": "node", "args": ["/path/to/recon-js/server.mjs"] }
 */
import { createTools } from "./index.mjs"; // 工具实现（与桌面进程内挂载同源）
import readline from "node:readline"; // 按行读 stdin

const TOOLS = createTools(); // 模块加载即构建（进程生命周期内不变）
const SERVER_INFO = { name: "table-recon", version: "1.1.0" }; // 版本随 canonical 走

/** 单帧输出：JSON + 换行（flush 由 write 保证） */
function send(msg) {
	process.stdout.write(JSON.stringify(msg) + "\n");
}

/** 请求分发（MCP 2024-11-05 最小实现） */
function handle(req) {
	const method = req.method, id = req.id; // 方法名 / 请求 id（通知无 id）
	if (method === "initialize") {
		// 握手：回报协议版本、能力集、服务信息
		send({ jsonrpc: "2.0", id, result: { protocolVersion: "2024-11-05", capabilities: { tools: {} }, serverInfo: SERVER_INFO } });
	} else if (method === "notifications/initialized") {
		/* 通知无响应（协议约定） */
	} else if (method === "ping") {
		send({ jsonrpc: "2.0", id, result: {} }); // 保活
	} else if (method === "tools/list") {
		// 工具清单（只回三要素，execute 不外泄）
		send({ jsonrpc: "2.0", id, result: { tools: TOOLS.map(({ name, description, inputSchema }) => ({ name, description, inputSchema })) } });
	} else if (method === "tools/call") {
		const tool = TOOLS.find((t) => t.name === req.params.name); // 按名查工具
		if (!tool) return send({ jsonrpc: "2.0", id, error: { code: -32602, message: `未知工具: ${req.params.name}` } }); // 协议级错误
		tool.execute(req.params.arguments || {}) // 业务执行（异步）
			.then((text) => send({ jsonrpc: "2.0", id, result: { content: [{ type: "text", text }], isError: false } })) // 成功载荷
			.catch((e) => send({ jsonrpc: "2.0", id, result: { content: [{ type: "text", text: `错误: ${e.message}` }], isError: true } })); // 业务错误回传给 AI（不崩进程）
	} else if (id !== undefined) {
		send({ jsonrpc: "2.0", id, error: { code: -32601, message: `未实现: ${method}` } }); // 其余有 id 请求 → 未实现
	}
}

// 主循环：逐行读 stdin → 解析 → 分发
const rl = readline.createInterface({ input: process.stdin });
rl.on("line", (line) => {
	line = line.trim(); // 容忍空白
	if (!line) return;
	try {
		handle(JSON.parse(line));
	} catch { /* 坏行丢弃（半事务性：坏行不影响后续） */ }
});

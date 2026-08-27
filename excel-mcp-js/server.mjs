#!/usr/bin/env node
/**
 * excel-mcp-js/server.mjs —— stdio MCP server（薄壳）
 * =============================================================================
 * 职责：把 tools.createTools() 暴露为按行 JSON-RPC 的 stdio MCP 服务
 *      （协议与 excel-mcp/server.py 完全一致）。
 * 注册示例（pi 的 mcp.json）：
 *   "excel": { "command": "node", "args": ["/path/to/excel-mcp-js/server.mjs"] }
 */
import { createTools } from "./tools.mjs"; // 13 工具实现（与桌面进程内挂载同源）
import readline from "node:readline"; // 按行读 stdin

const TOOLS = createTools(); // 进程生命周期内不变
const SERVER_INFO = { name: "excel-mcp", version: "1.0.0" }; // 与 Python 版同名同版本线

/** 单帧输出：JSON + 换行 */
function send(msg) {
	process.stdout.write(JSON.stringify(msg) + "\n");
}

/** 请求分发（MCP 2024-11-05 最小实现） */
function handle(req) {
	const method = req.method, id = req.id; // 方法 / 请求 id
	if (method === "initialize") {
		send({ jsonrpc: "2.0", id, result: { protocolVersion: "2024-11-05", capabilities: { tools: {} }, serverInfo: SERVER_INFO } }); // 握手
	} else if (method === "notifications/initialized") {
		/* 通知无响应 */
	} else if (method === "ping") {
		send({ jsonrpc: "2.0", id, result: {} }); // 保活
	} else if (method === "tools/list") {
		send({ jsonrpc: "2.0", id, result: { tools: TOOLS.map(({ name, description, inputSchema }) => ({ name, description, inputSchema })) } }); // 工具三要素清单
	} else if (method === "tools/call") {
		const tool = TOOLS.find((t) => t.name === req.params.name); // 按名查
		if (!tool) return send({ jsonrpc: "2.0", id, error: { code: -32602, message: `未知工具: ${req.params.name}` } });
		tool.execute(req.params.arguments || {})
			.then((text) => send({ jsonrpc: "2.0", id, result: { content: [{ type: "text", text }], isError: false } })) // 成功
			.catch((e) => send({ jsonrpc: "2.0", id, result: { content: [{ type: "text", text: `错误: ${e.message}` }], isError: true } })); // 业务错误不崩进程
	} else if (id !== undefined) {
		send({ jsonrpc: "2.0", id, error: { code: -32601, message: `未实现: ${method}` } }); // 未实现
	}
}

// 主循环：逐行 stdin → JSON → 分发；坏行丢弃
const rl = readline.createInterface({ input: process.stdin });
rl.on("line", (line) => {
	line = line.trim();
	if (!line) return;
	try {
		handle(JSON.parse(line));
	} catch { /* 坏行丢弃 */ }
});

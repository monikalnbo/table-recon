/**
 * AI 全链路冒烟（Electron 主进程内 AgentService → excel_* 工具 → 真实 API）
 * 用法: xvfb-run -a npx electron test-ai-live.mjs --no-sandbox
 */
import { app } from "electron";
import path from "node:path";
import { fileURLToPath } from "node:url";
import os from "node:os";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");

app.whenReady().then(async () => {
	try {
		const { AgentService } = await import("./main/agent-service.mjs");
		const ud = path.join(os.tmpdir(), `tr-ai-${process.pid}`);
		const svc = new AgentService(ud);
		svc.setConfig({
			baseUrl: "https://api.siliconflow.cn/v1",
			apiKey: "$SILICONFLOW_API_KEY",
			model: "Qwen/Qwen2.5-72B-Instruct",
			contextWindow: 32768, maxTokens: 4096,
		});
		let deltas = 0, tools = [];
		await svc.run(
			`用 excel_query_rows 查 ${path.join(root, "测试_B方.xlsx")} 中实称重量大于 700 的单号，中文简述结果。不要读文件内容到对话里，用工具。`,
			{},
			{
				onDelta: () => deltas++,
				onToolStart: (n) => { tools.push(n); console.log(`  [tool] ${n}`); },
				onToolEnd: (n, isErr, s) => isErr && console.error(`  [tool-err] ${n}: ${s}`),
				onDone: () => console.log("  [done]"),
				onError: (m) => { console.error("  [error]", m); process.exitCode = 1; },
			},
		);
		const excelUsed = tools.some((t) => t.startsWith("excel_"));
		console.log(excelUsed && deltas > 0 ? "✔ 第一轮：工具调用+流式" : `✖ 第一轮失败 tools=${tools} deltas=${deltas}`);
		if (!(excelUsed && deltas > 0)) process.exitCode = 1;

		/* 复用同一会话跑第二轮（验证事件重绑定，bug #2 回归） */
		let deltas2 = 0;
		await svc.run("不用任何工具，直接回复两个字：收到", {}, {
			onDelta: () => deltas2++,
			onError: (m) => { console.error("  [error-2]", m); process.exitCode = 1; },
		});
		console.log(deltas2 > 0 ? `✔ 第二轮：流式正常（${deltas2} deltas，会话复用+新回调生效）` : "✖ 第二轮无输出（事件重绑定回归）");
		if (deltas2 <= 0) process.exitCode = 1;
		svc.dispose();
		app.exit(process.exitCode ?? 0);
	} catch (err) {
		console.error("AI-SMOKE-FAIL:", err?.stack || err);
		app.exit(1);
	}
});

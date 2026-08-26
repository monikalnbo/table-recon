/**
 * 最简单的 pi Agent —— 双表核对版
 *
 * pi SDK 的 createAgentSession() 默认自带 read/write/edit/bash 四个工具，
 * AI 自己就能调用本仓库的 core/recon.py 完成双表核对。
 * 若已在 ~/.pi/agent/mcp.json 注册 table-recon MCP server，AI 也会自动用上 MCP 工具。
 *
 * 用法:
 *   npm install
 *   export ANTHROPIC_API_KEY=sk-ant-...   # 或其他已配置的 provider
 *   node recon-agent.mjs
 *   node recon-agent.mjs "用 A.xlsx 和 B.xlsx，按订单号核对，重量区间用range、数量用exact"
 */

import { createAgentSession } from "@earendil-works/pi-coding-agent";

const task =
	process.argv[2] ??
	"用 python 运行 ../core/recon.py 完成双表核对：" +
	"A方=../测试_A方.xlsx(关联列:订单号)，B方=../测试_B方.xlsx(关联列:单号)，" +
	"规则1: range 重量区间↔实称重量 容差0.05 A方裸数按kg B方裸数按g；" +
	"规则2: exact 数量↔数量。报告输出到 ../核对报告.xlsx，最后总结三类异常各有几条";

console.log(`▶ 任务: ${task}\n`);

const { session } = await createAgentSession();

try {
	session.subscribe((event) => {
		if (event.type === "message_update" && event.assistantMessageEvent.type === "text_delta") {
			process.stdout.write(event.assistantMessageEvent.delta);
		}
	});
	await session.prompt(task);
	console.log("\n\n✔ 完成");
} finally {
	session.dispose();
}

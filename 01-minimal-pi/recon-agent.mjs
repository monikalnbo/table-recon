/**
 * pi Agent · 双表核对
 *
 * 工作目录 = 仓库根。会加载：
 *   - .pi/skills/table-recon/SKILL.md
 *   - .mcp.json 里的 excel MCP + table-recon MCP
 *   （若本机已装 pi-mcp-adapter，SDK 会话会自动带上 MCP 工具）
 *
 * 用法:
 *   npm install
 *   node recon-agent.mjs
 *   node recon-agent.mjs "用 A.xlsx 和 B.xlsx 按订单号核对重量和数量"
 */

import { createAgentSession } from "@earendil-works/pi-coding-agent";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const task =
	process.argv[2] ??
	[
		"请先阅读 .pi/skills/table-recon/SKILL.md，再做双表核对。",
		`A方=${path.join(root, "测试_A方.xlsx")}`,
		`B方=${path.join(root, "测试_B方.xlsx")}`,
		"用 excel MCP 查看两表的 sheet、列名和前几行；",
		"然后用 table-recon 的 compare_tables：关联列 A=订单号 B=单号；",
		"规则1 range 重量区间↔实称重量 容差0.05 unit_a=kg unit_b=g；",
		"规则2 exact 数量↔数量。",
		`报告输出到 ${path.join(root, "核对报告.xlsx")}。`,
		"最后用中文总结三类异常。不要 bash 直接跑 recon.py，也不要用 Excel 公式自己比对。",
	].join("");

console.log(`▶ cwd: ${root}`);
console.log(`▶ 任务: ${task}\n`);

const { session } = await createAgentSession({ cwd: root });

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

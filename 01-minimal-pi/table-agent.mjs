/**
 * 最简单的 pi Agent —— 表格版
 *
 * 全部代码就这么多：createAgentSession() 默认自带
 * read / write / edit / bash 四个工具，加上你的模型配置，
 * 它就能自己动手读写 Excel（通过 bash 调 python/openpyxl）。
 *
 * 用法:
 *   npm install
 *   node table-agent.mjs "把 sales.csv 整理成 Excel 并按月汇总"
 *   node table-agent.mjs                        # 不带参数用默认任务
 *
 * 认证（二选一，同 pi CLI）:
 *   export ANTHROPIC_API_KEY=sk-ant-...   # 或其他 provider 的 key
 *   # 或已有订阅: 先在 pi CLI 里执行 /login
 */

import { createAgentSession } from "@earendil-works/pi-coding-agent";

const task =
	process.argv[2] ?? "用 python+openpyxl 在当前目录创建 demo.xlsx：一个季度销售表，含 月份/销售额 两列共 3 行数据，并在底部用公式求合计";

console.log(`▶ 任务: ${task}\n`);

const { session } = await createAgentSession();

try {
	// 流式打印 AI 回复
	session.subscribe((event) => {
		if (event.type === "message_update" && event.assistantMessageEvent.type === "text_delta") {
			process.stdout.write(event.assistantMessageEvent.delta);
		}
	});

	await session.prompt(task); // 发送任务，等待完成（含所有工具调用）
	console.log("\n\n✔ 完成");
} finally {
	session.dispose();
}

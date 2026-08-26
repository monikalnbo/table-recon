# 01 · 最简单的 pi Agent（双表核对版）

用 **~25 行代码**得到一个能自己完成双表核对的 AI Agent。

pi SDK 的 `createAgentSession()` 默认自带 `read` / `write` / `edit` / `bash` 四个工具，
AI 自己就会调用 `../core/recon.py` 做核对；如果注册了本仓库的 table-recon MCP server，
AI 还会自动用上 MCP 工具（两种方式都能跑，殊途同归）。

## 运行

```bash
pip install openpyxl
npm install                        # 安装 pi SDK
export ANTHROPIC_API_KEY=...      # 或已在 pi 里 /login 的订阅

node recon-agent.mjs              # 默认任务：核对仓库自带的测试表
node recon-agent.mjs "用 A.xlsx 和 B.xlsx 按订单号核对，重量range、金额exact，容差0.05"
```

## 核心就这么短

```js
import { createAgentSession } from "@earendil-works/pi-coding-agent";

const { session } = await createAgentSession();
session.subscribe((e) => {
  if (e.type === "message_update" && e.assistantMessageEvent.type === "text_delta")
    process.stdout.write(e.assistantMessageEvent.delta);
});
await session.prompt("用 ../core/recon.py 核对两份表，规则是……");
session.dispose();
```

自然语言即可描述规则（关联列、区间/精确、容差、单位），AI 会翻译成 recon.py 的参数。

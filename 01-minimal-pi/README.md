# 01 · pi Agent（双表核对）

用 pi SDK 起一个 Agent：读本仓库 skill，调 **excel MCP** 看表，调 **table-recon MCP** 核对。

```
自然语言
    ↓
pi Agent  +  skill table-recon
    ├─ excel MCP          读结构 / 预览 / 写格式
    └─ table-recon MCP    compare_tables → 标红报告
```

## 前置

1. 已登录 pi（或设置好模型 API key）
2. 已装 [pi-mcp-adapter](https://www.npmjs.com/package/pi-mcp-adapter)（`pi install npm:pi-mcp-adapter`）
3. 已装 `uv`（excel MCP 用 `uvx excel-mcp-server stdio`）
   ```bash
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```
4. `~/.pi/agent/mcp.json` 或仓库 `.mcp.json` 里同时注册 `excel` 与 `table-recon`（见根 README）

```bash
pip install openpyxl
npm install
node recon-agent.mjs
node recon-agent.mjs "用 A.xlsx 和 B.xlsx，按订单号核对重量区间和数量"
```

Agent 会先 `read` skill，再用 MCP。不要指望它去手写核对代码。

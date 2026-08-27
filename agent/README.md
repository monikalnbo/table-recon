# agent · pi Excel Agent（自定义 baseUrl + apiKey）

一个进程 = 完整的 pi Agent + 内置 excel MCP（`../excel-mcp/server.py`，查询/修改 Excel）。
模型端点完全自定义，任何 OpenAI 兼容服务都能接。

```
node excel-agent.mjs "你的任务"
        │
   pi SDK 会话（模型来自 config.json / 环境变量）
        │
   tools: read / bash / mcp
        ├─ excel MCP（本仓库 excel-mcp/server.py，13 个工具）
        └─ table-recon MCP（业务核对，可选）
```

## 配置模型

```bash
cd agent
cp config.example.json config.json
```

`config.json`：

```json
{
  "baseUrl": "https://api.siliconflow.cn/v1",
  "apiKey": "$SILICONFLOW_API_KEY",
  "model": "Qwen/Qwen2.5-7B-Instruct",
  "api": "openai-completions",
  "reasoning": false,
  "contextWindow": 128000,
  "maxTokens": 8192
}
```

- `baseUrl`：任意 OpenAI 兼容端点（SiliconFlow、DeepSeek、OpenRouter、vLLM、Ollama、中转站…）
- `apiKey`：直接写密钥，或 `"$环境变量名"` 引用（pi 自动展开，密钥不落盘）
- `api`：`openai-completions`（绝大多数）；Anthropic 兼容端点用 `anthropic-messages`
- `model`：端点上的模型 id

环境变量优先级更高：

```bash
export EXCEL_AGENT_BASE_URL=https://api.deepseek.com/v1
export EXCEL_AGENT_API_KEY=sk-xxx
export EXCEL_AGENT_MODEL=deepseek-chat
node excel-agent.mjs
```

## 运行

```bash
cd agent
npm install
node excel-agent.mjs                       # 默认演示任务（对仓库测试表查询+修改）
node excel-agent.mjs "把 /tmp/订单.xlsx 里金额>1000 的行标红"
```

前置：`pip install openpyxl`（excel MCP 唯一依赖）。**不需要 pi-mcp-adapter** —— agent 内置极简 MCP stdio 客户端，直接 spawn `excel-mcp/server.py` 与 `03-mcp-server/server.py`，把工具注册进 pi 会话（工具名 `excel_*` / `recon_*`）。

## Excel 工具一览（excel-mcp）

| 类 | 工具 | 说明 |
|---|---|---|
| 查 | `inspect_workbook` | sheet 列表 / 表头 / 预览 |
| 查 | `read_range` | 按 A1 区域读网格（可读公式原文） |
| 查 | `query_rows` | 条件筛选行，op 含 eq/ne/gt/lt/ge/le/contains/in/between/is_empty… |
| 改 | `write_data` / `append_rows` | 写区域 / 追加行（支持按表头字典） |
| 改 | `update_rows` | **查询式批量改**：筛选 + 改列值 |
| 改 | `set_cells` / `apply_formula` | 精确补丁 / 写公式 |
| 改 | `format_range` / `find_replace` / `delete_rows` | 格式化 / 替换 / 删行 |
| 结构 | `create_workbook` / `manage_sheets` | 新建簿 / create·copy·rename·delete sheet |

xlsx 与 csv 都支持（csv 无公式/格式，会明确报错）。自测：`python3 ../excel-mcp/server.py --selftest`。

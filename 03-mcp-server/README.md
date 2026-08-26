# 03 · table-recon MCP Server

把双表核对引擎暴露为 MCP 工具（stdio），供 pi / Claude / Cursor / Codex 等任意 MCP 客户端调用。
仅依赖 `openpyxl`，单文件 `server.py`。

## 工具

### inspect_sheet
读表格结构：列名、行数、前 5 行预览。核对前让 AI 先看列名选关联列。

```json
{ "path": "A.xlsx", "sheet": "Sheet1" }
```

### compare_tables
双表核对主工具：

```json
{
  "file_a": "测试_A方.xlsx",
  "file_b": "测试_B方.xlsx",
  "key_a": "订单号",
  "key_b": "单号",
  "rules": [
    { "name": "重量", "type": "range", "col_a": "重量区间", "col_b": "实称重量",
      "tolerance": 0.05, "unit_a": "kg", "unit_b": "g" },
    { "name": "数量", "type": "exact", "col_a": "数量", "col_b": "数量" }
  ],
  "output": "核对报告.xlsx"
}
```

返回摘要 + 三类异常明细 JSON（`range_mismatch` / `exact_mismatch` / `only_in_a` / `only_in_b`），
并生成标红 Excel 报告。

- `range`：A方=范围，B方=具体值，B值 ∉ A范围 = 不符（方向固定；单位自动换算）
- `exact`：两列完全一致
- `tolerance`：容差（基准单位 kg）；`unit_a/unit_b`：裸数字默认单位（kg/g/公斤/吨/lb）

## 注册

pi（`~/.pi/agent/mcp.json`）：

```json
"table-recon": {
  "command": "python3",
  "args": ["/abs/path/table-toolkit/03-mcp-server/server.py"],
  "env": { "TABLE_RECON_HOME": "/abs/path/table-toolkit" }
}
```

Claude Desktop / Cursor 等同理（stdio）。`TABLE_RECON_HOME` 指向仓库根目录；
不设则按 `server.py` 上一级自动定位 `core/recon.py`。

## 手动测试（无需客户端）

```bash
printf '%s\n' \
 '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"t","version":"0"}}}' \
 '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
 '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' \
 | python3 server.py
```

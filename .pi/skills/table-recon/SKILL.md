---
name: table-recon
description: 双表核对 + Excel 查询修改。用 pi + 内置 excel MCP（inspect/query/update/format）读写表格，用 table-recon MCP 做业务核对。当用户提到对账、核对 Excel、两表比对、重量区间、单号匹配、查表格、改表格、标红、写公式时使用。
---

# Excel 核对与编辑（pi + 内置 excel MCP + table-recon）

分工（不要混用）：

| 谁 | 干什么 | 不要干什么 |
|---|---|---|
| **excel MCP**（本仓库 `excel-mcp/server.py`，内置） | `inspect_workbook` 看结构；`query_rows` 条件查行；`update_rows` 查询式改值；`write_data`/`append_rows`/`set_cells` 写；`apply_formula` 公式；`format_range` 格式；`find_replace`/`delete_rows`/`manage_sheets` | 不要用公式拼「按键关联+区间包含」这种对账语义 |
| **table-recon MCP** | `compare_tables` 业务核对 + 标红报告 | 不要当通用 Excel 编辑器 |
| **core/recon.py** | 引擎本体；MCP 不可用时的 CLI 兜底 | 有 MCP 时不要 bash 跑 recon.py |

## 何时加载

- 两份 Excel 按某列（订单号/快递单号）对齐核对 → compare_tables
- 查表格（筛选、统计、找某几行）→ query_rows
- 改表格（批量改值、标红、加公式、追加行）→ update_rows / format_range / apply_formula

## Excel 工具速查

- 看结构：`inspect_workbook {path}`
- 条件查：`query_rows {path, filters:[{column, op, value}], columns?, limit?}`
  - op：`eq ne gt lt ge le contains not_contains in not_in between is_empty not_empty`；数值列自动按数值比
- 批量改：`update_rows {path, filters, updates:{列:新值}}`
- 写区域：`write_data {path, data:[[..]], start_cell}`；追加：`append_rows {path, rows:[{列:值}]}`
- 补丁：`set_cells {path, cells:{"B3":值}}`；公式：`apply_formula {path, cell, formula}`
- 标红：`format_range {path, start_cell, end_cell, bg_color:"FFC7CE", font_color:"9C0006", bold}`
- 换文字：`find_replace {path, find, replace, column?}`

## 核对流程（对账）

1. excel：`inspect_workbook` 两份表 → 确认 sheet 和列名
2. 选关联列与规则（列名对不上就问用户）
   - `range`：A 列是范围（`0.5-1kg`/`500-800g`/`小于1`），B 列是具体值（实称）
   - `exact`：数量、金额等必须一致
   - 实称重量若是裸数字（`800`），`unit_b` 用 `g`；区间裸数默认 `kg`
3. 只调一次 `table-recon_compare_tables`
4. 中文总结 summary + 三类异常条数；报告路径给用户。改格式/另存用 excel MCP。

## 禁止

- 不要用 apply_formula / 逐行比对代替 compare_tables
- 不要没看列名就猜 col_a/col_b
- A/B 方向反了会把具体值当 `[x,x]` 区间，产出假异常。区间必须在 A 方
- 区间必须在 A 方、点值在 B 方；实称裸数记得 unit_b=g

## 自检数据

仓库根 `测试_A方.xlsx` / `测试_B方.xlsx`。正确参数下应约：匹配 4、区间不符 2（DD002、DD008）、仅A DD004、仅B DD005。
excel MCP 自测：`python3 excel-mcp/server.py --selftest`。

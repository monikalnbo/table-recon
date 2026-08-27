---
name: table-recon
description: 双表核对（订单号/快递单号关联、重量区间 vs 实称、精确列比对、三类异常标红报告）。用 pi + excel MCP 读表，用 table-recon MCP 做业务核对。当用户提到对账、核对 Excel、两表比对、重量区间、单号匹配、核对报告时使用。
---

# 双表核对（pi + Excel MCP + table-recon）

分工（不要混用）：

| 谁 | 干什么 | 不要干什么 |
|---|---|---|
| **excel MCP**（haris-musa `excel-mcp-server`） | 读工作簿结构、预览单元格、写/格式化/公式/图表 | 不要自己用公式拼「按键关联 + 区间包含」 |
| **table-recon MCP** | `compare_tables` 业务核对 + 标红报告 | 不要当通用 Excel 编辑器 |
| **core/recon.py** | 引擎本体；MCP 不可用时的 CLI 兜底 | 有 MCP 时不要 `bash python recon.py` |

## 何时加载

用户要：两份 Excel 按某列（订单号/快递单号）对齐，核对重量区间、数量、金额等，并给出异常清单 / 标红报告。

## 标准流程

1. **先看表，再定规则**
   - excel：`get_workbook_metadata`（filepath）看 sheet 名
   - excel：`read_data_from_excel`（filepath, sheet_name, preview_only=true）看列名和前几行
   - 若 excel MCP 未连接，退回 `table-recon_inspect_sheet`
2. **选关联列与规则**（列名对不上就问用户）
   - 关联键：名称含 订单/单号/编号/id 的列，A/B 可以不同名
   - `range`：A 列是范围（`0.5-1kg` / `500-800g` / `小于1`），B 列是具体值（实称）
   - `exact`：数量、金额、SKU 等必须一致
   - 实称重量若是裸数字（`800`、`1200`），`unit_b` 用 `g`；区间裸数默认 `kg`
3. **只调一次核对**
   - `table-recon_compare_tables`：
     ```json
     {
       "file_a": "…/A.xlsx",
       "file_b": "…/B.xlsx",
       "key_a": "订单号",
       "key_b": "单号",
       "rules": [
         {"name": "重量", "type": "range", "col_a": "重量区间", "col_b": "实称重量",
          "tolerance": 0.05, "unit_a": "kg", "unit_b": "g"},
         {"name": "数量", "type": "exact", "col_a": "数量", "col_b": "数量"}
       ],
       "output": "…/核对报告.xlsx"
     }
     ```
4. **用中文总结** `summary` + 三类异常条数；把 `report` 路径给用户。需要改格式/另存时再用 excel MCP。

## 规则语义

- **range**：B 值必须落在 A 区间内（方向固定：A=范围，B=点值）。容差单位是 kg，`0.05` = ±50g。
- **exact**：两列完全一致（能解析成数就按数值比，否则文本归一化）。
- 单位：`g/kg/公斤/吨/lb` 自动换算；裸数字看 `unit_a` / `unit_b`。

## 禁止

- 不要用 excel MCP 的 `apply_formula` / 手工逐行比对代替 `compare_tables`
- 不要在 JS/Python 里再抄一套核对逻辑
- 不要在没看列名的情况下瞎猜 `col_a`/`col_b`
- A/B 方向反了会把具体值当成 `[x,x]` 区间，结果是假异常。区间必须在 A 方

## 自检数据

仓库根目录 `测试_A方.xlsx` / `测试_B方.xlsx`。正确参数下应约：匹配 4、区间不符 2（DD002、DD008）、仅A DD004、仅B DD005。

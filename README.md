# Table Recon · 双表核对工具箱

> 两份 Excel，指定一列（如订单号）做关联，自定义核对规则，一键找出所有差异。
> 同一套核心引擎，三种用法：**AI Agent（pi SDK）** / **图形界面（浏览器）** / **MCP 工具（任意 AI 客户端）**。

## 核对逻辑

```
A方表格 ──┐
          ├─ 按关联列（订单号）匹配 ──┬─ 两边都有 → 逐条执行核对规则
B方表格 ──┘                          ├─ 仅A方有 → 异常①
                                     └─ 仅B方有 → 异常②
规则不通过 → 异常③（区间不符 / 项不一致）
```

| 规则类型 | 语义 | 例 |
|---|---|---|
| **range 区间** | A方该列是**范围**，B方该列是**具体值**；B值不落在A范围内 = 不符 | A=`0.5-1kg`，B=`800g` → 0.8∈[0.5,1] ✓ |
| **exact 精确** | 两列必须完全一致（数值按值比，文本忽略大小写/空格） | A=`5`，B=`6` → 不一致 ✗ |

**单位自动区分与换算**：`800g`、`0.8kg`、`0.8公斤` 等价；区间两侧可各带单位（`500g-800g`、`1-2kg`）；
纯数字无单位时可按规则指定默认单位（如 B 方实称重量列裸数按 g、A 方区间列裸数按 kg）。
区间格式支持：`1-2` `0.5~1.5` `1至2` `小于3` `不超过2.5` `≥2` `3以上`，容差可调（如 ±0.05kg）。

**输出三类异常 + 标红报告**（`核对报告.xlsx`，7 个工作表：总览 / 区间不符 / 项不一致 / 仅A方 / 仅B方 / A方标注 / B方标注，异常行整行标红）。

## 目录

| 目录 | 是什么 | 怎么用 |
|---|---|---|
| `core/` | 核对引擎 recon.py（仅依赖 openpyxl） | CLI 直接跑 / 被下面三者复用 |
| `01-minimal-pi/` | 最简 pi Agent（~25 行 JS） | 自然语言下任务，AI 自己调引擎 |
| `02-gui/` | 单文件网页图形界面 | 浏览器打开，零安装，数据不出本机 |
| `03-mcp-server/` | MCP 工具（stdio，2 个 tools） | 接入 pi / Claude / Cursor 等 AI 客户端 |

## 快速开始

```bash
# 1) 命令行直接核对
python3 core/recon.py 测试_A方.xlsx 测试_B方.xlsx \
    --key-a 订单号 --key-b 单号 \
    --rule range:重量区间:实称重量:0.05:kg:g \
    --rule exact:数量:数量 \
    -o 核对报告.xlsx

# 2) 图形界面
open 02-gui/index.html        # 或双击，浏览器直接用

# 3) AI Agent
cd 01-minimal-pi && npm install && node recon-agent.mjs

# 4) MCP 工具（~/.pi/agent/mcp.json）
"table-recon": {
  "command": "python3",
  "args": ["/path/to/table-toolkit/03-mcp-server/server.py"],
  "env": { "TABLE_RECON_HOME": "/path/to/table-toolkit" }
}
```

各子目录有独立 README。`测试_A方.xlsx` / `测试_B方.xlsx` 是自带测试数据（内含区间不符、单边缺失等用例）。

## MCP 工具一览

| 工具 | 作用 |
|---|---|
| `inspect_sheet` | 读表头/行数/前5行预览，帮 AI 选列 |
| `compare_tables` | 双表核对：关联列 + 规则列表 → 异常清单 JSON + 标红报告 |

## 为什么自己写而不用现成的？

调研过 GitHub 主流 Excel MCP（haris-musa 4.1k★、negokaz 1k★、sbroenne 570★ 等），
它们强在通用 Excel 读写/格式化，但**都没有内置"双表按键关联 + 区间方向性核对 + 单位换算 + 三类异常"这种业务语义**，
交给 AI 逐步拼单元格既慢又易错。所以核心逻辑固化成本引擎，AI 只负责理解需求 → 填参数 → 解读结果。

## 依赖

- core / MCP：Python 3.10+，`pip install openpyxl`
- GUI：现代浏览器（SheetJS + ExcelJS 走 CDN）
- pi Agent：Node 20+，`npm install @earendil-works/pi-coding-agent`

# Table Recon · 双表核对工具箱

> 两份 Excel，指定一列（如订单号）做关联，自定义核对规则，一键找出所有差异。
>
> AI 路径：**pi Agent + excel MCP（读表）+ table-recon MCP（核对）**。  
> 离线路径：浏览器 GUI / 桌面 App / iOS（不经过 pi）。

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
| `core/` | 核对引擎 recon.py（仅依赖 openpyxl） | CLI 兜底；table-recon MCP 调它 |
| `excel-mcp/` | **内置 Excel MCP**（13 工具：查/改/公式/格式/sheet 管理） | 零依赖，`python3 excel-mcp/server.py --selftest` 自测 |
| `agent/` | **pi Agent**（自定义 baseUrl + apiKey + model） | `node excel-agent.mjs "任务"` |
| `.mcp.json` | pi 用的 MCP 注册 | `excel`（本仓库内置）+ `table-recon`（核对） |
| `.pi/skills/table-recon/` | pi skill | 先读表再核对/改表，禁止手写公式对账 |
| `01-minimal-pi/` | pi SDK 示例 Agent（用已登录 pi） | `node recon-agent.mjs` |
| `03-mcp-server/` | 业务核对 MCP（stdio） | `compare_tables` / `inspect_sheet` |
| `02-gui/` | 单文件网页图形界面 | 浏览器打开，离线，数据不出本机 |
| `desktop/` | 电脑版 App（Electron） | Actions 云打包 exe / Mac zip / AppImage |
| `ios/` | iPhone/iPad App | Actions 打包 ipa / Xcode |
| `TableRecon.swiftpm/` | Swift Playgrounds 免签版 | 无需 Mac |
| `scripts/` | `setup-pi.sh` | 检查环境 |

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

# 3) pi Agent（自定义 baseUrl/apiKey + 内置 excel MCP 查改 Excel）
pip install -r requirements.txt
cd agent && cp config.example.json config.json   # 填 baseUrl / apiKey / model
npm install && node excel-agent.mjs "查询并修改你的 Excel"

# 5) 桌面版 App（Windows exe / Mac zip / Linux AppImage）
# Actions → Build Desktop App → 下载对应 Artifact，详见 desktop/README.md
# Mac：解压 zip 得到「双表核对.app」，右键→打开（未签名）

# 6) iPhone App
# push 后 GitHub Actions 自动在 macOS 云机打包，Actions 页下载 TableRecon.ipa（未签名）
# 用 Sideloadly/AltStore 以自己 Apple ID 签名安装；或把 TableRecon.swiftpm 传入 Swift Playgrounds 免签直接运行
# 详见 ios/README.md 与 TableRecon.swiftpm/README.md
```

各子目录有独立 README。`测试_A方.xlsx` / `测试_B方.xlsx` 是自带测试数据（内含区间不符、单边缺失等用例）。

## MCP 分工

| Server | 工具 | 作用 |
|---|---|---|
| **excel**（内置 `excel-mcp/`，零依赖） | `inspect_workbook` `read_range` `query_rows` `write_data` `append_rows` `update_rows` `set_cells` `apply_formula` `format_range` `find_replace` `delete_rows` `create_workbook` `manage_sheets` | Excel/CSV 查询与修改：条件筛选、查询式批量改、公式、格式标红、sheet 管理 |
| **table-recon** | `compare_tables` / `inspect_sheet` | 按关联列 + range/exact 规则核对，产出三类异常和标红报告 |

excel MCP 管格子级读写；table-recon 管对账语义。pi skill（`.pi/skills/table-recon`）强制「先看表 → 再动手」。
想换 [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)（`uvx excel-mcp-server stdio`）也可以，但内置版无 uv 依赖且支持查询式修改。

## 为什么核对逻辑自己写？

Excel MCP（haris-musa 4.1k★ 等）强在通用读写/格式化，**没有**「按键关联 + 区间方向性核对 + 单位换算 + 三类异常」。
那部分固化在 `core/recon.py`，AI 只负责理解需求 → 填参数 → 解读结果。

## 依赖

- core / MCP / agent：Python 3.10+，`pip install -r requirements.txt`（openpyxl）
- pi Agent：Node 20+，`cd agent && npm install`；模型端点任意 OpenAI 兼容（baseUrl + apiKey 见 `agent/README.md`）；无需 pi-mcp-adapter
- GUI：现代浏览器（SheetJS + ExcelJS 走 CDN）

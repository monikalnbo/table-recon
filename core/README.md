# core · 双表核对引擎

`recon.py` —— 核对引擎。CLI 可直接跑；**table-recon MCP** 调它做 `compare_tables`。
pi Agent 的正常路径是 excel MCP 读表 + table-recon MCP 核对，不要再让 AI 用 bash 拼 `recon.py` 参数（MCP 挂了才用 CLI）。

```bash
python3 recon.py A.xlsx B.xlsx \
    --key-a 订单号 --key-b 单号 \
    --rule range:重量区间:实称重量:0.05:kg:g \   # 类型:A列:B列[:容差[:A裸数单位[:B裸数单位]]]
    --rule exact:数量:数量 \
    -o 核对报告.xlsx --json                     # --json 输出供程序消费
```

## API

```python
from recon import compare, export_report

result = compare("A.xlsx", "B.xlsx", key_a="订单号", key_b="单号", rules=[
    {"name": "重量", "type": "range", "col_a": "重量区间", "col_b": "实称重量",
     "tolerance": 0.05, "unit_a": "kg", "unit_b": "g"},
    {"name": "数量", "type": "exact", "col_a": "数量", "col_b": "数量"},
])
export_report(result, "核对报告.xlsx")   # 7 表标红报告
```

`compare()` 返回 dict：`summary`（计数）、`range_mismatch` / `exact_mismatch` /
`only_in_a` / `only_in_b`（明细，含原因）、`_raw`（导出用，勿序列化）。

## 关键函数

| 函数 | 作用 |
|---|---|
| `parse_value(text, default_unit)` | 具体值 → kg 基准 float，单位自动换算 |
| `parse_range(text, default_unit)` | 区间 → (low, high)，两侧可各带单位 |
| `check_range(va, vb, tol, ua, ub)` | A范围 ∋ B值？（方向性；容错自动互换） |
| `check_exact(va, vb)` | 数值比值 / 文本归一化比 |

单位支持：mg / g / 克 / kg / 公斤 / 千克 / 吨 / t / lb / 磅（NFKC 归一化，全角兼容）。

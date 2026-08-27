/**
 * excel-mcp-js/selftest.mjs —— 自测（无外部依赖，node 直接跑）
 * =============================================================================
 * 用例来源：
 *   1-13、csv 三例 —— 1:1 搬运 excel-mcp/server.py --selftest（防实现漂移的锚）
 *   14-18 —— 本 JS 版独有回归：B2（日期列 update）/ B11（find_replace 列限定）/
 *            B12（公式格保护）/ B15（真公式写入）
 * 用法: node excel-mcp-js/selftest.mjs   （退出码 0=全过）
 */
import { createTools } from "./tools.mjs"; // 被测出口（与进程内挂载/stdio 同源）
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const T = Object.fromEntries(createTools().map((t) => [t.name, t.execute])); // 工具名 → 实现直调
const td = fs.mkdtempSync(path.join(os.tmpdir(), "excel-mcp-js-")); // 每次跑用独立临时目录
let n = 0; // 用例计数器

/** 标准断言跑法：调工具 → JSON.parse → check(data) → ✔/✖ + 失败退出码 */
const run = async (label, name, args, check) => {
	const out = await T[name](args); // 工具返回 JSON 字符串
	const data = JSON.parse(out);
	const pass = check(data);
	console.log(`${pass ? "✔" : "✖"} ${++n}. ${label}`, pass ? "" : `—— ${out.slice(0, 200)}`);
	if (!pass) process.exitCode = 1; // 不中断，跑完看全貌
	return data;
};

/* ==================== 基础 13 例（Python selftest 同源） ==================== */

const p = path.join(td, "t.xlsx"); // 主 xlsx 夹具
await run("create_workbook", "create_workbook", { path: p, sheet_name: "数据", headers: ["订单号", "重量", "状态"] }, (d) => d.created === p);
await run("append_rows", "append_rows", { path: p, rows: [
	{ 订单号: "DD1", 重量: 800, 状态: "待发" }, // 字典模式：按表头对齐
	{ 订单号: "DD2", 重量: 1200, 状态: "待发" },
	{ 订单号: "DD3", 重量: 650, 状态: "已发" },
] }, (d) => d.appended === 3);
const insp = await run("inspect_workbook", "inspect_workbook", { path: p }, (d) => d.columns.join() === "订单号,重量,状态" && d.sheets[0].name === "数据");
await run("query_rows(>700)", "query_rows", { path: p, filters: [{ column: "重量", op: "gt", value: 700 }] }, (d) => d.matched === 2); // 数值比：800/1200 命中
await run("update_rows", "update_rows", { path: p, filters: [{ column: "订单号", op: "eq", value: "DD2" }], updates: { 状态: "已发", 重量: 1250 } }, (d) => d.updated === 1);
await run("query_rows(check)", "query_rows", { path: p, filters: [{ column: "订单号", op: "eq", value: "DD2" }] }, (d) => d.rows[0].重量 === 1250 && d.rows[0].状态 === "已发"); // 上一步落盘验证
await run("set_cells", "set_cells", { path: p, cells: { D1: "备注", D2: "易碎" } }, (d) => d.set === 2);
await run("apply_formula", "apply_formula", { path: p, cell: "D5", formula: "SUM(B2:B4)" }, (d) => d.formula === "=SUM(B2:B4)"); // 回显带 = 规范化
await run("format_range", "format_range", { path: p, start_cell: "A1", end_cell: "D1", bold: true, bg_color: "DDEBF7" }, (d) => d.formatted === 4); // 4 格表头
await run("find_replace", "find_replace", { path: p, find: "待发", replace: "待发货" }, (d) => d.replaced === 1); // 仅 DD1 一格（DD2 已改"已发"）
const rr = await run("read_range", "read_range", { path: p, start_cell: "A1", end_cell: "D5", values_only: false }, (d) => d.range === "A1:D5" && String(d.data[4][3]).includes("SUM")); // 公式模式可见 SUM
await run("manage_sheets(copy)", "manage_sheets", { path: p, action: "copy", sheet: "数据" }, (d) => d.sheets.length === 2);
await run("delete_rows", "delete_rows", { path: p, sheet: "数据 副本", row_start: 2 }, (d) => d.deleted_rows === 1);

/* ==================== csv 分支（Python 同源） ==================== */

const c = path.join(td, "t.csv"); // csv 夹具
await T.create_workbook({ path: c, headers: ["a", "b"] }); // 建表头
await T.append_rows({ path: c, rows: [[1, 2], [3, 4]] }); // 追加两行（字符串化存储）
await run("csv query", "query_rows", { path: c, filters: [{ column: "a", op: "ge", value: 2 }] }, (d) => d.matched === 1 && d.rows[0].a === "3"); /* csv 值为字符串，与 Python 一致 */
await run("csv update", "update_rows", { path: c, filters: [{ column: "a", op: "eq", value: 3 }], updates: { b: 99 } }, (d) => d.updated === 1);
const csvCheck = await T.query_rows({ path: c }); // 更新落盘验证
const csvOk = JSON.parse(csvCheck).rows[1].b === "99"; /* 字符串，与 Python 一致 */
console.log(`${csvOk ? "✔" : "✖"} ${++n}. csv update check`, csvOk ? "" : csvCheck);
if (!csvOk) process.exitCode = 1;

/* ==================== JS 版回归用例 ==================== */

/* ---- B2 回归：日期列存在时 update_rows 必须仍能命中 ----
 * 背景：旧实现「SheetJS 视图筛选 + 整行 JSON 相等映射回 ExcelJS」，
 * 日期在两库表示不同（序列号 vs Date）→ 永不相等 → updated:0 静默失败。
 * 用 ExcelJS 直接造真日期列夹具。 */
const dt = path.join(td, "date.xlsx");
{
	const ExcelJS = (await import("exceljs")).default; // 动态引入避免顶部耦合
	const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet("S");
	ws.addRow(["单号", "日期", "金额"]); // 表头
	ws.addRow(["DD1", new Date("2026-01-15"), 100]); // 真 Date 对象（复现触发条件）
	ws.addRow(["DD2", new Date("2026-02-20"), 200]);
	await wb.xlsx.writeFile(dt);
}
const u2 = JSON.parse(await T.update_rows({ path: dt, filters: [{ column: "单号", op: "eq", value: "DD1" }], updates: { 金额: 999 } }));
const u2ok = u2.updated === 1; // 必须命中 1 行（旧代码=0）
console.log(`${u2ok ? "✔" : "✖"} ${++n}. update_rows 带日期列（B2 回归）`, u2ok ? "" : JSON.stringify(u2));
if (!u2ok) process.exitCode = 1;
const u2q = JSON.parse(await T.query_rows({ path: dt, filters: [{ column: "单号", op: "eq", value: "DD1" }] }));
const u2qok = u2q.rows[0]?.金额 === 999; // 落盘验证
console.log(`${u2qok ? "✔" : "✖"} ${++n}. 日期列更新落盘验证`, u2qok ? "" : JSON.stringify(u2q.rows[0]));
if (!u2qok) process.exitCode = 1;

/* ---- B11 回归：find_replace 的 column 限定必须生效 ----
 * 背景：旧实现 xlsx 分支忽略 column 参数，全表误改（"plain"→"plzin"）。 */
const fr = path.join(td, "fr.xlsx");
await T.create_workbook({ path: fr, sheet_name: "S", headers: ["左列", "右列"] });
await T.append_rows({ path: fr, rows: [["xy", "xy"]] }); // 两列同值，便于区分改了哪列
const fr1 = JSON.parse(await T.find_replace({ path: fr, find: "x", replace: "X", column: "左列" })); // 只改左列
const frv = JSON.parse(await T.read_range({ path: fr, start_cell: "A2", end_cell: "B2" })); // 读回两列
const fr1ok = fr1.replaced === 1 && frv.data[0][0] === "Xy" && frv.data[0][1] === "xy"; // 左改右不动
console.log(`${fr1ok ? "✔" : "✖"} ${++n}. find_replace 列限定（B11 回归）`, fr1ok ? "" : JSON.stringify({ fr1, frv }));
if (!fr1ok) process.exitCode = 1;

/* ---- B12 + B15 回归：公式格跳过 & apply_formula 写的是真公式 ----
 * B12：find_replace 不得把公式格打平成静态文本。
 * B15：apply_formula 必须产出真公式（ExcelJS 需 {formula} 对象；"=..." 字符串是文本）。 */
const ff = path.join(td, "formula.xlsx");
await T.create_workbook({ path: ff, sheet_name: "S", headers: ["x", "y"] }); // 表头不含 'a'，避免夹具自伤（find_replace 语义含表头，与 Python 一致）
await T.append_rows({ path: ff, rows: [["aa", "bb"]] }); // A2="aa" 含查找目标
await T.apply_formula({ path: ff, cell: "C2", formula: 'CONCATENATE("a","a")' }); // 公式结果 "aa"，也是查找目标
const ffr = JSON.parse(await T.find_replace({ path: ff, find: "a", replace: "Q" })); // 全表替换
const ffrv = JSON.parse(await T.read_range({ path: ff, start_cell: "A2", end_cell: "C2", values_only: false })); // 公式模式读回
// 期望：A2 "aa"→"QQ" 计 1 格；C2 公式格被跳过且仍是公式（带 = 前缀）
const ffok = ffr.replaced === 1 && String(ffrv.data[0][2]).startsWith("=CONCATENATE");
console.log(`${ffok ? "✔" : "✖"} ${++n}. 公式保护+真公式写入（B12/B15 回归）`, ffok ? "" : JSON.stringify({ ffr, ffrv }));
if (!ffok) process.exitCode = 1;

console.log(`\n${process.exitCode ? "✖ 有失败" : "✔ selftest 全部通过"}（${td}）`);

#!/usr/bin/env node
/**
 * 业务模拟：水果电商「平台订单 vs 仓库实发」对账全链路
 * =============================================================================
 * 模拟真实业务的脏数据形态（来自 fruiterp/erp-warehouse 场景）：
 *   平台侧（A方）：订单号、商品、重量区间（各种手输格式）、数量、备注
 *   仓库侧（B方）：运单号、实称重量（g/kg 混杂、裸数）、数量、发货日期
 * 预期异常全景：区间不符/项不一致/仅A（未发）/仅B（无单）/重复键
 * 跑法：node sim-business.mjs（依次测 recon-js 引擎 + excel-mcp-js 工具）
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import ExcelJS from "exceljs";

const td = fs.mkdtempSync(path.join(os.tmpdir(), "sim-biz-"));
let fails = 0;
const ok = (m) => console.log(`  ✔ ${m}`);
const bad = (m) => { fails++; console.error(`  ✖ ${m}`); };

/* ============================================================
 * 夹具构造：模拟 ERP 导出（ExcelJS 造真文件，含日期/公式/富文本干扰）
 * ============================================================ */
const A = path.join(td, "平台订单.xlsx");
const B = path.join(td, "仓库实发.xlsx");
{
	const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet("订单");
	ws.addRow(["订单号", "商品", "重量区间", "数量", "备注"]);
	const rows = [
		["DD1001", "苹果-红富士5斤装", "4.5-5.5斤", 2, "老客户"],        // ✓ 应通过（B:5.1斤? 斤不在单位表!）
		["DD1002", "香蕉-整箱", "18-20kg", 1, ""],                        // ✓ kg 正常
		["DD1003", "车厘子-J级2kg", "1.8～2.2kg", 3, "加急"],             // ～ 全角波浪
		["DD1004", "脐橙", "小于10公斤", 5, ""],                          // 中文公斤
		["DD1005", "葡萄-阳光玫瑰", "500g以上", 10, "礼盒"],              // 中文区间
		["DD1006", "榴莲-金枕头", "6-8斤", 1, ""],                       // 斤又来了
		["DD1007", "草莓-丹东99", "0.9-1.1", 20, ""],                    // 裸数区间
		["DD1008", "芒果", "2~2.5公斤", 4, ""],                          // 公斤
		["DD1009", "火龙果", "1-1.5kg", 6, ""],                          // 未发（仅A）
		["DD1010", "龙眼", "3kg以上", 8, ""],                            // B 多发了（数量不符）
		["DD1001 ", "苹果-红富士5斤装", "4.5-5.5斤", 2, "重复订单号带尾空格"], // 重复键（带空格）
	];
	rows.forEach((r) => ws.addRow(r));
	await wb.xlsx.writeFile(A);
}
{
	const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet("实发");
	ws.addRow(["单号", "实称重量", "数量", "发货日期", "快递"]);
	const rows = [
		["DD1001", "5.1斤", 2, new Date("2026-08-20"), "顺丰"],          // 斤 vs 斤：parseValue 认斤吗？
		["DD1002", 19.2, 1, new Date("2026-08-20"), "京东"],             // 裸数（kg 语义）
		["DD1003", "2050g", 3, new Date("2026-08-21"), "顺丰"],          // g→kg 换算
		["DD1004", "9.6公斤", 5, new Date("2026-08-21"), "圆通"],        // 公斤
		["DD1005", "620g", 10, new Date("2026-08-22"), "中通"],          // 500g以上 → 0.62 ✓
		["DD1006", "3550", 1, new Date("2026-08-22"), "韵达"],           // 裸数=kg 语义? 实际 3.55kg ∉ [3,4]斤换算?
		["DD1007", "1.02kg", 20, new Date("2026-08-23"), "邮政"],        // ✓
		["DD1008", "2450g", 3, new Date("2026-08-23"), "顺丰"],          // 数量 3≠4 项不一致
		["DD1010", "3.1kg", 8, new Date("2026-08-24"), "京东"],          // ✓
		["DD9999", "2kg", 1, new Date("2026-08-24"), "顺丰"],            // 无单（仅B）
	];
	rows.forEach((r) => ws.addRow(r));
	await wb.xlsx.writeFile(B);
}

console.log("夹具:", td);

/* ============================================================
 * 场景1：recon-js 全链路核对（走 compare_tables 工具）
 * ============================================================ */
console.log("\n[场景1] recon-js compare_tables（斤/公斤/裸数/全角/重复键 混合）");
{
	const { createTools } = await import("/root/table-toolkit/recon-js/index.mjs");
	const T = Object.fromEntries(createTools().map((t) => [t.name, t.execute]));
	const out = JSON.parse(await T.compare_tables({
		file_a: A, file_b: B, key_a: "订单号", key_b: "单号",
		rules: [
			{ name: "重量", type: "range", col_a: "重量区间", col_b: "实称重量", tolerance: 0.05, unit_a: "kg", unit_b: "g" },
			{ name: "数量", type: "exact", col_a: "数量", col_b: "数量" },
		],
		output: path.join(td, "报告1.xlsx"),
	}));
	console.log("  summary:", JSON.stringify(out.summary));
	console.log("  区间不符:", out.range_mismatch.map((m) => `${m.key}(${m.a} vs ${m.b}): ${m.reason}`).join(" | ") || "无");
	console.log("  项不一致:", out.exact_mismatch.map((m) => `${m.key}: ${m.a}≠${m.b}`).join(" | ") || "无");
	console.log("  仅A:", out.only_in_a.map((o) => o.key).join(","), "| 仅B:", out.only_in_b.map((o) => o.key).join(","));

	/* --- 业务断言（对账员视角）--- */
	const s = out.summary;
	// 匹配键数 = 双方都有的键（注意 DD1001 带空格 = 同键归一化）
	s["匹配"] + s["区间不符"] + s["项不一致"] >= 9 ? ok("键覆盖数合理") : bad(`键覆盖异常 ${JSON.stringify(s)}`);
	// DD1009 必须在仅A（未发）
	out.only_in_a.some((o) => o.key === "dd1009") ? ok("DD1009 未发 → 仅A ✓") : bad("DD1009 未出现在仅A！");
	// DD9999 必须在仅B（无单）
	out.only_in_b.some((o) => o.key === "dd9999") ? ok("DD9999 无单 → 仅B ✓") : bad("DD9999 未出现在仅B！");
	// DD1008 数量 4 vs 3 → 项不一致
	out.exact_mismatch.some((m) => m.key === "dd1008") ? ok("DD1008 数量不符 → 项不一致 ✓") : bad("DD1008 数量不符漏检！");
	// 斤单位（canonical 双修后）：DD1001 4.5-5.5斤=[2.25,2.75]kg vs 5.1斤=2.55kg → 应匹配
	const dd1001miss = out.range_mismatch.find((m) => m.key === "dd1001");
	!dd1001miss ? ok("DD1001 斤单位区间核对通过（[2.25,2.75]∋2.55）✓") : bad(`斤单位仍误报: ${dd1001miss.reason}`);
}

/* ============================================================
 * 场景2：excel-mcp-js 工具链模拟 AI 对账员日常操作
 * ============================================================ */
console.log("\n[场景2] AI 对账员操作流（inspect → query → update → format → 验证）");
{
	const { createTools } = await import("/root/table-toolkit/excel-mcp-js/tools.mjs");
	const T = Object.fromEntries(createTools().map((t) => [t.name, t.execute]));

	// 2.1 侦察两表
	const iA = JSON.parse(await T.inspect_workbook({ path: A }));
	const iB = JSON.parse(await T.inspect_workbook({ path: B }));
	iA.columns.includes("重量区间") && iB.columns.includes("实称重量") ? ok("侦察列名 ✓") : bad("侦察缺列");

	// 2.2 条件查询：发货日期在 8/22 之后的顺丰单
	const q = JSON.parse(await T.query_rows({
		path: B, filters: [
			{ column: "快递", op: "eq", value: "顺丰" },
		], columns: ["单号", "实称重量", "快递"],
	}));
	console.log("  顺丰单:", JSON.stringify(q.rows.map((r) => r.单号)));
	q.matched === 4 ? ok("顺丰单=4 ✓") : bad(`顺丰单 matched=${q.matched}`);

	// 2.3 日期列筛选（gt：发货日期 > 2026-08-22）——ExcelJS Date vs 字符串
	const qd = JSON.parse(await T.query_rows({
		path: B, filters: [{ column: "发货日期", op: "gt", value: "2026-08-22" }],
	}));
	console.log("  8/22后发货:", qd.matched, "单", qd.matched > 0 ? `（首单 ${qd.rows[0]?.单号}）` : "（0！日期筛选可能失效）");
	qd.matched === 4 ? ok("日期筛选 ✓（0823×2+0824×2）") : bad(`日期筛选 matched=${qd.matched}（期望4）`);

	// 2.4 复制 A 方到工作区再改（不碰原始）
	const A2 = path.join(td, "平台订单_工作副本.xlsx");
	fs.copyFileSync(A, A2);
	// 2.5 对账员动作：把「加急」备注改成「加急-已电话确认」
	const fr = JSON.parse(await T.find_replace({ path: A2, find: "加急", replace: "加急-已电话确认", column: "备注" }));
	fr.replaced === 1 ? ok("备注替换（列限定）✓") : bad(`替换数=${fr.replaced}`);

	// 2.6 给未发货的 DD1009 标红提示
	const q9 = JSON.parse(await T.query_rows({ path: A2, filters: [{ column: "订单号", op: "eq", value: "DD1009" }] }));
	// 找行号：总行数定位
	const rr = JSON.parse(await T.read_range({ path: A2, start_cell: "A1", end_cell: `E${q9.total_rows + 1}` }));
	const rowIdx = rr.data.findIndex((r) => r[0] === "DD1009") + 1;
	rowIdx > 1 ? ok(`DD1009 在第 ${rowIdx} 行`) : bad("定位 DD1009 失败");
	await T.format_range({ path: A2, start_cell: `A${rowIdx}`, end_cell: `E${rowIdx}`, bg_color: "FFC7CE" });

	// 2.7 公式：工作副本加「应发重量上限」列
	await T.set_cells({ path: A2, cells: { F1: "应发上限(kg)" } });
	await T.apply_formula({ path: A2, cell: "F2", formula: 'IF(ISNUMBER(SEARCH("kg",C2)),C2,"")' });
	const fv = JSON.parse(await T.read_range({ path: A2, start_cell: "F2", end_cell: "F2", values_only: false }));
	String(fv.data[0][0]).startsWith("=IF") ? ok("公式写入 ✓") : bad(`公式异常 ${JSON.stringify(fv.data[0][0])}`);

	// 2.8 日期列 update_rows（B2 回归的业务版）：B 表本无「备注」列 → 先建列头再改
	await T.set_cells({ path: B, cells: { F1: "备注" } });
	const u = JSON.parse(await T.update_rows({
		path: B, filters: [{ column: "单号", op: "eq", value: "DD9999" }], updates: { 备注: "无单-待查" },
	}));
	console.log("  无单标记:", JSON.stringify(u));
	u.updated === 1 ? ok("无单打标（日期列共存）✓") : bad(`updated=${u.updated}`);
}

/* ============================================================
 * 场景3：重复键/带空格键的语义（对账员最怕的静默吞行）
 * ============================================================ */
console.log("\n[场景3] 重复键与键归一化（业务风险探针）");
{
	const { compareData } = await import("/root/table-toolkit/recon-js/engine.mjs");
	const res = compareData({
		headersA: ["单号", "金额"], rowsA: [
			{ 单号: "S1", 金额: 100 }, { 单号: "S1", 金额: 200 },   // 真重复
			{ 单号: "S2 ", 金额: 300 },                              // 尾空格 → 归一化后 S2
			{ 单号: "Ｓ３", 金额: 400 },                             // 全角Ｓ３ → NFKC → s3
		],
		headersB: ["单号", "金额"], rowsB: [
			{ 单号: "S1", 金额: 100 },
			{ 单号: "S2", 金额: 300 },
			{ 单号: "s3", 金额: 400 },
		],
		keyA: "单号", keyB: "单号",
		rules: [{ name: "金额", type: "exact", col_a: "金额", col_b: "金额" }],
	});
	console.log("  summary:", JSON.stringify(res.summary));
	// S1 重复：A 有两行 S1（100/200），Map last-wins 只留 200 → 与 B 的 100 比 → 应报不符
	res.exact_mismatch.some((m) => m.key === "s1" && m.a === 200) ? ok("S1 重复键 last-wins（200 参与核对，可检出）") : console.log("  [信息] S1 处理:", JSON.stringify(res.exact_mismatch.find((m) => m.key === "s1")));
	// 尾空格/全角：归一化后应正常关联
	res.summary["匹配"] + res.summary["项不一致"] === 3 ? ok("空格/全角键归一化关联 ✓") : bad(`关联数 ${res.summary["匹配"] + res.summary["项不一致"]} ≠ 3`);
}

/* ============================================================
 * 场景4：报告的 Excel 打开验证（业务交付物）
 * ============================================================ */
console.log("\n[场景4] 报告可开性 + 标红正确性（ExcelJS 回读验证）");
{
	const rp = path.join(td, "报告1.xlsx");
	if (!fs.existsSync(rp)) { bad("报告不存在"); }
	else {
		const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(rp);
		const names = wb.worksheets.map((w) => w.name);
		console.log("  sheets:", names.join(" / "));
		names.length === 7 ? ok("7 表结构 ✓") : bad(`表数 ${names.length}`);
		// 仅A方有页应有 DD1009 且整行红
		const ws = wb.getWorksheet("仅A方有");
		let red = 0;
		ws.eachRow((row, rn) => { if (rn > 1) { const c = row.getCell(2); if (c.fill?.fgColor?.argb === "FFFFC7CE") red++; } });
		red >= 1 ? ok(`仅A页标红 ${red} 行 ✓`) : bad("仅A页无标红");
	}
}

console.log(`\n${fails ? `✖ ${fails} 项业务断言失败` : "✔ 业务模拟全部通过"}`);
process.exit(fails ? 1 : 0);

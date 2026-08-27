/**
 * recon-js/report.mjs —— 7 表标红报告导出（ExcelJS），对齐 core/recon.py export_report。
 * 表结构：总览 / 区间不符 / 项不一致 / 仅A方有 / 仅B方有 / A方标注 / B方标注
 * 标红语义：异常清单全红；「标注」页 = 原表整行副本，命中行整行标红。
 */
import ExcelJS from "exceljs"; // 写侧（样式）
import { keyof } from "./engine.mjs"; // 复用引擎归一化（NFKC/trim/小写/去空格/全角符号替换，防标注页键错位）

// 标红样式：粉底 + 深红字（Excel 经典「坏」配色，与 Python FFC7CE/9C0006 一致）
const RED = { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } }, font: { color: { argb: "FF9C0006" } } };
const HEAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } }; // 表头底色（浅蓝）

/**
 * result = compareData() 返回值（含 _raw）；meta = {fileA, fileB}（总览页显示）。
 * 返回 outPath（调用方回填到结果 JSON 的 report 字段）。
 */
export async function exportReport(result, outPath, meta = {}) {
	const raw = result._raw; // 导出原料（headers/rows/maps）
	const wb = new ExcelJS.Workbook(); // 新簿（不复用读入簿，隔离污染）

	/** 建表页：表头行 + 加粗 + 底色 + 冻结首行 */
	const mk = (title, heads) => {
		const ws = wb.addWorksheet(title); // 中文名即 sheet 名
		ws.addRow(heads); // 表头
		ws.getRow(1).font = { bold: true }; // 加粗
		ws.getRow(1).fill = HEAD_FILL; // 底色
		ws.views = [{ state: "frozen", ySplit: 1 }]; // 冻结首行
		return ws;
	};
	/** 追加数据行；bad=true 整行标红 */
	const put = (ws, vals, bad) => {
		const row = ws.addRow(vals);
		if (bad) row.eachCell((c) => { c.fill = RED.fill; c.font = RED.font; }); // 逐格上样式
	};

	// ---- 页1：总览（文件、关联列、七项计数、规则回显）----
	mk("总览", ["项目", "数量"]).addRows([
		["A方文件", meta.fileA ?? ""], ["B方文件", meta.fileB ?? ""], // 文件路径（缺省空串）
		["关联列", `A:${result.keys.a} ↔ B:${result.keys.b}`], // 列名对照
		...Object.entries(result.summary), // 七项计数直接摊平
		["核对规则", result.rules.map((r) => `${r.name ?? r.type}(${r.type}: ${r.col_a}↔${r.col_b})`).join("; ")], // 规则一行摘要
	]);

	// ---- 页2/3：两类不符明细（全行红）----
	for (const [title, arr] of [["区间不符", result.range_mismatch], ["项不一致", result.exact_mismatch]]) {
		const ws = mk(title, ["关联键", "规则", "A方值", "B方值", "说明"]);
		arr.forEach((m) => put(ws, [m.key, m.rule, m.a, m.b, m.reason], true)); // bad=true
	}
	// ---- 页4/5：单边缺失（关联键 + 整行原始数据，全红）----
	for (const [title, arr, heads] of [["仅A方有", result.only_in_a, raw.headersA], ["仅B方有", result.only_in_b, raw.headersB]]) {
		const ws = mk(title, ["关联键", ...heads]); // 键列 + 原表全列
		arr.forEach((o) => put(ws, [o.key, ...heads.map((h) => o.row[h] ?? null)], true));
	}

	// ---- 页6/7：A/B 方标注（原表副本 + 命中行整行红）----
	// 命中集合 = 两类不符键 ∪ 本侧单边键
	const badA = new Set([...result.range_mismatch, ...result.exact_mismatch].map((m) => m.key).concat(result.only_in_a.map((o) => o.key)));
	const badB = new Set([...result.range_mismatch, ...result.exact_mismatch].map((m) => m.key).concat(result.only_in_b.map((o) => o.key)));
	const keyOfRaw = (row, key) => keyof(row, key); // 包装：复用引擎归一化（含 ～/— 替换）
	for (const [title, heads, rows, key, bad] of [ // A/B 两侧同构处理
		["A方标注", raw.headersA, raw.rowsA, raw.keyA, badA],
		["B方标注", raw.headersB, raw.rowsB, raw.keyB, badB],
	]) {
		const ws = mk(title, heads); // 与原表同列
		rows.forEach((row) => put(ws, heads.map((h) => row[h] ?? null), bad.has(keyOfRaw(row, key)))); // 命中才红
	}

	await wb.xlsx.writeFile(outPath); // 落盘
	return outPath;
}

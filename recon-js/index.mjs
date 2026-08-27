/**
 * recon-js/index.mjs —— canonical 入口（双挂载核心）
 * =============================================================================
 * 出口 1：readSheet()          文件 → {sheet, headers, rows}（读侧统一入口）
 * 出口 2：createTools()        MCP 形态工具（inspect_sheet / compare_tables）
 *                              供 server.mjs（stdio MCP）与桌面进程内挂载共用
 * 读：SheetJS（xlsx/csv）；核对：engine.compareData；报告：report.exportReport。
 */
import * as XLSX from "xlsx"; // SheetJS：只读
import { compareData } from "./engine.mjs"; // 纯函数核对引擎
import { exportReport } from "./report.mjs"; // 7 表标红报告

export { compareData, exportReport } from "./exports.mjs"; // 再出口便捷入口（供 import "recon-js" 一站式）

/** 1 → "A", 27 → "AA"（空表头列兜底命名用；与 excel-mcp-js 同实现） */
const colLetter = (n) => {
	let s = "";
	while (n > 0) {
		const m = (n - 1) % 26; // 本位
		s = String.fromCharCode(65 + m) + s; // 前插
		n = (n - 1 - m) / 26; // 进位
	}
	return s;
};

/**
 * 读 xlsx/csv → { sheet, headers, rows }；首行为表头。
 * ⚠️ F7 修复：曾用 dict 模式（sheet_to_json 默认）取表头——
 *    「只有表头无数据」的文件 rows=[] → headers=[] → 上游报「关联列不存在」。
 *    现改网格模式（header:1），与 excel-mcp-js/sheetRows 同一策略：
 *    表头永远取首行网格；空表头列兜底「列X」；全空行跳过（对齐 Python read_sheet）。
 */
export function readSheet(file, sheet) {
	const wb = XLSX.read(file, { type: "buffer" }); // Buffer → 工作簿
	if (sheet && !wb.SheetNames.includes(sheet))
		throw new Error(`工作表不存在: '${sheet}'，可选: ${wb.SheetNames.join(" / ")}`); // 指定 sheet 必须存在（对齐 Python：报错不静默回退）
	const name = sheet || wb.SheetNames[0]; // 目标 sheet：指定或第一个
	const ws = wb.Sheets[name];
	// 网格模式读入（raw:true 原始值；defval:null 补空；blankrows:false 去全空行）
	const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false });
	// 表头：首行网格，空单元格兜底命名（防 AI/用户拿空列名操作）
	const headers = (grid[0] || []).map((h, i) => (h != null && String(h).trim() !== "" ? String(h).trim() : `列${colLetter(i + 1)}`));
	// 数据：网格转字典行（按表头对齐，缺省 null；跳过全空行）
	const rows = grid.slice(1)
		.filter((r) => r.some((c) => c != null && String(c).trim() !== ""))
		.map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? null])));
	return { sheet: name, headers, rows };
}

/**
 * 工具工厂：产出 MCP 形态工具数组（name/description/inputSchema/execute）。
 * readSheetImpl 可注入（测试替身用）；默认真实读盘实现。
 * schema 与 03-mcp-server/server.py 逐字段一致（pi skill 生态通用）。
 */
export function createTools({ readSheetImpl = readSheet } = {}) {
	return [
		{
			// 工具1：结构侦察（帮 AI/用户选关联列和规则列）
			name: "inspect_sheet",
			description: "读取 Excel/CSV 表格结构：返回列名、总行数、前 5 行预览。用于核对前选列。",
			inputSchema: {
				type: "object",
				properties: {
					path: { type: "string", description: "表格文件路径（.xlsx/.csv）" },
					sheet: { type: "string", description: "工作表名，缺省为第一个" },
				},
				required: ["path"],
			},
			execute: async (args) => {
				const { readFileSync } = await import("node:fs"); // 按需引入（保持纯浏览器不可用的部分隔离）
				const { sheet, headers, rows } = readSheetImpl(readFileSync(args.path), args.sheet); // 读盘 → 统一结构
				return JSON.stringify({ // 结构摘要（preview 5 行防 token 爆炸）
					path: args.path, sheet, columns: headers, row_count: rows.length,
					preview: rows.slice(0, 5),
				}, null, 2);
			},
		},
		{
			// 工具2：双表核对主入口（业务核心，详细语义见 description）
			name: "compare_tables",
			description: "双表关联核对：以指定列（如订单号）为关联键，逐条执行核对规则，输出三类异常（区间不符/项不一致/单边缺失）并生成标红 Excel 报告。规则两种: type=range —— A方该列是范围(如 0.5-1kg、500-800g、小于1)，B方该列是具体值，B值不在A范围内即不符（单位自动换算，kg/g/公斤/吨/lb，裸数字默认kg，可用 unit_a/unit_b 指定）; type=exact —— 两列必须完全一致(数值按值比,文本归一化比)。",
			inputSchema: {
				type: "object",
				properties: {
					file_a: { type: "string", description: "A方表格路径（区间规则中 A=范围一方）" },
					file_b: { type: "string", description: "B方表格路径（区间规则中 B=具体值一方）" },
					key_a: { type: "string", description: "A方关联列名，如 订单号" },
					key_b: { type: "string", description: "B方关联列名" },
					rules: { // 规则数组（子 schema）
						type: "array",
						description: "核对规则列表",
						items: {
							type: "object",
							properties: {
								name: { type: "string", description: "规则显示名，如 重量" },
								type: { type: "string", enum: ["range", "exact"] }, // 两种规则
								col_a: { type: "string", description: "A方列名" },
								col_b: { type: "string", description: "B方列名" },
								tolerance: { type: "number", description: "容差(基准单位kg)，如0.05=±50g，仅range" },
								unit_a: { type: "string", description: "A方裸数字的默认单位，如 kg/g/公斤/吨/lb，缺省kg，仅range" },
								unit_b: { type: "string", description: "B方裸数字的默认单位，缺省kg，仅range" },
							},
							required: ["type", "col_a", "col_b"],
						},
					},
					sheet_a: { type: "string", description: "A方工作表名，缺省第一个" },
					sheet_b: { type: "string", description: "B方工作表名，缺省第一个" },
					tolerance: { type: "number", description: "全局容差，规则未指定时生效" },
					output: { type: "string", description: "报告输出路径，缺省 核对报告.xlsx（写在A方文件同目录或工作目录）" },
				},
				required: ["file_a", "file_b", "key_a", "key_b", "rules"],
			},
			execute: async (args) => {
				const { readFileSync, statSync } = await import("node:fs"); // 按需引入
				const { dirname, resolve } = await import("node:path");
				const a = readSheetImpl(readFileSync(args.file_a), args.sheet_a); // 读 A 方
				const b = readSheetImpl(readFileSync(args.file_b), args.sheet_b); // 读 B 方
				// 关联列存在性前置校验（错误信息带可选列，帮 AI 自纠）
				for (const [side, headers, key] of [["A", a.headers, args.key_a], ["B", b.headers, args.key_b]]) {
					if (!headers.includes(key)) throw new Error(`${side}方关联列 '${key}' 不存在，可选列: ${headers}`);
				}
				// 核心核对（纯函数）
				const result = compareData({
					headersA: a.headers, rowsA: a.rows, headersB: b.headers, rowsB: b.rows,
					keyA: args.key_a, keyB: args.key_b,
					rules: args.rules, tolerance: args.tolerance ?? 0,
				});
				// 报告路径：指定 > A方同目录 > 当前目录（与 Python 决策一致）
				let out = args.output;
				if (!out) {
					const dirA = dirname(args.file_a); // A 方文件目录
					out = (statSync(dirA, { throwIfNoEntry: false })?.isDirectory() ? resolve(dirA, "核对报告.xlsx") : "核对报告.xlsx");
				}
				await exportReport(result, out, { fileA: args.file_a, fileB: args.file_b }); // 7 表标红报告
				const slim = { ...result, _raw: undefined }; // 剔除导出原料（不可序列化：Map）
				slim.report = resolve(out); // 回填报告绝对路径
				return JSON.stringify(slim, null, 2); // MCP 文本载荷
			},
		},
	];
}

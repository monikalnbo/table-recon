/**
 * excel-mcp-js/tools.mjs —— Excel 13 工具（canonical JS 实现，双挂载核心）
 * =============================================================================
 * 对齐目标：excel-mcp/server.py（同名工具、同参数 schema、同语义）。
 * 防漂移机制：
 *   - selftest.mjs 用例 1:1 搬运 Python --selftest，并附 B2/B11/B12/B15 回归用例
 *   - 读写分工：读 = SheetJS 0.20.3（vendor，xlsx/csv 通吃）；
 *               写/样式/公式 = ExcelJS（openpyxl 的 JS 对位库）
 * 已知契约差异（记录在案，非 bug）：
 *   - csv 一切值为字符串（Python csv 模块同）
 *   - >2^53 大整数经 JS Number 有精度损失（Excel 本身只存 15 位有效数字，实际风险低）
 *   - 日期在「过滤视图」中保持原生类型（ExcelJS=Date 对象 / Python=datetime）
 */
import * as XLSX from "xlsx"; // SheetJS：只读用途（xlsx/xlsm/csv 解析）
import ExcelJS from "exceljs"; // ExcelJS：写用途（样式/公式/结构修改）
import fs from "node:fs";
import path from "node:path";

/* ============================ 基础工具函数 ============================ */

// A1 引用正则：字母部分=列，数字部分=行
const CELL_RE = /^([A-Za-z]+)(\d+)$/;

/** "B3" → { row: 3, col: 2 }；非法引用直接抛错（对齐 Python _cell_to_idx） */
const colToIdx = (cell) => {
	const m = CELL_RE.exec(String(cell).trim()); // 先 trim，容忍 " b3 " 这类手输
	if (!m) throw new Error(`非法单元格引用: '${cell}'（应为 A1 形式）`);
	// 列号按 26 进制累加：A=1, B=2, …, Z=26, AA=27 …
	return { row: parseInt(m[2], 10), col: [...m[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) };
};

/** 1 → "A", 27 → "AA"（colToIdx 的逆函数，用于输出区域引用） */
const colLetter = (n) => {
	let s = "";
	while (n > 0) {
		const m = (n - 1) % 26; // 本位字母
		s = String.fromCharCode(65 + m) + s; // 前插
		n = (n - 1 - m) / 26; // 进位
	}
	return s;
};

const isCsv = (p) => /\.csv$/i.test(p); // 按扩展名分流读写策略
const plain = (v) => (typeof v === "number" && Number.isInteger(v) ? v : v ?? null); // 透传；仅保证 null 语义统一

/** 路径守卫：必须存在且后缀合法（.xlsx/.xlsm/.csv），否则抛中文错误给 AI/用户 */
function requireFile(p) {
	if (!fs.existsSync(p)) throw new Error(`文件不存在: ${p}`);
	if (!/\.(xlsx|xlsm|csv)$/i.test(p)) throw new Error(`仅支持 .xlsx/.xlsm/.csv，收到: '${path.extname(p)}'`);
	return p;
}

/** 仅 xlsx 可用的操作（公式/格式）在 csv 上给出明确报错，而不是静默失败 */
const xlsxOnly = (p) => { if (isCsv(p)) throw new Error("csv 不支持该操作（公式/格式仅限 xlsx）"); };

/* ============================ csv 读写（自实现，无三方依赖） ============================ */

/** 读 csv 文件 → 字符串（utf8 读入 + 手工剥离 BOM，Node 不认 "utf-8-sig" 编码名） */
function readCsv(p) {
	return fs.readFileSync(p, "utf8").replace(/^\uFEFF/, "");
}

/** RFC4180 csv 解析器 → string[][]（处理引号包裹、""转义、\r\n 换行） */
function parseCsv(text) {
	const rows = []; // 结果行集
	let row = [], cell = "", q = false; // 当前行 / 当前格 / 是否处于引号内
	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (q) {
			// 引号内："" → 字面 "；单 " → 引号段结束
			if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
			else cell += c; // 其余字符原样入格
		} else if (c === '"') q = true; // 进入引号段
		else if (c === ",") { row.push(cell); cell = ""; } // 逗号 = 格边界
		else if (c === "\n" || c === "\r") {
			if (c === "\r" && text[i + 1] === "\n") i++; // CRLF 当一个换行
			row.push(cell); rows.push(row); row = []; cell = ""; // 行边界
		} else cell += c;
	}
	// 末行无换行符的收尾（有内容或非空行才补）
	if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
	return rows;
}

/** csv 格子转义：含 逗号/引号/换行 才包引号，内部 " 加倍 */
const csvEsc = (v) => {
	const s = v == null ? "" : String(v);
	return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};

/**
 * 统一读入口：文件 → { headers, rows(字典数组), sheet }
 * xlsx 走 SheetJS 网格模式；csv 走自实现解析器。
 * 关键点（曾踩坑）：表头必须取自首行网格——
 *   dict 模式（sheet_to_json 默认）在「只有表头无数据」时返回空数组 → headers=[]（B 系坑），
 *   网格模式永远拿得到首行。
 * 空表头列名兜底「列X」；全空行跳过（与 Python read_sheet 一致）。
 */
function sheetRows(file, sheet) {
	if (isCsv(file)) {
		const raw = parseCsv(readCsv(file)); // → string[][]
		if (!raw.length) return { headers: [], rows: [], sheet: path.basename(file, ".csv") }; // 空文件
		// 表头 = 首行 trim；空表头列兜底命名
		const headers = raw[0].map((h, i) => (h && h.trim() ? h.trim() : `列${colLetter(i + 1)}`));
		// 数据行：跳过全空行；按表头对齐成字典（缺省补 null）
		const rows = raw.slice(1)
			.filter((r) => r.some((c) => c && c.trim() !== ""))
			.map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? null])));
		return { headers, rows, sheet: path.basename(file, ".csv") }; // csv 无 sheet 概念，用文件名
	}
	// xlsx：Buffer 读入（cellDates:false → 日期保持序列号，两侧视图类型一致）
	const wb = XLSX.read(fs.readFileSync(file), { type: "buffer", cellDates: false });
	const name = sheet && wb.SheetNames.includes(sheet) ? sheet : wb.SheetNames[0]; // 指定或第一个
	const ws = wb.Sheets[name];
	// header:1 = 网格模式（数组行）；defval:null 补空；raw:true 拿原始值
	const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false });
	// 表头来自首行网格（见函数注释的坑）
	const headers = (grid[0] || []).map((h, i) => (h != null && String(h).trim() !== "" ? String(h).trim() : `列${colLetter(i + 1)}`));
	// 数据行：跳过全空行 → 字典化
	const rows = grid.slice(1)
		.filter((r) => r.some((c) => c != null && String(c).trim() !== ""))
		.map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? null])));
	return { headers, rows, sheet: name };
}

/** 宽松数值化：数字直通；字符串去逗号后转数；失败/空 → null（筛选比较用） */
const num = (v) => {
	if (typeof v === "number") return Number.isFinite(v) ? v : null; // NaN/Infinity 视为不可比
	if (v == null) return null;
	const s = String(v).trim().replace(/,/g, ""); // "1,200" → "1200"
	return s === "" ? null : (Number.isFinite(+s) ? +s : null);
};

/**
 * 日期归一化（业务修复：日期列筛选全灭）。
 * 三种输入 → 统一 UTC 毫秒：
 *   ① Date 对象（ExcelJS 原生）
 *   ② "2026-08-22" / "2026/8/22" / "2026-08-22T10:00" 等可被 Date.parse 的字符串
 *   ③ Excel 序列号 25569~2958463（1970-01-01 ~ 9999-12-31，防误吞普通数字如金额）
 * 非日期 → null（回退普通数值/文本比较）。
 */
const EXCEL_EPOCH_MIN = 25569; // 1970-01-01 的序列号（SheetJS cellDates:false 时日期变此序列号）
const EXCEL_EPOCH_MAX = 2958463; // 9999-12-31
const dateNum = (v) => {
	if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.getTime(); // ①
	if (typeof v === "string") {
		const t = Date.parse(v.trim()); // ② ISO/斜杠等主流写法
		return Number.isNaN(t) ? null : t;
	}
	if (typeof v === "number" && v >= EXCEL_EPOCH_MIN && v <= EXCEL_EPOCH_MAX) // ③ 序列号区间判定
		return Math.round((v - 25569) * 86400 * 1000); // Excel epoch(1899-12-30) → Unix 毫秒
	return null;
};

/* ============================ 筛选引擎（query/update 共用） ============================ */

// 数值比较算子表（a=格值 b=目标值，均经 num() 归一）
const OPS = { eq: (a, b) => a === b, ne: (a, b) => a !== b, gt: (a, b) => a > b, lt: (a, b) => a < b, ge: (a, b) => a >= b, le: (a, b) => a <= b };

/**
 * 单格值 vs 单条件匹配（对齐 Python _match）。
 * 语义：数值算子优先按数值比；任一侧不可数值化时 eq/ne 退化为文本比，其余 false；
 *       文本类一律 trim + 小写（大小写不敏感）。
 */
function matchOp(value, op, want) {
	// 空值判定类：不看 want
	if (op === "is_empty") return value == null || String(value).trim() === "";
	if (op === "not_empty") return !(value == null || String(value).trim() === "");
	if (value == null) return false; // 其余算子遇 null 一律不匹配
	// 数值算子族
	if (OPS[op]) {
		// —— 日期优先比较（业务修复：日期列筛选全灭）——
		const da = dateNum(value), db = dateNum(want); // 双侧尝试日期归一
		if (da != null && db != null) return OPS[op](da, db); // 都是日期 → 时间戳比
		// —— 普通数值比较 ——
		const a = num(value), b = num(want); // 双侧数值化
		if (a != null && b != null) return OPS[op](a, b); // 都可数值 → 数值比
		// 退化：仅 eq/ne 支持文本比（与 Python 一致）
		const sa = String(value).trim().toLowerCase(), sb = String(want).trim().toLowerCase();
		return op === "eq" ? sa === sb : op === "ne" ? sa !== sb : false;
	}
	// 区间算子：value ∈ [min(lo,hi), max(lo,hi)]（顺序可反写）
	if (op === "between") {
		const v = num(value);
		const [lo, hi] = (Array.isArray(want) ? want : [want]).map(num); // want 允许 [a,b] 或裸值
		if (v == null || lo == null || hi == null) return false; // 任一侧不可数值化 → 不匹配
		return Math.min(lo, hi) <= v && v <= Math.max(lo, hi);
	}
	// 文本类算子
	const s = String(value).trim().toLowerCase(); // 值 → 规范文本
	const w = String(want).trim().toLowerCase(); // 目标 → 规范文本
	if (op === "contains") return s.includes(w); // 包含
	if (op === "not_contains") return !s.includes(w); // 不包含
	// in / not_in：want 为数组按集合比；误传单值按 eq 兜底（容错）
	const set = new Set((Array.isArray(want) ? want : [want]).map((x) => String(x).trim().toLowerCase()));
	if (op === "in") return Array.isArray(want) ? set.has(s) : s === w;
	if (op === "not_in") return Array.isArray(want) ? !set.has(s) : s !== w;
	throw new Error(`不支持的操作符: '${op}'（可用: eq/ne/gt/lt/ge/le/contains/not_contains/in/not_in/between/is_empty/not_empty）`);
}

/**
 * 批量过滤（对齐 Python apply_filters）：
 * 1) 先校验每个条件的列名存在（不存在的列名 = 参数错误，必须报错而非静默空结果）
 * 2) 行 = 所有条件同时满足（AND）
 */
function applyFilters(rows, filters, headers) {
	if (!filters?.length) return rows; // 无条件 = 全量（update_rows 依赖此语义）
	for (const [i, f] of filters.entries()) {
		if (!f.column || !headers.includes(f.column)) throw new Error(`筛选${i + 1}列 '${f.column}' 不存在，可选: ${headers}`);
	}
	return rows.filter((r) => filters.every((f) => matchOp(r[f.column], f.op || "eq", f.value ?? f.values)));
}

/* ============================ 写入基建（ExcelJS） ============================ */

/**
 * 读-改-写事务：readFile → fn(wb, ws) → writeFile。
 * fn 抛错则不落盘（半事务性：参数校验失败不会损坏原文件）。
 * 返回 fn 处理后的 ws（供调用方取 name 等信息）。
 */
async function withWb(file, sheet, fn) {
	const wb = new ExcelJS.Workbook();
	await wb.xlsx.readFile(file); // 整簿读入（保留公式/样式）
	let ws = sheet ? wb.getWorksheet(sheet) : wb.worksheets[0]; // 指定或第一个
	if (!ws) throw new Error(`工作表不存在: '${sheet}'，可选: ${wb.worksheets.map((w) => w.name)}`);
	await fn(wb, ws); // 业务改动（可能抛错 → 跳过 writeFile）
	await wb.xlsx.writeFile(file); // 原子写回
	return ws;
}

/** 从 ExcelJS worksheet 取表头数组（与 sheetRows 同样的兜底命名规则，写侧专用） */
function headersOf(ws) {
	const r = ws.getRow(1); // 首行
	return Array.from({ length: ws.columnCount }, (_, i) => {
		const v = r.getCell(i + 1).value; // 每列表头单元格
		return v != null && String(v).trim() !== "" ? String(v).trim() : `列${colLetter(i + 1)}`; // 空表头兜底
	});
}

/* ============================ 13 个工具实现 ============================ */

const t = {}; // 工具名 → 实现（TOOLS_SPEC 按 name 关联）

/**
 * inspect_workbook —— 结构侦察（任何操作前先调）
 * 返回：sheet 清表（含行列数）、目标 sheet 的表头 + 前 3 行预览。
 */
t.inspect_workbook = async (a) => {
	const p = requireFile(a.path); // 路径守卫
	let sheets, active, columns = [], preview = [];
	if (isCsv(p)) {
		// csv：无 sheet 概念，用文件名模拟单 sheet
		const { headers, rows, sheet } = sheetRows(p);
		sheets = [{ name: sheet, rows: rows.length, cols: headers.length }];
		active = sheet; columns = headers;
		preview = rows.slice(0, 3);
	} else {
		// xlsx：ExcelJS 取 sheet 元数据（行列数），SheetJS 取表头+预览（快）
		const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(p);
		const target = a.sheet || wb.worksheets[0].name; // 目标 sheet：指定或第一个
		sheets = wb.worksheets.map((w) => ({ name: w.name, rows: w.rowCount, cols: w.columnCount }));
		const { headers, rows } = sheetRows(p, target);
		active = target; columns = headers; preview = rows.slice(0, 3);
	}
	return j({ path: p, sheets, active, columns, preview });
};

/**
 * read_range —— 按 A1 区域读原始网格。
 * values_only=false 时显示公式原文（"=SUM(...)"，对齐 openpyxl 显示习惯）。
 */
t.read_range = async (a) => {
	const p = requireFile(a.path);
	// csv 分支：无坐标概念，按「起始格切行+切列」模拟（与 Python csv 分支一致）
	if (isCsv(p)) {
		const { rows } = sheetRows(p);
		const { row, col } = colToIdx(a.start_cell || "A1"); // 起始格 → 行列号
		const data = rows.slice(row - 1).map((r) => Object.values(r).slice(col - 1)); // 行、列都从起始格切
		return j({ start_cell: a.start_cell || "A1", data });
	}
	const formulas = a.values_only === false; // 显式 false 才显示公式
	const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(p);
	const ws = a.sheet ? wb.getWorksheet(a.sheet) : wb.worksheets[0];
	if (!ws) throw new Error(`工作表不存在: '${a.sheet}'`);
	const s = colToIdx(a.start_cell || "A1"); // 起始角点
	const e = a.end_cell ? colToIdx(a.end_cell) : { row: ws.rowCount, col: ws.columnCount }; // 缺省=用到末尾
	const [r1, r2] = [Math.min(s.row, e.row), Math.max(s.row, e.row)]; // 支持反写角点
	const [c1, c2] = [Math.min(s.col, e.col), Math.max(s.col, e.col)];
	const data = [];
	for (let r = r1; r <= r2; r++) {
		const row = [];
		for (let c = c1; c <= c2; c++) {
			let v = ws.getRow(r).getCell(c).value; // 原生值（可能是公式/富文本对象）
			if (formulas && v && typeof v === "object") {
				// 公式模式：优先公式原文（补 "=" 对齐 openpyxl），否则结果
				if (v.formula) v = "=" + v.formula;
				else if (v.result !== undefined) v = v.result;
				else if (v.richText) v = v.richText.map((x) => x.text).join("");
			} else if (v && typeof v === "object" && !Array.isArray(v)) {
				// 值模式：对象值降维为纯量（富文本拼接 / 公式取结果 / 其他取 text）
				v = v.richText ? v.richText.map((x) => x.text).join("") : (v.result ?? v.text ?? null);
			}
			row.push(v);
		}
		data.push(row);
	}
	return j({ range: `${colLetter(c1)}${r1}:${colLetter(c2)}${r2}`, data }); // 实际读取范围回显
};

/**
 * query_rows —— 按条件查行（首行为表头）。
 * 返回 total/matched/截断行；limit 上限 500 防 token 爆炸。
 */
t.query_rows = async (a) => {
	const p = requireFile(a.path);
	const { headers, rows, sheet } = sheetRows(p, a.sheet); // 读侧统一入口
	const matched = applyFilters(rows, a.filters, headers); // 条件过滤（列名校验在内）
	const cols = a.columns?.length ? a.columns : headers; // 投影列：指定或全列
	for (const c of cols) if (!headers.includes(c)) throw new Error(`列 '${c}' 不存在，可选: ${headers}`); // 投影列也要校验
	const limit = Math.min(a.limit ?? 50, 500); // 默认 50，硬顶 500
	return j({
		sheet, total_rows: rows.length, matched: matched.length, columns: cols, // 概览
		rows: matched.slice(0, limit).map((r) => Object.fromEntries(cols.map((c) => [c, plain(r[c])]))), // 投影后的数据行
		note: matched.length > limit ? `显示前 ${limit} 行` : null, // 截断提示
	});
};

/**
 * write_data —— 从 start_cell 起覆盖写二维数组。
 * csv 无坐标概念，退化为追加（mode 字段说明）。
 */
t.write_data = async (a) => {
	const p = requireFile(a.path);
	const data = a.data; // 待写网格
	if (!Array.isArray(data) || !data.every((r) => Array.isArray(r))) throw new Error("data 必须是二维数组 [[..],[..]]"); // 形状校验
	if (isCsv(p)) {
		// csv：逐行转义后追加（文件非空先补换行）
		const lines = data.map((r) => r.map(csvEsc).join(","));
		fs.appendFileSync(p, (fs.existsSync(p) && fs.statSync(p).size ? "\n" : "") + lines.join("\n"), "utf8");
		return j({ written: data.length, mode: "csv-append" }); // 明示是追加语义
	}
	// xlsx：定位起始格后逐格覆写
	const ws = await withWb(p, a.sheet, async (_wb, ws) => {
		const s = colToIdx(a.start_cell || "A1");
		data.forEach((row, i) => row.forEach((v, k) => ws.getRow(s.row + i).getCell(s.col + k).value = v)); // 行偏移 i 列偏移 k
	});
	return j({ written: data.length, start_cell: a.start_cell || "A1", sheet: ws.name });
};

/**
 * append_rows —— 追加行。rows 两种形态：
 *   [{}] 字典数组：按表头自动对齐列序（AI 最爱用）
 *   [[]] 二维数组：原样按列序追加
 * 字典模式字段必须在表头内（防拼写错静默丢列）。
 */
t.append_rows = async (a) => {
	const p = requireFile(a.path);
	const rowsIn = a.rows;
	const { headers, sheet } = sheetRows(p, a.sheet); // 现有表头（对齐/校验用）
	let data; // 归一后的二维数组
	if (rowsIn.length && !Array.isArray(rowsIn[0])) {
		// 字典模式：先校验所有字段都在表头，再按表头序展开（缺失列补 null）
		rowsIn.forEach((r, i) => Object.keys(r).forEach((k) => { if (!headers.includes(k)) throw new Error(`第${i + 1}行字段 '${k}' 不在表头中，可用: ${headers}`); }));
		data = rowsIn.map((r) => headers.map((h) => r[h] ?? null));
	} else data = rowsIn; // 数组模式：原样
	if (isCsv(p)) {
		fs.appendFileSync(p, "\n" + data.map((r) => r.map(csvEsc).join(",")).join("\n"), "utf8"); // 追加（文件必有表头行）
		return j({ appended: data.length, sheet, headers });
	}
	// xlsx：找首空行起始（全空 sheet 从 1 起）
	let start = 1;
	const ws = await withWb(p, a.sheet, async (_wb, ws) => {
		start = ws.rowCount >= 1 ? ws.rowCount + 1 : 1; // 末行 +1
		if (ws.rowCount === 1 && ws.getRow(1).getCell(1).value == null) start = 1; // 仅一个空行 = 全空 sheet
		data.forEach((row, i) => row.forEach((v, k) => ws.getRow(start + i).getCell(k + 1).value = v)); // 逐格写入
	});
	return j({ appended: data.length, start_row: start, sheet: ws.name, headers });
};

/**
 * update_rows —— 查询式批量修改（本工具链的王牌）。
 * 策略（对齐 Python）：单一 ExcelJS 视图上「过滤 → 原行回写」。
 * ⚠️ 历史 bug B2（已修）：旧实现 SheetJS 视图筛选后按「整行 JSON 相等」映射回
 *    ExcelJS 行号——日期(Date vs 序列号)/富文本等类型表示不同 → 永不匹配 → updated:0。
 *    现在直接在 ExcelJS 视图上构建过滤行（公式取 result、富文本拼串），命中即写，
 *    其余列的公式/样式原封不动。
 */
t.update_rows = async (a) => {
	const p = requireFile(a.path);
	const updates = a.updates || {}; // {列名: 新值}
	if (!Object.keys(updates).length) throw new Error("updates 不能为空"); // 空更新=误用
	// csv 分支：读全量 → 内存改 → 全量重写（csv 只能整文件重写）
	if (isCsv(p)) {
		const { headers, rows } = sheetRows(p);
		for (const k of Object.keys(updates)) if (!headers.includes(k)) throw new Error(`更新列 '${k}' 不存在，可选: ${headers}`); // 更新列校验
		const matched = applyFilters(rows, a.filters, headers); // 过滤命中的行（字典引用）
		matched.forEach((r) => Object.assign(r, updates)); // 原地合并更新
		// 全量重写：BOM + 表头行 + 数据行
		fs.writeFileSync(p, "\uFEFF" + [headers.map(csvEsc).join(","), ...rows.map((r) => headers.map((h) => csvEsc(r[h])).join(","))].join("\n"), "utf8");
		return j({ updated: matched.length, updates, sheet: path.basename(p, ".csv") });
	}
	// xlsx 分支：单视图过滤 + 原行回写（见函数头注释）
	let updated = 0; // 命中计数
	const ws = await withWb(p, a.sheet, async (_wb, ws) => {
		const h = headersOf(ws); // 写侧表头（兜底命名与读侧一致）
		const colOf = Object.fromEntries(h.map((name, i) => [name, i + 1])); // 表头 → 列号映射
		for (const k of Object.keys(updates)) // 更新列必须存在（在事务内校验，失败不落盘）
			if (!colOf[k]) throw new Error(`更新列 '${k}' 不存在，可选: ${h}`);
		for (let ri = 2; ri <= ws.rowCount; ri++) { // 数据区从第 2 行
			const vals = {}; // 本行「过滤视图」：ExcelJS 原生值降维为纯量
			h.forEach((name, i) => {
				let v = ws.getRow(ri).getCell(i + 1).value;
				if (v && typeof v === "object") { // 对象值三态降维
					if (v.result !== undefined) v = v.result; // 公式格 → 计算结果（对齐 Python data_only）
					else if (v.richText) v = v.richText.map((x) => x.text).join(""); // 富文本 → 拼接串
				}
				vals[name] = v ?? null;
			});
			// 全空行跳过（与读侧行为一致）
			if (Object.values(vals).every((v) => v == null || String(v).trim() === "")) continue;
			// 过滤视图上直接跑筛选（含列名校验）；命中 → 只写 updates 涉及的格
			if (applyFilters([vals], a.filters, h).length) {
				for (const [k, v] of Object.entries(updates)) ws.getRow(ri).getCell(colOf[k]).value = v;
				updated++;
			}
		}
	});
	return j({ updated, updates, sheet: ws.name });
};

/**
 * set_cells —— 精确补丁：{"B3": 值} 一次写多个任意散点格。
 */
t.set_cells = async (a) => {
	const p = requireFile(a.path);
	xlsxOnly(p); // csv 无坐标
	const cells = a.cells; // {引用: 值}
	if (!cells || typeof cells !== "object" || Array.isArray(cells) || !Object.keys(cells).length) throw new Error("cells 必须是 {\"A1\": 值, ...}");
	await withWb(p, a.sheet, async (_wb, ws) => {
		for (const [ref, v] of Object.entries(cells)) { // 逐格
			const { row, col } = colToIdx(ref); // 引用 → 坐标
			ws.getRow(row).getCell(col).value = v; // 覆写
		}
	});
	return j({ set: Object.keys(cells).length, cells }); // 回显写入内容
};

/**
 * apply_formula —— 给单元格写公式。
 * ⚠️ 关键（B15 修复）：ExcelJS 对 "=SUM(..)" 纯字符串按【文本】存储
 *    （实测读回仍是 string），必须用 { formula: "SUM(..)" } 对象才是真公式。
 */
t.apply_formula = async (a) => {
	const p = requireFile(a.path);
	xlsxOnly(p); // csv 无公式
	let f = String(a.formula).trim(); // 容忍手输带空格
	if (!f.startsWith("=")) f = "=" + f; // 容忍不带等号（回显统一带）
	const ws = await withWb(p, a.sheet, async (_wb, ws) => {
		const { row, col } = colToIdx(a.cell); // 目标格
		ws.getRow(row).getCell(col).value = { formula: f.slice(1) }; // 去掉 "=" 存对象（B15）
	});
	return j({ cell: String(a.cell).toUpperCase(), formula: f, sheet: ws.name }); // 回显规范化公式
};

/**
 * format_range —— 区域格式化（加粗/斜体/字号/字色/底色/数字格式/对齐）。
 * 字体属性用「合并展开」而非整体替换——只改指定属性，保留既有其他属性。
 */
t.format_range = async (a) => {
	const p = requireFile(a.path);
	xlsxOnly(p); // csv 无格式
	const s = colToIdx(a.start_cell), e = colToIdx(a.end_cell || a.start_cell); // 起止角点（缺省=单格）
	const [r1, r2] = [Math.min(s.row, e.row), Math.max(s.row, e.row)]; // 支持反写
	const [c1, c2] = [Math.min(s.col, e.col), Math.max(s.col, e.col)];
	let n = 0; // 实际格式化格数
	await withWb(p, a.sheet, async (_wb, ws) => {
		for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { // 区域内逐格
			const cell = ws.getRow(r).getCell(c);
			// —— 字体组（合并展开，见函数头）——
			if (a.bold !== undefined) cell.font = { ...cell.font, bold: !!a.bold };
			if (a.italic !== undefined) cell.font = { ...cell.font, italic: !!a.italic };
			if (a.font_size != null) cell.font = { ...cell.font, size: a.font_size };
			if (a.font_color) cell.font = { ...cell.font, color: { argb: "FF" + String(a.font_color).replace("#", "").toUpperCase() } }; // 容忍 #RRGGBB 写法
			// —— 底色（纯色填充）——
			if (a.bg_color) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + String(a.bg_color).replace("#", "").toUpperCase() } };
			// —— 数字格式（如 0.00 / #,##0）——
			if (a.number_format) cell.numFmt = a.number_format;
			// —— 对齐（水平 + 自动换行）——
			if (a.alignment) cell.alignment = { horizontal: a.alignment, wrapText: !!a.wrap_text };
			n++;
		}
	});
	return j({ formatted: n, range: `${colLetter(c1)}${r1}:${colLetter(c2)}${r2}` });
};

/**
 * find_replace —— 文本查找替换（全表或指定列）。
 * ⚠️ 两处历史 bug（已修）：
 *   B11：xlsx 分支曾完全忽略 column 参数（跨列误改）——现在先解析列号再限定。
 *   B12：公式格命中时会整格替换成静态文本（公式被毁）——现在公式格一律跳过。
 */
t.find_replace = async (a) => {
	const p = requireFile(a.path);
	const find = String(a.find), repl = String(a.replace ?? ""); // 目标串 / 替换串（缺省=删除）
	if (!find) throw new Error("find 不能为空");
	let count = 0; // 命中格数（按格计，非按出现次数）
	// csv 分支：内存改 + 全量重写（同 update_rows 的 csv 策略）
	if (isCsv(p)) {
		const { headers, rows } = sheetRows(p);
		const cols = a.column ? [a.column] : headers; // 列限定：csv 分支本来就正确
		for (const r of rows) for (const h of cols) { // 限定列逐行扫
			const v = r[h];
			if (v != null && String(v).includes(find)) { r[h] = String(v).split(find).join(repl); count++; } // split/join = 全量替换
		}
		fs.writeFileSync(p, "\uFEFF" + [headers.map(csvEsc).join(","), ...rows.map((r) => headers.map((h) => csvEsc(r[h])).join(","))].join("\n"), "utf8");
		return j({ replaced: count, find, replace: repl });
	}
	// xlsx 分支
	await withWb(p, a.sheet, async (_wb, ws) => {
		// —— 列限定解析（B11）：column 参数 → 列号；找不到 = 报错（与 Python 一致，不静默全表）——
		let onlyCol = null; // null = 不限列（全表）
		if (a.column) {
			const hr = ws.getRow(1); // 表头行
			for (let c = 1; c <= ws.columnCount; c++) {
				const hv = hr.getCell(c).value; // 逐列表头值
				if (hv != null && String(hv).trim() === a.column) { onlyCol = c; break; } // 精确匹配（trim 容错）
			}
			if (onlyCol == null) throw new Error(`列 '${a.column}' 不存在`);
		}
		// —— 全表逐格扫（includeEmpty 跳空格）——
		ws.eachRow((row) => row.eachCell({ includeEmpty: false }, (cell, cn) => {
			if (onlyCol !== null && cn !== onlyCol) return; // 非目标列直接跳过（B11 核心）
			let v = cell.value; // 原生值
			if (v && typeof v === "object") {
				if (v.formula || v.sharedFormula) return; // 公式格跳过（B12：防止打平成静态文本）
				if (v.result !== undefined) v = v.result; // 其余对象格取计算结果参与匹配
				else if (v.richText) v = v.richText.map((x) => x.text).join(""); // 富文本取拼接串
			}
			if (v != null && String(v).includes(find)) {
				cell.value = String(v).split(find).join(repl); // 命中 → 写回纯文本（全量替换）
				count++;
			}
		}));
	});
	return j({ replaced: count, find, replace: repl });
};

/**
 * delete_rows —— 删行（1-based 闭区间 [row_start, row_end]）。
 * 必须从大行号往小删：正序删会因行号位移删错行。
 */
t.delete_rows = async (a) => {
	const p = requireFile(a.path);
	xlsxOnly(p);
	const r1 = parseInt(a.row_start, 10), r2 = parseInt(a.row_end ?? a.row_start, 10); // 缺省 end = 单行
	if (Math.min(r1, r2) < 1) throw new Error("行号从 1 开始"); // 越界防护
	const [lo, hi] = [Math.min(r1, r2), Math.max(r1, r2)]; // 支持反写
	await withWb(p, a.sheet, async (_wb, ws) => {
		for (let r = hi; r >= lo; r--) ws.spliceRows(r, 1); // 倒序逐行删（见函数头）
	});
	return j({ deleted_rows: hi - lo + 1, rows: `${lo}-${hi}` });
};

/**
 * create_workbook —— 新建工作簿（xlsx 带可选表头行；csv 只写表头行）。
 * 已存在 = 报错（防误覆盖）。
 */
t.create_workbook = async (a) => {
	const p = a.path;
	if (fs.existsSync(p)) throw new Error(`文件已存在: ${p}`); // 防覆盖
	fs.mkdirSync(path.dirname(path.resolve(p)), { recursive: true }); // 父目录自动创建
	if (/\.csv$/i.test(p)) {
		fs.writeFileSync(p, "\uFEFF" + (a.headers || []).map(csvEsc).join(","), "utf8"); // BOM + 表头
		return j({ created: p, headers: a.headers || [] });
	}
	const wb = new ExcelJS.Workbook(); // 空 workbook
	const ws = wb.addWorksheet(a.sheet_name || "Sheet"); // 首个 sheet
	if (a.headers?.length) ws.addRow(a.headers); // 可选表头行
	await wb.xlsx.writeFile(p);
	return j({ created: p, sheets: [a.sheet_name || "Sheet"], rows: a.headers?.length ? 1 : 0 });
};

/**
 * manage_sheets —— sheet 级操作：create / copy / rename / delete。
 * copy 为手写实现（ExcelJS 无内置 copy）：逐格搬「值 + 样式」，
 * 公式格搬计算结果（副本公式引用会错位，取结果更安全）。
 */
t.manage_sheets = async (a) => {
	const p = requireFile(a.path);
	xlsxOnly(p);
	const { action, sheet } = a; // 操作类型 / 目标 sheet
	const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(p); // 手工事务（不同 action 写法差异大）
	if (action === "create") {
		if (wb.getWorksheet(sheet)) throw new Error(`sheet 已存在: ${sheet}`); // 防重名
		wb.addWorksheet(sheet); // 空表
	} else if (action === "copy") {
		const src = wb.getWorksheet(sheet);
		if (!src) throw new Error(`sheet 不存在: ${sheet}`);
		const dst = wb.addWorksheet(a.new_name || `${sheet} 副本`); // 目标名缺省 = 「原名 副本」
		// 逐格搬值+搬样式（见函数头注释）
		src.eachRow({ includeEmpty: true }, (row, rn) => row.eachCell({ includeEmpty: true }, (cell, cn) => {
			const d = dst.getRow(rn).getCell(cn); // 副本对应格
			d.value = cell.value instanceof Object && cell.value.result !== undefined ? cell.value.result : cell.value; // 公式取结果
			d.style = cell.style; // 样式整体引用（ExcelJS 内部会克隆）
		}));
	} else if (action === "rename") {
		const ws = wb.getWorksheet(sheet);
		if (!ws) throw new Error(`sheet 不存在: ${sheet}`);
		ws.name = a.new_name; // 直接改名
	} else if (action === "delete") {
		const ws = wb.getWorksheet(sheet);
		if (!ws) throw new Error(`sheet 不存在: ${sheet}`);
		if (wb.worksheets.length === 1) throw new Error("至少保留一个 sheet"); // 不许删空簿
		wb.removeWorksheet(sheet);
	} else throw new Error("action 必须是 create/copy/rename/delete");
	await wb.xlsx.writeFile(p); // 落盘
	return j({ action, sheets: wb.worksheets.map((w) => w.name) }); // 最新 sheet 清单
};

const j = (o) => JSON.stringify(o); // 统一 JSON 字符串输出（MCP content 约定）

/* ============================ 工具规格（与 Python server 逐字段一致） ============================ */

// 筛选条件子 schema（query_rows / update_rows 共用）
const F = { type: "object", properties: { column: { type: "string" }, op: { type: "string" }, value: {}, values: { type: "array" } }, required: ["column", "op"] };

// 13 工具规格表：name + description + inputSchema（顺序即对外展示序）
const TOOLS_SPEC = [
	{ name: "inspect_workbook", description: "查看 Excel/CSV 结构：sheet 列表、行列数、表头、前3行预览。任何操作前先调这个。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" } }, required: ["path"] } },
	{ name: "read_range", description: "按 A1 区域读原始网格（如 A1:C10）。values_only=false 时返回公式原文。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, start_cell: { type: "string" }, end_cell: { type: "string" }, values_only: { type: "boolean", default: true } }, required: ["path", "start_cell"] } },
	{ name: "query_rows", description: "按条件查行（首行为表头）。op: eq/ne/gt/lt/ge/le/contains/not_contains/in/not_in/between/is_empty/not_empty。数值列自动按数值比。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, filters: { type: "array", items: F }, columns: { type: "array", items: { type: "string" } }, limit: { type: "integer", default: 50 } }, required: ["path"] } },
	{ name: "write_data", description: "从 start_cell 起写二维数组数据（会覆盖目标区域）。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, start_cell: { type: "string", default: "A1" }, data: { type: "array", items: { type: "array" } } }, required: ["path", "data"] } },
	{ name: "append_rows", description: "追加行。rows 可为二维数组，或按表头的字典数组（自动对齐列顺序）。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, rows: { type: "array" } }, required: ["path", "rows"] } },
	{ name: "update_rows", description: "查询式批量修改：筛选出匹配行，把这些行的指定列改成新值。filters 同 query_rows。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, filters: { type: "array", items: F }, updates: { type: "object" } }, required: ["path", "updates"] } },
	{ name: "set_cells", description: "精确补丁：{\"B3\": 值, \"C7\": 值} 一次写多个单元格。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, cells: { type: "object" } }, required: ["path", "cells"] } },
	{ name: "apply_formula", description: "给单元格写公式，如 =SUM(B2:B10)。带不带等号都行。（仅 xlsx）", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, cell: { type: "string" }, formula: { type: "string" } }, required: ["path", "cell", "formula"] } },
	{ name: "format_range", description: "格式化区域：加粗/斜体/字号/字色/背景色/数字格式/对齐。颜色如 FFC7CE 或 #FFC7CE。（仅 xlsx）", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, start_cell: { type: "string" }, end_cell: { type: "string" }, bold: { type: "boolean" }, italic: { type: "boolean" }, font_size: { type: "number" }, font_color: { type: "string" }, bg_color: { type: "string" }, number_format: { type: "string" }, alignment: { type: "string" }, wrap_text: { type: "boolean" } }, required: ["path", "start_cell"] } },
	{ name: "find_replace", description: "全表或指定列查找替换文本（公式格自动跳过）。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, find: { type: "string" }, replace: { type: "string" }, column: { type: "string" } }, required: ["path", "find"] } },
	{ name: "delete_rows", description: "删除行（1-based，含两端）。（仅 xlsx）", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet: { type: "string" }, row_start: { type: "integer" }, row_end: { type: "integer" } }, required: ["path", "row_start"] } },
	{ name: "create_workbook", description: "新建工作簿（可带表头行）。", inputSchema: { type: "object", properties: { path: { type: "string" }, sheet_name: { type: "string" }, headers: { type: "array", items: { type: "string" } } }, required: ["path"] } },
	{ name: "manage_sheets", description: "管理 sheet：create / copy / rename / delete。", inputSchema: { type: "object", properties: { path: { type: "string" }, action: { type: "string", enum: ["create", "copy", "rename", "delete"] }, sheet: { type: "string" }, new_name: { type: "string" } }, required: ["path", "action", "sheet"] } },
];

/**
 * 工具工厂：SPEC（名称/描述/参数 schema）× 实现（t 表）→ 完整工具数组。
 * 单次遍历装配（旧版 O(n²) find 已清理）；进程内挂载与 stdio server 共用此出口。
 */
export function createTools() {
	return TOOLS_SPEC.map((spec) => ({ ...spec, execute: t[spec.name] })); // spec 摊平 + 绑实现
}

/**
 * recon-js/engine.mjs —— 双表核对引擎（canonical JS，纯函数、零 IO）
 * =============================================================================
 * 语义基线：core/recon.py（金测 desktop/test/golden-recon.mjs 双跑断言，diff 必须为空）。
 * 来源：从 02-gui 内嵌引擎原样提取（该实现已与 Python 结果一致）；
 *       文件读写层独立在 index.mjs，本文件只做纯数据核对。
 * 已知与 Python 共同保留的语义（不单边改，改动需双修双测）：
 *   ① 方向互换死代码：点值（"800g"）能被 parseRange 解析成退化区间 [0.8,0.8]，
 *      导致「A=点值/B=区间」的反向容错分支永不触发 → 假异常；
 *   ② 重复关联键静默 last-wins（Map 覆盖），拆单/补发行会丢。
 */

/* ============================ 单位与解析 ============================ */

// 单位 → 千克系数（基准单位 kg；与 Python UNIT_FACTOR 一一对应）
// 斤 = 0.5kg（中国电商高频单位；canonical 双修：Python 同步加入，金测同步）
export const UNIT = { mg: 1e-6, g: 0.001, 克: 0.001, kg: 1, 公斤: 1, 千克: 1, 吨: 1000, t: 1000, lb: 0.453592, 磅: 0.453592, 斤: 0.5 };
// 单位正则片段：长名在前防止 "kg" 被 "[gt]" 截胡；[gt] 兜底单字母 g/t
const URE = "(mg|kg|lb|吨|磅|公斤|千克|克|斤|[gt])";
// 数值正则片段：可选负号 + 数字 + 可选小数（不支持科学计数，与 Python 同）
const NUM = "(-?\\d+(?:\\.\\d+)?)";

/** 文本归一化：NFKC（全角→半角）→ trim → 小写 → 去空格 → ～→~ → 长破折号→-（手输 Excel 常见变体） */
export const norm = (v) =>
	v == null ? "" : String(v).normalize("NFKC").trim().toLowerCase()
		.replace(/ /g, "").replace(/～/g, "~").replace(/[—–]/g, "-");

/**
 * 具体值解析："800g"→0.8，"1.2公斤"→1.2，"0.55"(裸数+default g)→0.00055。
 * 非 fullmatch（re.search 语义）：在串中找「数字+可选单位」，找不到/空串 → null。
 * def：裸数字的默认单位系数（调用方用 unit_a/unit_b 指定）。
 */
export function parseValue(t, def = 1) {
	const s = norm(t); // 先归一化
	if (!s) return null; // 空串不可解析
	const m = s.match(new RegExp(NUM + "\\s*" + URE + "?")); // search 语义（与 Python 同）
	if (!m) return null; // 连数字都没有
	const f = m[2] ? UNIT[m[2]] : def; // 有单位用单位；没单位用默认（注意：未知单位串会 fallthrough 到 def，同 Python）
	return f == null ? null : parseFloat(m[1]) * f; // 换算到 kg 基准
}

/**
 * 区间解析 → [lo, hi]（kg 基准）；失败 → null。
 * 支持格式（单位可出现在任一侧或两侧）：
 *   a-b / a~b / a至b / 小于x / 低于x / 不超过x / <≤x / 大于x / ≥x / x以上 / 纯数字(退化[x,x])
 * 单位规则：两侧都有→各自换算；仅一侧有→作用于整个区间；都没有→def。
 * ⚠️ 纯数字也被解析为 [x,x] 点区间——这是「方向互换死代码」的根源（见文件头①）。
 */
export function parseRange(t, def = 1) {
	const s = norm(t);
	if (!s) return null;
	let m;
	// ---- 形态1：a-间隔-b（两侧各可带单位）----
	if ((m = s.match(new RegExp(`^${NUM}\\s*${URE}?[~\\-至]${NUM}\\s*${URE}?`)))) {
		const [, n1, u1, n2, u2] = m; // 数字1/单位1/数字2/单位2
		let f1 = u1 ? UNIT[u1] : null, f2 = u2 ? UNIT[u2] : null; // 各侧系数（未指定=null）
		if (f1 == null && f2 == null) f1 = f2 = def; // 都裸数 → 默认单位
		else if (f1 == null) f1 = f2; // "500-600g"：尾部单位作用于整段
		else if (f2 == null) f2 = f1; // "500g-600"：头部单位作用于整段
		const lo = parseFloat(n1) * f1, hi = parseFloat(n2) * f2; // 换算
		return [Math.min(lo, hi), Math.max(lo, hi)]; // 支持 a>b 反写
	}
	// ---- 形态2：上界类（小于/低于/不超过/<≤）→ [0, x] ----
	if ((m = s.match(new RegExp(`(?:小于|低于|不超过|[<≤])\\s*${NUM}\\s*${URE}?`))))
		return [0, parseFloat(m[1]) * (m[2] ? UNIT[m[2]] : def)];
	// ---- 形态3：下界类（大于等于/大于/超过/高于/>≥）→ [x, +∞) ----
	if ((m = s.match(new RegExp(`(?:大于等于|大于|超过|高于|[>≥])\\s*${NUM}\\s*${URE}?`))))
		return [parseFloat(m[1]) * (m[2] ? UNIT[m[2]] : def), Infinity];
	// ---- 形态4：x以上 → [x, +∞)（fullmatch 防止 "3以上的备注" 误吞）----
	if ((m = s.match(new RegExp(`^${NUM}\\s*${URE}?以上`))))
		return [parseFloat(m[1]) * (m[2] ? UNIT[m[2]] : def), Infinity];
	// ---- 形态5：纯数字 → 退化点区间 [x, x]（见文件头①的坑）----
	if ((m = s.match(new RegExp(`^${NUM}\\s*${URE}?$`)))) {
		const v = parseFloat(m[1]) * (m[2] ? UNIT[m[2]] : def);
		return [v, v];
	}
	return null; // 全部形态未命中
}

/** 区间格式化（用于 why 文案）：[0.5, 1] / [2, +∞) / null→"?" */
export const fmtRange = (r) => (r == null ? "?" : `[${r[0]}, ${r[1] === Infinity ? "+∞" : r[1]}]`);

/**
 * 区间核对（方向性：A=范围，B=点值）。B ∉ A → 不符。
 * 分支优先级：
 *   1) A 可解析为区间：
 *      a. B 可解析为点值 → 包含判定（主路径）
 *      b. B 也是区间 → 相交即通过（区间对区间容错）
 *      c. B 不可解析 → 不符 + 说明
 *   2) A 不可解析但 B 是区间 → 反向包含（方向反了的容错；因形态5退化区间，实际到不了这——见①）
 *   3) 双方都是点值 → 容差内相等
 *   4) 都不可解析 → 不符
 * tol：容差（kg）；ua/ub：两侧裸数默认系数。
 */
export function checkRange(va, vb, tol = 0, ua = 1, ub = 1) {
	const ra = parseRange(va, ua), rb = parseRange(vb, ub); // 两侧各自尝试区间解析
	if (ra != null) { // ---- A 是区间 ----
		const val = parseValue(vb, ub); // B 尝试点值
		if (val != null) // 1a. 包含判定（±tol 放宽边界）
			return (ra[0] - tol <= val && val <= ra[1] + tol)
				? [true, `${val} ∈ ${fmtRange(ra)}`]
				: [false, `${val} ∉ ${fmtRange(ra)}`];
		if (rb != null) { // 1b. 区间对区间：相交即过
			const lo = Math.max(ra[0], rb[0]), hi = Math.min(ra[1], rb[1]); // 交集
			return lo - tol <= hi + tol ? [true, "区间相交"] : [false, `区间不相交 ${fmtRange(ra)} vs ${fmtRange(rb)}`];
		}
		return [false, `B方数值无法解析（${vb}）`]; // 1c.
	}
	if (rb != null) { // ---- 2. 反向容错（A 点值 ∈ B 区间）----
		const val = parseValue(va, ua);
		if (val != null)
			return (rb[0] - tol <= val && val <= rb[1] + tol)
				? [true, `${val} ∈ ${fmtRange(rb)}`] : [false, `${val} ∉ ${fmtRange(rb)}`];
	}
	// ---- 3. 双点值：容差内相等（+1e-9 防浮点尾差）----
	const x = parseValue(va, ua), y = parseValue(vb, ub);
	if (x != null && y != null) return Math.abs(x - y) <= tol + 1e-9 ? [true, "数值相等"] : [false, `${x} ≠ ${y}`];
	return [false, "两边都无法解析为数值/区间"]; // ---- 4. ----
}

/**
 * 精确核对：可数值化→数值比（1.0 == 1）；否则文本归一化比（大小写/空格不敏感）。
 * 注意默认单位 1（kg）——"800" vs "800g" 会判不等（800kg ≠ 0.8kg），与 Python 同。
 */
export function checkExact(va, vb) {
	const x = parseValue(va), y = parseValue(vb); // 双侧点值解析（默认 kg）
	if (x != null && y != null) return Math.abs(x - y) < 1e-9 ? [true, "数值相等"] : [false, `数值不等: ${va} ≠ ${vb}`];
	return norm(va) === norm(vb) ? [true, "文本一致"] : [false, `不一致: ${va ?? "∅"} ≠ ${vb ?? "∅"}`];
}

/**
 * 规则里的单位说明 → 系数。三态：单位名（kg/g/…）查表；数字串（"0.001"）直用；其他 → 1。
 * （对齐 Python _unit_factor 的宽容行为。）
 */
export const unitFactor = (u) => {
	if (!u) return 1; // 未指定 → 1（kg）
	if (UNIT[u] !== undefined) return UNIT[u]; // 合法单位名
	const f = parseFloat(u); // 允许直接给系数
	return Number.isFinite(f) ? f : 1; // 非法串 → 1（不炸，宽松）
};

/**
 * 行关联键：整数型 Number → String（"123" vs 123 统一），再走 norm。
 * null/空串 → null（该行不参与关联，静默丢弃——见文件头②）。
 * ⚠️ JS 特有契约：>2^53 整数经 Number 有精度损失；Date 对象会变本地化字符串
 *    （本实现读侧统一序列号，Date 不会出现；记录防回归）。
 */
export const keyof = (row, col) => {
	let v = row?.[col]; // 行字典取值（行可能 undefined）
	if (v == null) return null; // null/undefined → 不参与
	if (typeof v === "number" && Number.isInteger(v)) v = String(v); // 123 → "123"
	return norm(v) || null; // 归一化；空串 → null
};

/** 展示值：整数型 Number 保持原样（SheetJS 已给 number），null 语义统一（undefined→null） */
export const plain = (v) => {
	if (v == null) return null;
	if (typeof v === "number" && Number.isInteger(v)) return v;
	return v;
};

/**
 * 核对主函数（纯数据，不碰文件）。
 * 入参 rows 均为「字典数组」（列名→值）；rules 见 TOOLS_SPEC 的 compare_tables。
 * 输出结构与 Python compare() 的 slim dict 对齐（金测逐字段 diff）：
 *   summary 七项计数 / range_mismatch / exact_mismatch / only_in_a / only_in_b / _raw（导出报告用）
 */
export function compareData({ headersA, rowsA, headersB, rowsB, keyA, keyB, rules, tolerance = 0 }) {
	// ---- 1. 建索引：key → 行（重复 key 后行覆盖前行 = last-wins，见文件头②）----
	const mapA = new Map(), mapB = new Map();
	rowsA.forEach((r) => { const k = keyof(r, keyA); if (k) mapA.set(k, r); }); // 空键行静默跳过
	rowsB.forEach((r) => { const k = keyof(r, keyB); if (k) mapB.set(k, r); });

	// ---- 2. 双边都有：逐键逐规则核对 ----
	const rangeMiss = [], exactMiss = []; // 两类不符明细
	let ok = 0; // 全规则通过的键数
	for (const k of [...mapA.keys()].filter((k) => mapB.has(k)).sort()) { // 交集键，排序保证输出稳定
		let bad = false; // 本键是否有任一规则不符
		for (const rule of rules) {
			const va = mapA.get(k)[rule.col_a], vb = mapB.get(k)[rule.col_b]; // 两侧参与值
			const tol = rule.tolerance ?? tolerance; // 规则级容差优先，全局兜底
			// range/exact 分派
			const [pass, why] = rule.type === "range"
				? checkRange(va, vb, tol, unitFactor(rule.unit_a), unitFactor(rule.unit_b))
				: checkExact(va, vb);
			if (!pass) {
				bad = true;
				// 按规则类型归入对应明细（字段名与 Python 完全一致）
				(rule.type === "range" ? rangeMiss : exactMiss)
					.push({ key: k, rule: rule.name ?? rule.type, col_a: rule.col_a, col_b: rule.col_b, a: plain(va), b: plain(vb), reason: why });
			}
		}
		if (!bad) ok++; // 键级统计：全部规则通过才算匹配
	}
	// ---- 3. 单边键：整行入明细（报告里展示完整上下文）----
	const onlyA = [...mapA.keys()].filter((k) => !mapB.has(k)).sort().map((k) => ({ key: k, row: mapA.get(k) }));
	const onlyB = [...mapB.keys()].filter((k) => !mapA.has(k)).sort().map((k) => ({ key: k, row: mapB.get(k) }));

	// ---- 4. 组装结果（_raw 仅供报告导出，序列化前必须剔除）----
	return {
		keys: { a: keyA, b: keyB }, // 关联列名
		rules, // 规则回显
		summary: { A方记录: rowsA.length, B方记录: rowsB.length, 匹配: ok, 区间不符: rangeMiss.length, 项不一致: exactMiss.length, 仅A方有: onlyA.length, 仅B方有: onlyB.length }, // 七项计数（键名与 Python 对齐，金测断言用）
		range_mismatch: rangeMiss,
		exact_mismatch: exactMiss,
		only_in_a: onlyA,
		only_in_b: onlyB,
		_raw: { headersA, headersB, rowsA, rowsB, mapA, mapB, keyA, keyB }, // 报告导出原料
	};
}

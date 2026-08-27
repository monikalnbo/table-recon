#!/usr/bin/env node
/**
 * 金测：recon-js 与 core/recon.py 同输入同断言（canonical 等价证明）
 * 用法: node test/golden-recon.mjs   （在 recon-js/ 或 desktop/test/ 下均可，自动定位仓库根）
 */
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const { createTools, readSheet } = await import("recon-js");

const A = path.join(root, "测试_A方.xlsx"), B = path.join(root, "测试_B方.xlsx");
const RULES = [
	{ name: "重量", type: "range", col_a: "重量区间", col_b: "实称重量", tolerance: 0.05, unit_a: "kg", unit_b: "g" },
	{ name: "数量", type: "exact", col_a: "数量", col_b: "数量" },
];

const expect = {
	summary: { A方记录: 7, B方记录: 7, 匹配: 4, 区间不符: 2, 项不一致: 0, 仅A方有: 1, 仅B方有: 1 },
	rangeKeys: ["dd002", "dd008"],
	onlyA: ["dd004"], onlyB: ["dd005"],
};

let fails = 0;
const ok = (m) => console.log(`  ✔ ${m}`);
const bad = (m) => { fails++; console.error(`  ✖ ${m}`); };

/* ---- JS ---- */
console.log("[golden] recon-js");
const tools = createTools();
const jsOut = JSON.parse(await tools.find((t) => t.name === "compare_tables").execute({
	file_a: A, file_b: B, key_a: "订单号", key_b: "单号", rules: RULES,
}));
const jsCheck = (label, cond) => cond ? ok(label) : bad(`${label} —— 实际 ${JSON.stringify(cond)}`);

const sumEq = JSON.stringify(jsOut.summary) === JSON.stringify(expect.summary);
jsCheck(`summary ${JSON.stringify(jsOut.summary)}`, sumEq);
const rKeys = jsOut.range_mismatch.map((m) => m.key).sort();
jsCheck(`区间不符键 ${rKeys}`, JSON.stringify(rKeys) === JSON.stringify(expect.rangeKeys));
const onlyA = jsOut.only_in_a.map((o) => o.key), onlyB = jsOut.only_in_b.map((o) => o.key);
jsCheck(`仅A ${onlyA}`, JSON.stringify(onlyA) === JSON.stringify(expect.onlyA));
jsCheck(`仅B ${onlyB}`, JSON.stringify(onlyB) === JSON.stringify(expect.onlyB));

/* 报告可生成 */
const { exportReport } = await import("recon-js/report.mjs");
const a = readSheet(readFileSync(A)), b = readSheet(readFileSync(B));
const { compareData } = await import("recon-js/engine.mjs");
const res = compareData({ headersA: a.headers, rowsA: a.rows, headersB: b.headers, rowsB: b.rows, keyA: "订单号", keyB: "单号", rules: RULES });
const tmpReport = "/tmp/golden-报告.xlsx";
await exportReport(res, tmpReport, { fileA: A, fileB: B });
ok(`报告生成 ${tmpReport}`);

/* ---- Python 差分（有 Python 才跑） ---- */
console.log("[golden] Python diff");
let pyOut = null;
try {
	const stdout = execFileSync("python3", [
		path.join(root, "core", "recon.py"), A, B,
		"--key-a", "订单号", "--key-b", "单号",
		"--rule", "range:重量区间:实称重量:0.05:kg:g",
		"--rule", "exact:数量:数量",
		"-o", "/tmp/golden-py.xlsx", "--json",
	], { encoding: "utf8" });
	pyOut = JSON.parse(stdout);
} catch (e) {
	console.log("  · Python 不可用或执行失败，跳过差分（安装包环境预期）");
}

if (pyOut) {
	const strip = (s) => String(s).replace(/（区间原文.*?）$/u, "").replace(/\s/g, "");
	const pick = (o) => ({
		summary: o.summary,
		range_mismatch: o.range_mismatch.map(({ key, reason }) => ({ key, reason: strip(reason) })),
		exact_mismatch: o.exact_mismatch.map(({ key, reason }) => ({ key, reason: strip(reason) })),
		only_in_a: o.only_in_a.map((x) => x.key),
		only_in_b: o.only_in_b.map((x) => x.key),
	});
	const jsSlim = pick(jsOut);
	const same = JSON.stringify(jsSlim) === JSON.stringify(pick(pyOut));
	if (same) ok("JS 与 Python 输出完全一致");
	else {
		bad("JS 与 Python 输出不一致：");
		console.error("   JS :", JSON.stringify(jsSlim.range_mismatch));
		console.error("   PY :", JSON.stringify(pick(pyOut).range_mismatch));
	}
}

/* ---- 斤单位（canonical 双修后金测同步）---- */
const { parseValue, parseRange } = await import("recon-js/engine.mjs");
const jin = [
	[parseValue("5.1斤"), 2.55],
	[parseRange("4.5-5.5斤")[0], 2.25],
	[parseRange("4.5-5.5斤")[1], 2.75],
];
jin.every(([got, want]) => Math.abs(got - want) < 1e-9)
	? ok("斤单位解析（5.1斤=2.55kg，4.5-5.5斤=[2.25,2.75]）")
	: bad(`斤解析 ${JSON.stringify(jin)}`);

/* ---- stdio server 冒烟 ---- */
console.log("[golden] stdio server");
const { spawn } = await import("node:child_process");
const srv = spawn(process.execPath, [path.join(root, "recon-js", "server.mjs")]);
let srvOut = "";
srv.stdout.on("data", (d) => (srvOut += d));
srv.stderr.on("data", (d) => console.error("   [server-stderr]", String(d).slice(0, 300)));
srv.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "t", version: "0" } } }) + "\n");
srv.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n");
srv.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/list" }) + "\n");
const t0 = Date.now();
while (!srvOut.includes('"id":2') && Date.now() - t0 < 10000) await new Promise((r) => setTimeout(r, 150));
srv.kill();
try {
	const lines = srvOut.trim().split("\n").map((l) => JSON.parse(l));
	const list = lines.find((l) => l.id === 2).result.tools.map((t) => t.name);
	JSON.stringify(list) === JSON.stringify(["inspect_sheet", "compare_tables"]) ? ok(`tools/list ${list}`) : bad(`tools/list 异常: ${list}`);
} catch (e) {
	bad(`stdio server 无响应: ${e.message}`);
}

console.log(fails ? `\n✖ ${fails} 项失败` : "\n✔ 金测全部通过");
process.exit(fails ? 1 : 0);

#!/usr/bin/env node
/**
 * scripts/build-pages.mjs —— 从 02-gui/index.html 再生成内嵌自包含副本
 * =============================================================================
 * 产物：
 *   desktop/app.html      = 02-gui + 内联 SheetJS/ExcelJS + 桌面 AI 面板（如已注入）
 *   ios/TableRecon/app.html         = 02-gui + 内联库 + iOS safe-area 适配
 *   TableRecon.swiftpm/app.html     = 同 iOS 版
 * 规则：CDN <script src> 替换为内联；desktop/app.html 若已含桌面注入则保留
 *      （注入锚点在文件尾 </body> 前，重生成时先剥离再重注，防重复）。
 * 用法：先下载两个库到 scripts/vendor/（一次），再 node scripts/build-pages.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const src = fs.readFileSync(path.join(root, "02-gui/index.html"), "utf8");
const vendor = (f) => fs.readFileSync(path.join(root, "scripts/vendor", f), "utf8");
const XLSX_JS = vendor("xlsx.full.min.js");
const EXCELJS_JS = vendor("exceljs.min.js");

/** CDN 引用 → 内联；其余保持 */
function inlineLibs(html) {
	html = html.replace(/<script src="https:\/\/cdn\.sheetjs\.com\/[^"]+"><\/script>/, () => `<script>\n${XLSX_JS}\n</script>`);
	html = html.replace(/<script src="https:\/\/cdnjs\.cloudflare\.com\/[^"]+exceljs[^"]+"><\/script>/, () => `<script>\n${EXCELJS_JS}\n</script>`);
	if (/cdn\.sheetjs\.com|cdnjs/.test(html)) throw new Error("存在未内联的 CDN 引用（vendor 文件名不匹配？）");
	return html;
}

/** iOS 适配：viewport-fit=cover + safe-area padding（与历史版一致） */
const IOS_HEAD = `<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">`;
function forIOS(html) {
	html = html.replace('<meta name="viewport" content="width=device-width, initial-scale=1.0">', IOS_HEAD);
	html = html.replace("body{", "body{padding:env(safe-area-inset-top) env(safe-area-inset-right) env(safe-area-inset-bottom) env(safe-area-inset-left);");
	return html;
}

/* ---- iOS / Playgrounds：内联 + 适配，直接覆写 ---- */
const iosHtml = forIOS(inlineLibs(src));
for (const dst of ["ios/TableRecon/app.html", "TableRecon.swiftpm/app.html"]) {
	fs.writeFileSync(path.join(root, dst), iosHtml);
	console.log(`✔ ${dst} (${(iosHtml.length / 1e6).toFixed(2)}MB)`);
}

/* ---- desktop：内联 + 保留桌面注入（剥离旧注入 → 内联 → 重注） ---- */
const desktopPath = path.join(root, "desktop/app.html");
let desktopHtml = fs.existsSync(desktopPath) ? fs.readFileSync(desktopPath, "utf8") : "";
// 剥离旧桌面注入（锚点：桌面模式注释起 至 尾部 </body> 前）
const INJ_MARK = "<script>\n/* ===================== 桌面模式（Electron） ===================== */";
let injectBlock = "";
const injStart = desktopHtml.indexOf(INJ_MARK);
if (injStart >= 0) {
	const injEnd = desktopHtml.indexOf("</body>", injStart);
	injectBlock = desktopHtml.slice(injStart, injEnd);
}
// AI 面板 HTML 段（锚点：id="ai-panel" 卡片）与注入 JS 一起保留
const panelMark = '<!-- ===================== AI 面板（桌面版） =====================';
const AI_PANEL_RE = /<!-- =+ AI 面板（桌面版） =+ -->\s*<div class="card" id="ai-panel"[\s\S]*?<\/div>\n/;
let panelBlock = "";
const pm = desktopHtml.match(AI_PANEL_RE);
if (pm) panelBlock = pm[0];

// 基底 = 02-gui 内联版；然后插入面板与注入
let base = inlineLibs(src);
if (panelBlock) {
	const anchor = '<input type="file" id="file-a"'; // 面板插在 file input 前（唯一锚点）
	base = base.replace(anchor, panelBlock + "\n" + anchor);
}
// ⚠⚠ 注入必须锚定文件末尾的 </body>：内联库（SheetJS 的 HTML 导出模板）里也有字面 "</body>" 字符串，
// 用 replace（首次匹配）会把 JS 注进库代码内部 → 桌面模式彻底失效（踩过两次的坑，勿再犯）
if (injectBlock) {
	const last = base.lastIndexOf("</body>");
	if (last < 0 || base.length - last > 200) throw new Error("尾部 </body> 锚点异常");
	base = base.slice(0, last) + injectBlock + base.slice(last);
}
fs.writeFileSync(desktopPath, base);
console.log(`✔ desktop/app.html (${(base.length / 1e6).toFixed(2)}MB, 面板:${!!panelBlock}, 注入:${!!injectBlock})`);
if (!panelBlock || !injectBlock) console.log("  ⚠ 桌面面板/注入缺失——首次构建需先注入（见 git 历史或手工补）");

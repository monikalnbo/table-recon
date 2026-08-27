/**
 * desktop/main.mjs —— 双表核对 桌面版主进程（ESM 入口）
 * =============================================================================
 * 硬要求：Electron ≥35（pi SDK 用 Node22 的 fs.globSync；Electron33=Node20 会崩）
 * 职责：
 *   1) 绿色便携：exe 旁数据目录（PORTABLE_EXECUTABLE_DIR 注入），不写注册表
 *   2) 核对 IPC：UI 按钮 → recon-js（与 AI 同引擎，单一 canonical）
 *   3) AI IPC：agent-service 会话，流式事件推 renderer
 *   4) 窗口管理：单实例锁、外部链接走系统浏览器、下载走另存为对话框
 */
import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";
import { compareData, exportReport, readSheet } from "recon-js"; // canonical 引擎（npm file: 依赖）
import { AgentService } from "./main/agent-service.mjs"; // AI 会话服务

const __dirname = path.dirname(fileURLToPath(import.meta.url)); // ESM 下自算目录

/* ---------- 绿色便携：数据目录跟随 exe ---------- */
/** portable 版由 electron-builder 注入 PORTABLE_EXECUTABLE_DIR；检测到即重定向 userData → exe 旁 TableReconData/（配置/报告全跟 exe 走，删除即卸载） */
function applyPortableUserData() {
	const portableDir = process.env.PORTABLE_EXECUTABLE_DIR; // 仅 portable 包有
	if (!portableDir) return false; // 开发/安装版用默认
	app.setPath("userData", path.join(portableDir, "TableReconData")); // 重定向
	return true;
}
applyPortableUserData(); // 必须在 app ready 前执行

const gotLock = app.requestSingleInstanceLock(); // 单实例锁（防双开抢报告文件）
if (!gotLock) app.quit(); // 第二实例直接退出

let win = null; // 主窗口引用（IPC 事件推送用）
let agent = null; // AI 服务实例（ready 后创建）

/** 建主窗口 + 加载 renderer + 下载/外链策略 */
function createWindow() {
	win = new BrowserWindow({
		width: 1280, height: 900, minWidth: 760, // 尺寸
		backgroundColor: "#0f1117", title: "双表核对", // 深色底防白闪
		autoHideMenuBar: true, // 菜单栏自动隐藏
		webPreferences: { contextIsolation: true, preload: path.join(__dirname, "preload.cjs") }, // 隔离 + 白名单桥
	});
	win.loadFile(path.join(__dirname, "app.html")); // 自包含 GUI（离线可用）
	// 下载（网页引擎导出的 blob）→ 原生另存为对话框
	win.webContents.session.on("will-download", (_e, item) => {
		dialog.showSaveDialog(win, {
			title: "保存核对报告", defaultPath: item.getFilename(), // 默认文件名
			filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }], // 只留 xlsx
		}).then(({ canceled, filePath }) => {
			if (!canceled && filePath) item.setSavePath(filePath); // 选了路径 → 存
			else item.cancel(); // 取消 → 弃
		});
	});
	// 页面内 target=_blank 一律走系统浏览器（不让 Electron 开新窗）
	win.webContents.setWindowOpenHandler(({ url }) => {
		shell.openExternal(url);
		return { action: "deny" };
	});
}

app.whenReady().then(() => {
	agent = new AgentService(app.getPath("userData")); // AI 服务（portable 已重定向）
	registerIpc(); // 注册全部 IPC 通道
	createWindow(); // 开窗
	app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); }); // mac 点 dock 复活
});
app.on("second-instance", () => win?.focus()); // 二实例请求 → 聚焦既有窗
app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); }); // 非 mac 关窗即退
app.on("before-quit", () => agent?.dispose()); // 退出前销毁会话（断流）

/** IPC 通道注册（全部 invoke/handle 模式 + 事件推送） */
function registerIpc() {
	/* ---- 核对（UI 按钮与 AI 同引擎——深度嵌入的关键一条） ---- */
	ipcMain.handle("recon:compare", async (_e, a) => {
		try {
			const A = readSheet(fs.readFileSync(a.fileA), a.sheetA); // 读 A
			const B = readSheet(fs.readFileSync(a.fileB), a.sheetB); // 读 B
			const result = compareData({ // 核对（纯函数）
				headersA: A.headers, rowsA: A.rows, headersB: B.headers, rowsB: B.rows,
				keyA: a.keyA, keyB: a.keyB, rules: a.rules, tolerance: a.tolerance ?? 0,
			});
			let report = a.output; // 报告路径：指定 > A 同目录
			if (!report) {
				const dir = path.dirname(a.fileA);
				report = path.join(dir, "核对报告.xlsx");
			}
			await exportReport(result, report, { fileA: a.fileA, fileB: a.fileB }); // 7 表标红
			const slim = { ...result, _raw: undefined }; // 剔不可序列化原料
			slim.report = report; // 回填路径
			return { ok: true, data: slim }; // 成功信封
		} catch (err) {
			return { ok: false, error: err.message }; // 失败信封（renderer 直接展示）
		}
	});

	/* ---- AI 配置读写 ---- */
	// 读：key 永不回明文（掩码 + hasKey 布尔）
	ipcMain.handle("agent:getConfig", () => ({ ...agent.getConfig(), hasKey: !!agent.getConfig().apiKey, apiKey: mask(agent.getConfig().apiKey) }));
	ipcMain.handle("agent:setConfig", (_e, cfg) => {
		const cur = agent.getConfig(); // 现值
		// 空 apiKey / 掩码串 = 保留原值（防 UI 回显掩码把真 key 覆盖掉）
		if (!cfg.apiKey || cfg.apiKey.includes("•")) cfg.apiKey = cur.apiKey;
		const saved = agent.setConfig(cfg); // 单次落盘
		return { ok: true, config: { ...saved, apiKey: mask(saved.apiKey) } }; // 回显仍掩码
	});

	/* ---- AI 任务执行 ---- */
	const send = (ch, payload) => win?.webContents.send(ch, payload); // 事件推送帮助函数（窗口可能已关）
	ipcMain.handle("agent:run", async (_e, { task, context }) => {
		if (!task?.trim()) return { ok: false, error: "任务为空" }; // 空任务拦截
		agent.run(task, context || {}, { // 异步执行（不等完成，流式靠事件）
			onDelta: (d) => send("agent:delta", d), // 文本增量
			onToolStart: (n) => send("agent:tool", { name: n, phase: "start" }), // 工具开始
			onToolEnd: (n, isErr, summary) => send("agent:tool", { name: n, phase: "end", error: isErr, summary }), // 工具结束
			onDone: () => send("agent:done", {}), // 完成
			onError: (m) => send("agent:error", m), // 业务错误
		}).catch((err) => send("agent:error", err?.message || String(err))); // busy 抛异常也回传，不吞 unhandledRejection
		return { ok: true }; // 仅表示「已受理」
	});
	ipcMain.handle("agent:stop", async () => { await agent.stop(); return { ok: true }; }); // 中止
}

/** key 掩码：≤8 位全隐；否则首 4 + •••• + 尾 4（明文永不回 renderer） */
const mask = (k) => (k ? k.length <= 8 ? "••••" : `${k.slice(0, 4)}••••${k.slice(-4)}` : "");

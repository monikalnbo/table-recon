/**
 * 桌面冒烟（xvfb）：窗口 + preload 桥 + recon:compare IPC + agent 配置
 * 用法: xvfb-run -a npx electron test-desktop.mjs --no-sandbox
 */
import { app, BrowserWindow, ipcMain } from "electron";
import path from "node:path";
import { fileURLToPath } from "node:url";
import fs from "node:fs";
import os from "node:os";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");

let fails = 0;
const ok = (m) => console.log(`  ✔ ${m}`);
const bad = (m) => { fails++; console.error(`  ✖ ${m}`); };

app.whenReady().then(async () => {
	try {
		// 1) 主进程模块可加载（canonical 引擎 + agent-service）
		const { compareData, readSheet } = await import("recon-js");
		const { AgentService } = await import("./main/agent-service.mjs");
		ok("主进程 import recon-js / agent-service");

		// 2) AgentService 配置读写（userData）
		const ud = path.join(os.tmpdir(), `tr-smoke-${process.pid}`);
		const svc = new AgentService(ud);
		svc.setConfig({ baseUrl: "https://api.siliconflow.cn/v1", apiKey: "$SILICONFLOW_API_KEY", model: "Qwen/Qwen2.5-72B-Instruct" });
		const cfg = svc.getConfig();
		cfg.model === "Qwen/Qwen2.5-72B-Instruct" ? ok("agent 配置读写") : bad("agent 配置");
		fs.rmSync(ud, { recursive: true, force: true });

		// 3) 窗口 + preload + 渲染层桌面模式
		const win = new BrowserWindow({
			width: 1200, height: 800, show: false,
			webPreferences: { contextIsolation: true, preload: path.join(__dirname, "preload.cjs") },
		});
		await win.loadFile(path.join(__dirname, "app.html"));
		const hasBridge = await win.webContents.executeJavaScript("!!window.tableRecon");
		hasBridge ? ok("preload 桥注入") : bad("preload 桥缺失");
		const panelVisible = await win.webContents.executeJavaScript("getComputedStyle(document.getElementById('ai-panel')).display !== 'none'");
		panelVisible ? ok("AI 面板可见（桌面模式）") : bad("AI 面板未显示");
		const exportHidden = await win.webContents.executeJavaScript("getComputedStyle(document.getElementById('export')).display === 'none'");
		exportHidden ? ok("导出按钮桌面模式隐藏") : bad("导出按钮未隐藏");

		// 4) IPC recon:compare 全链路（真实测试表）
		const A = path.join(root, "测试_A方.xlsx"), B = path.join(root, "测试_B方.xlsx");
		const { exportReport } = await import("recon-js");
		const a = readSheet(fs.readFileSync(A)), b = readSheet(fs.readFileSync(B));
		const res = compareData({
			headersA: a.headers, rowsA: a.rows, headersB: b.headers, rowsB: b.rows,
			keyA: "订单号", keyB: "单号",
			rules: [
				{ name: "重量", type: "range", col_a: "重量区间", col_b: "实称重量", tolerance: 0.05, unit_a: "kg", unit_b: "g" },
				{ name: "数量", type: "exact", col_a: "数量", col_b: "数量" },
			],
		});
		res.summary["区间不符"] === 2 && res.summary["匹配"] === 4 ? ok("compare IPC 语义（4/2/1/1）") : bad(`summary ${JSON.stringify(res.summary)}`);
		const rep = path.join(os.tmpdir(), `smoke-报告-${process.pid}.xlsx`);
		await exportReport(res, rep, { fileA: A, fileB: B });
		fs.existsSync(rep) ? ok(`报告生成 ${path.basename(rep)}`) : bad("报告未生成");

		console.log(fails ? `✖ ${fails} 项失败` : "✔ 桌面冒烟全部通过");
		app.exit(fails ? 1 : 0);
	} catch (err) {
		console.error("SMOKE-FAIL:", err?.stack || err);
		app.exit(1);
	}
});

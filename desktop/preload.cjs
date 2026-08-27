/**
 * desktop/preload.cjs —— contextBridge 白名单桥（必须是 CJS：sandboxed preload 限制）
 * =============================================================================
 * 原则：只暴露最小面；apiKey 永不经过 renderer（main 侧掩码）；
 *      File.path 在 Electron 33+ 已移除 → 用 webUtils.getPathForFile 取路径。
 */
const { contextBridge, ipcRenderer, webUtils } = require("electron"); // Electron API

contextBridge.exposeInMainWorld("tableRecon", {
	// Electron 33+ File.path 已移除，经 webUtils 取真实路径（拖拽/选择的文件）
	getPathForFile: (file) => {
		try { return webUtils.getPathForFile(file); } catch { return ""; } // 拿不到给空串（上游兜底 file.name）
	},

	// 核对（与 AI 同引擎，深度嵌入）：{fileA,fileB,keyA,keyB,rules,...} → {ok,data|error}
	reconCompare: (args) => ipcRenderer.invoke("recon:compare", args),

	// AI 配置：get 返回掩码 key；set 空 key = 保留原值
	agentGetConfig: () => ipcRenderer.invoke("agent:getConfig"),
	agentSetConfig: (cfg) => ipcRenderer.invoke("agent:setConfig", cfg),

	// AI 任务：run 立即返回（受理）；过程靠下方事件流
	agentRun: (task, context) => ipcRenderer.invoke("agent:run", { task, context }),
	agentStop: () => ipcRenderer.invoke("agent:stop"),

	// 事件订阅（main → renderer 推送）
	onDelta: (cb) => ipcRenderer.on("agent:delta", (_e, d) => cb(d)), // 流式文本增量
	onTool: (cb) => ipcRenderer.on("agent:tool", (_e, t) => cb(t)), // 工具开始/结束
	onDone: (cb) => ipcRenderer.on("agent:done", () => cb()), // 完成
	onAgentError: (cb) => ipcRenderer.on("agent:error", (_e, m) => cb(m)), // 错误

	isDesktop: true, // 桌面模式哨兵（app.html 据此启用 AI 面板/IPC 核对）
});

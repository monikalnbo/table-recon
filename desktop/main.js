const { app, BrowserWindow, dialog } = require("electron");
const path = require("path");

// 双表核对 · 桌面版：Electron 壳加载自包含 app.html（离线，数据不出本机）
app.whenReady().then(() => {
	const win = new BrowserWindow({
		width: 1280,
		height: 880,
		minWidth: 760,
		backgroundColor: "#0f1117",
		title: "双表核对",
		autoHideMenuBar: true,
		webPreferences: { contextIsolation: true },
	});
	win.loadFile(path.join(__dirname, "app.html"));

	// 报告导出：拦截网页下载 → 弹原生「另存为」对话框
	win.webContents.session.on("will-download", (_event, item) => {
		dialog
			.showSaveDialog(win, {
				title: "保存核对报告",
				defaultPath: item.getFilename(),
				filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
			})
			.then(({ canceled, filePath }) => {
				if (!canceled && filePath) item.setSavePath(filePath);
				else item.cancel();
			});
	});
});

app.on("window-all-closed", () => app.quit());

# Desktop · 双表核对 桌面版（Windows / macOS / Linux）

Electron 壳加载自包含 `app.html`（依赖已全部内联，**完全离线**，表格数据不出本机）。
功能与网页/iOS 版完全一致；导出报告走**原生「另存为」对话框**。

## 安装包获取（无需本地环境）

push 后 GitHub Actions 自动在三平台云机打包，到
[Actions](https://github.com/monikalnbo/table-recon/actions) → **Build Desktop App** → 最新运行 → Artifacts 下载：

| Artifact | 内容 | 安装 |
|---|---|---|
| `TableRecon-desktop-windows-latest` | `TableRecon-1.0.0-win-x64.exe`（NSIS 安装器） | 双击安装；SmartScreen 提示选「更多信息→仍要运行」 |
| `TableRecon-desktop-ubuntu-latest` | `TableRecon-1.0.0-mac-arm64.zip` / `mac-x64.zip` + Linux AppImage | Mac：解压出 `.app`，右键→打开；Intel 用 x64，Apple 芯片用 arm64。Linux：`chmod +x *.AppImage && ./…` |

> 均为未签名包（个人使用足够）；需要正式分发可自行接入证书。

## 本地运行 / 打包

```bash
cd desktop
npm install
npm start            # 开发运行
npx electron-builder # 本机平台打包，产物在 dist/
```

## 结构

```
desktop/
├── main.js       # Electron 主进程：窗口 + 导出另存为对话框
├── app.html      # 自包含核对界面（与 iOS 版同一份）
└── package.json  # electron-builder 配置（win nsis / mac dmg / linux AppImage）
```

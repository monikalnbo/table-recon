# iOS · 双表核对 App

SwiftUI + WKWebView 壳，内嵌自包含 `app.html`（SheetJS + ExcelJS 已内联，**完全离线**，表格数据不离开手机）。
功能与网页版一致：导入两表 → 选关联列 → 配区间/精确规则 → 核对 → 导出标红报告（走系统分享面板，可存到「文件」）。

## 三种安装方式

| 方式 | 需要 | 说明 |
|---|---|---|
| **A. 云端打包 ipa**（推荐） | 无 Mac | push 后 GitHub Actions 在 macOS 云机上产出 `TableRecon.ipa`（未签名）。Actions 页 → Build iOS App → Artifacts 下载。用 [Sideloadly](https://sideloadly.io/)（Win/Mac）或 AltStore 用你自己的 Apple ID 签名安装；免费账号 7 天重签一次 |
| **B. Xcode 编译** | Mac + Xcode 16 | `brew install xcodegen && cd ios && xcodegen generate`，打开 `TableRecon.xcodeproj`，选自己的 Team（免费个人 Team 即可）直接真机运行 |
| **C. Swift Playgrounds** | iPad/iPhone | 无需 Mac、免签名：把 `../TableRecon.swiftpm` 整个文件夹 AirDrop 到设备，Swift Playgrounds 打开运行 |

> App Store 上架需 Apple Developer Program（$99/年）+ 正式签名 + 审核；本仓库默认产出未签名包，签名自由掌握。
> 若配置了正式证书，可在 Actions Secrets 里加签名变量改造 workflow 走 TestFlight。

## 结构

```
ios/
├── project.yml          # XcodeGen 工程定义（CI 上生成 .xcodeproj）
└── TableRecon/
    ├── App.swift        # SwiftUI 入口
    ├── ContentView.swift# WKWebView + 导出桥(messageHandlers.export → 分享面板)
    └── app.html         # 自包含网页（离线，含 iOS safe-area 适配）
```

## 本地验证 HTML（无需 Mac）

任何浏览器打开 `TableRecon/app.html` 即可调试界面与导出逻辑（浏览器里走下载，App 里走分享）。

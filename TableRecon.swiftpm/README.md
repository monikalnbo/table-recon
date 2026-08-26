# TableRecon.swiftpm · Swift Playgrounds 版

**免 Mac、免签名、免开发者账号**——iPhone/iPad 装了免费的 [Swift Playgrounds](https://apps.apple.com/app/swift-playgrounds/id908519492) 就能直接运行的本 App。

## 安装步骤

1. 把整个 `TableRecon.swiftpm` 文件夹弄到 iPhone/iPad 上（AirDrop / iCloud 云盘 /「文件」App 均可）
2. 在「文件」里长按该文件夹 → 共享 → 用 **Swift Playgrounds** 打开
   （或直接在 Swift Playgrounds 里点 + → 从文件导入）
3. 点 ▶ 运行，即得到完整的双表核对 App

> 文件夹内含 `app.html`（1.6MB，内联全部依赖）——必须和 Swift 文件一起拷贝，不能只拷 .swift。

## 代码说明

- `Package.swift`：Swift Playgrounds App 格式声明
- `MyApp.swift`：SwiftUI 壳 + WKWebView，加载 `app.html`
- `app.html`：核对界面本体（导出走系统分享面板）

import SwiftUI
import WebKit

@main
struct TableReconApp: App {
	var body: some Scene {
		WindowGroup { ContentView() }
	}
}

/// WKWebView 承载同目录 app.html（内联 SheetJS + ExcelJS，完全离线）
struct ContentView: UIViewRepresentable {
	func makeUIView(context: Context) -> WKWebView {
		let config = WKWebViewConfiguration()
		config.userContentController.add(context.coordinator, name: "export")
		let webView = WKWebView(frame: .zero, configuration: config)
		webView.uiDelegate = context.coordinator
		webView.isOpaque = false
		webView.backgroundColor = UIColor(red: 0.06, green: 0.07, blue: 0.09, alpha: 1)
		if let url = Bundle.main.url(forResource: "app", withExtension: "html") {
			webView.loadFileURL(url, allowingReadAccessTo: url.deletingLastPathComponent())
		}
		return webView
	}

	func updateUIView(_ webView: WKWebView, context: Context) {}
	func makeCoordinator() -> Coordinator { Coordinator() }

	final class Coordinator: NSObject, WKScriptMessageHandler, WKUIDelegate {
		func userContentController(_ userContentController: WKUserContentController,
		                           didReceive message: WKScriptMessage) {
			guard message.name == "export",
			      let body = message.body as? [String: Any],
			      let name = body["name"] as? String,
			      let base64 = body["base64"] as? String,
			      let data = Data(base64Encoded: base64) else { return }
			let url = FileManager.default.temporaryDirectory.appendingPathComponent(name)
			try? data.write(to: url)
			DispatchQueue.main.async { self.share(url: url) }
		}

		private func share(url: URL) {
			guard let scene = UIApplication.shared.connectedScenes
				.compactMap({ $0 as? UIWindowScene }).first,
			      let root = scene.keyWindow?.rootViewController
				?? scene.windows.first?.rootViewController else { return }
			let picker = UIActivityViewController(activityItems: [url], applicationActivities: nil)
			var top = root
			while let presented = top.presentedViewController { top = presented }
			top.present(picker, animated: true)
		}
	}
}

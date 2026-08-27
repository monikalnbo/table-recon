/** Spike 1: pi SDK 在 Electron 主进程跑通（无窗口，流式输出到 stdout） */
import { app } from "electron";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

app.whenReady().then(async () => {
	try {
		const { createAgentSession, ModelRuntime } = await import("@earendil-works/pi-coding-agent");
		const cfg = JSON.parse(fs.readFileSync(path.join(root, "agent", "config.json"), "utf8"));
		const pid = cfg.provider || "custom-llm";
		const modelsJson = {
			providers: {
				[pid]: {
					name: "Spike",
					baseUrl: cfg.baseUrl,
					apiKey: cfg.apiKey,
					api: cfg.api || "openai-completions",
					models: [{
						id: cfg.model, name: cfg.model, reasoning: false, input: ["text"],
						contextWindow: cfg.contextWindow || 32768, maxTokens: 2048,
						cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
					}],
				},
			},
		};
		const modelsPath = path.join(app.getPath("userData"), "models.json");
		fs.mkdirSync(path.dirname(modelsPath), { recursive: true });
		fs.writeFileSync(modelsPath, JSON.stringify(modelsJson));

		const t0 = Date.now();
		const rt = await ModelRuntime.create({ modelsPath });
		const model = rt.getModel(pid, cfg.model);
		if (!model) throw new Error("模型注册失败");
		console.log(`[spike] ModelRuntime ok (${Date.now() - t0}ms): ${pid}/${cfg.model}`);

		const { session } = await createAgentSession({ model, modelRuntime: rt });
		let saw = "";
		session.subscribe((e) => {
			if (e.type === "message_update" && e.assistantMessageEvent.type === "text_delta") {
				saw += e.assistantMessageEvent.delta;
				process.stdout.write(e.assistantMessageEvent.delta);
			}
		});
		await session.prompt("只回复两个字符:OK");
		session.dispose();
		console.log(`\n[spike] prompt ok, reply=${JSON.stringify(saw.trim())}`);
		console.log("SPIKE1-PASS");
		app.exit(0);
	} catch (err) {
		console.error("SPIKE1-FAIL:", err?.stack || err);
		app.exit(1);
	}
});

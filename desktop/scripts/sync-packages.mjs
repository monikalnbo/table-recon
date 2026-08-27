#!/usr/bin/env node
/**
 * desktop/scripts/sync-packages.mjs —— 打包前置：canonical → app-src 同步副本
 * =============================================================================
 * 为什么需要：desktop 依赖 file:../recon-js（软链出项目）。干净克隆/CI 下，
 * Node 沿 realpath 向上找 node_modules 不经过 desktop/，xlsx/exceljs 解析会断。
 * 方案：同步为 desktop/app-src/ 实体目录（package.json 指向 file:./app-src/…），
 *      依赖链全部落在 desktop/node_modules（npm 提升），electron-builder 走标准路径。
 * 开发纪律：改了 canonical（recon-js / excel-mcp-js）必须重跑本脚本再测 desktop！
 * 用法: node scripts/sync-packages.mjs   （desktop/npm install 之前跑一次）
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const desktop = path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."); // desktop/
const repo = path.resolve(desktop, ".."); // 仓库根
const appSrc = path.join(desktop, "app-src"); // 同步目标（已 gitignore）

const PKGS = ["recon-js", "excel-mcp-js"]; // canonical 包清单

fs.rmSync(appSrc, { recursive: true, force: true }); // 全量重建（防残留旧文件）
for (const name of PKGS) {
	const src = path.join(repo, name); // 源（canonical）
	const dst = path.join(appSrc, name); // 目标
	fs.mkdirSync(dst, { recursive: true });
	for (const f of fs.readdirSync(src)) { // 逐项拷贝
		if (f === "node_modules" || f === ".git") continue; // 排除重目录
		fs.cpSync(path.join(src, f), path.join(dst, f), { recursive: true }); // 递归拷
	}
	console.log(`✔ ${name} → app-src/${name}`);
}
console.log("app-src 同步完成（package.json 依赖指向 file:./app-src/…）");

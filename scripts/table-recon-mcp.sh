#!/usr/bin/env bash
# 启动本仓库 table-recon MCP（业务核对）。路径相对本脚本，克隆后不用改。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
export TABLE_RECON_HOME="$ROOT"
exec python3 "$ROOT/03-mcp-server/server.py"

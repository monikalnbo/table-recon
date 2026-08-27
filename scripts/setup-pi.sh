#!/usr/bin/env bash
# 检查 pi + 内置 excel MCP + table-recon MCP 能否在本机跑起来
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"

ok() { printf '  ✔ %s\n' "$1"; }
bad() { printf '  ✖ %s\n' "$1"; exit 1; }

echo "table-recon pi 集成检查"
echo "  root: $ROOT"

command -v python3 >/dev/null || bad "需要 python3"
python3 -c "import openpyxl" 2>/dev/null || bad "需要 openpyxl：pip install -r requirements.txt"
ok "openpyxl"

[[ -f "$ROOT/excel-mcp/server.py" ]] || bad "缺少 excel-mcp/server.py"
python3 "$ROOT/excel-mcp/server.py" --selftest >/dev/null || bad "excel MCP 自测失败"
ok "excel MCP（13 工具，自测通过）"

[[ -f "$ROOT/03-mcp-server/server.py" ]] || bad "缺少 03-mcp-server/server.py"
ok "table-recon MCP"

[[ -f "$ROOT/.pi/skills/table-recon/SKILL.md" ]] || bad "缺少 skill"
ok "skill table-recon"

[[ -f "$ROOT/.mcp.json" ]] || bad "缺少 .mcp.json"
ok ".mcp.json（excel + table-recon）"

command -v node >/dev/null && ok "node $(node --version)" || printf '  · node 未装（pi agent 需要 Node 20+）\n'
command -v pi >/dev/null && ok "pi $(pi --version 2>/dev/null | head -1)" || printf '  · pi CLI 未在 PATH（SDK 仍可用）\n'

echo
echo "下一步："
echo "  1) cd agent && cp config.example.json config.json   # 填 baseUrl / apiKey / model"
echo "  2) cd agent && npm install && node excel-agent.mjs"

#!/usr/bin/env bash
# 检查 pi + excel MCP + table-recon MCP 能否在本机跑起来
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
export PATH="${HOME}/.local/bin:${HOME}/.cargo/bin:${PATH}"

ok() { printf '  ✔ %s\n' "$1"; }
bad() { printf '  ✖ %s\n' "$1"; exit 1; }

echo "table-recon pi 集成检查"
echo "  root: $ROOT"

command -v python3 >/dev/null || bad "需要 python3"
python3 -c "import openpyxl" 2>/dev/null || bad "需要 openpyxl：pip install openpyxl"
ok "openpyxl"

command -v uvx >/dev/null || bad "需要 uvx：curl -LsSf https://astral.sh/uv/install.sh | sh"
ok "$(uv --version 2>/dev/null || echo uvx present)"

uvx excel-mcp-server --help >/dev/null || bad "uvx excel-mcp-server 无法启动"
ok "excel-mcp-server (haris-musa)"

[[ -f "$ROOT/03-mcp-server/server.py" ]] || bad "缺少 03-mcp-server/server.py"
ok "table-recon MCP"

[[ -f "$ROOT/.pi/skills/table-recon/SKILL.md" ]] || bad "缺少 skill"
ok "skill table-recon"

[[ -f "$ROOT/.mcp.json" ]] || bad "缺少 .mcp.json"
ok ".mcp.json"

command -v pi >/dev/null && ok "pi $(pi --version 2>/dev/null | head -1)" || printf '  · pi CLI 未在 PATH（SDK 仍可用）\n'

echo
echo "下一步："
echo "  1) 把 .mcp.json 里的两个 server 合并进 ~/.pi/agent/mcp.json"
echo "  2) cd 01-minimal-pi && npm install && node recon-agent.mjs"

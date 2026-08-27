#!/usr/bin/env bash
# 启动 haris-musa/excel-mcp-server（stdio）。需要 uvx：curl -LsSf https://astral.sh/uv/install.sh | sh
set -euo pipefail
export PATH="${HOME}/.local/bin:${HOME}/.cargo/bin:${PATH}"
if ! command -v uvx >/dev/null 2>&1; then
  echo "excel MCP 需要 uvx。安装: curl -LsSf https://astral.sh/uv/install.sh | sh" >&2
  exit 1
fi
exec uvx excel-mcp-server stdio "$@"

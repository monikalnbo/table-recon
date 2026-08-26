#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
table-recon MCP Server（stdio）
==============================

把 core/recon.py 的双表核对能力暴露成 MCP 工具，供 AI Agent（pi / Claude / Cursor…）调用。

工具:
  1. inspect_sheet   —— 读表头+行数+前几行预览，帮 AI 选列
  2. compare_tables  —— 双表核对主工具：关联列 + 规则列表，返回异常清单并生成标红报告

配置（pi 的 ~/.pi/agent/mcp.json）:
  {
    "mcpServers": {
      "table-recon": {
        "command": "python3",
        "args": ["/绝对路径/table-toolkit/03-mcp-server/server.py"],
        "env": { "TABLE_RECON_HOME": "/绝对路径/table-toolkit" }
      }
    }
  }

  （不设 TABLE_RECON_HOME 时，按 server.py 所在目录的上一级查找 core/recon.py）

协议: MCP stdio (JSON-RPC 2.0, 2024-11-05)。仅依赖 openpyxl。
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

HOME = Path(os.environ.get("TABLE_RECON_HOME") or Path(__file__).resolve().parent.parent)
sys.path.insert(0, str(HOME / "core"))

import recon  # noqa: E402

PROTOCOL = "2024-11-05"
SERVER_INFO = {"name": "table-recon", "version": "1.0.0"}
CAPABILITIES = {"tools": {}}

TOOLS_SPEC = [
    {
        "name": "inspect_sheet",
        "description": "读取 Excel/CSV 表格结构：返回列名、总行数、前 5 行预览。用于核对前选列。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "表格文件路径（.xlsx/.xls/.csv）"},
                "sheet": {"type": "string", "description": "工作表名，缺省为第一个"},
            },
            "required": ["path"],
        },
    },
    {
        "name": "compare_tables",
        "description": (
            "双表关联核对：以指定列（如订单号）为关联键，逐条执行核对规则，输出三类异常"
            "（区间不符/项不一致/单边缺失）并生成标红 Excel 报告。"
            "规则两种: type=range —— A方该列是范围(如 0.5-1kg、500-800g、小于1)，B方该列是具体值，"
            "B值不在A范围内即不符（单位自动换算，kg/g/公斤/吨/lb，裸数字默认kg，可用 unit_a/unit_b 指定）; "
            "type=exact —— 两列必须完全一致(数值按值比,文本归一化比)。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_a": {"type": "string", "description": "A方表格路径（区间规则中 A=范围一方）"},
                "file_b": {"type": "string", "description": "B方表格路径（区间规则中 B=具体值一方）"},
                "key_a": {"type": "string", "description": "A方关联列名，如 订单号"},
                "key_b": {"type": "string", "description": "B方关联列名"},
                "rules": {
                    "type": "array",
                    "description": "核对规则列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "规则显示名，如 重量"},
                            "type": {"type": "string", "enum": ["range", "exact"]},
                            "col_a": {"type": "string", "description": "A方列名"},
                            "col_b": {"type": "string", "description": "B方列名"},
                            "tolerance": {"type": "number", "description": "容差(基准单位kg)，如0.05=±50g，仅range"},
                            "unit_a": {"type": "string", "description": "A方裸数字的默认单位，如 kg/g/公斤/吨/lb，缺省kg，仅range"},
                            "unit_b": {"type": "string", "description": "B方裸数字的默认单位，缺省kg，仅range"},
                        },
                        "required": ["type", "col_a", "col_b"],
                    },
                },
                "sheet_a": {"type": "string", "description": "A方工作表名，缺省第一个"},
                "sheet_b": {"type": "string", "description": "B方工作表名，缺省第一个"},
                "tolerance": {"type": "number", "description": "全局容差，规则未指定时生效"},
                "output": {"type": "string", "description": "报告输出路径，缺省 核对报告.xlsx（写在A方文件同目录或工作目录）"},
            },
            "required": ["file_a", "file_b", "key_a", "key_b", "rules"],
        },
    },
]


# ---------------------------------------------------------------- 工具实现

def _read_any(path: str, sheet: str | None):
    """xlsx/xls 直接读；csv 转 Excel 内存形式读"""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"文件不存在: {p}")
    if p.suffix.lower() == ".csv":
        import csv
        with open(p, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        headers = [str(h).strip() if h is not None else "" for h in (rows[0] if rows else [])]
        data = [dict(zip(headers, r)) for r in rows[1:] if any(str(c).strip() for c in r)]
        return headers, data
    return recon.read_sheet(p, sheet)


def tool_inspect_sheet(args: dict) -> str:
    headers, data = _read_any(args["path"], args.get("sheet"))
    preview = [{k: recon._plain(v) for k, v in r.items()} for r in data[:5]]
    return json.dumps(
        {"path": args["path"], "columns": headers, "row_count": len(data), "preview": preview},
        ensure_ascii=False, indent=2,
    )


def tool_compare_tables(args: dict) -> str:
    out = args.get("output")
    if not out:
        dir_a = Path(args["file_a"]).parent
        out = str(dir_a / "核对报告.xlsx") if dir_a.is_dir() else "核对报告.xlsx"
    result = recon.compare(
        args["file_a"], args["file_b"],
        key_a=args["key_a"], key_b=args["key_b"],
        rules=args["rules"], tolerance=args.get("tolerance", 0.0),
        sheet_a=args.get("sheet_a"), sheet_b=args.get("sheet_b"),
    )
    recon.export_report(result, out)
    slim = {k: v for k, v in result.items() if not k.startswith("_")}
    slim["report"] = str(Path(out).resolve())
    return json.dumps(slim, ensure_ascii=False, indent=2)


TOOL_IMPLS = {"inspect_sheet": tool_inspect_sheet, "compare_tables": tool_compare_tables}


# ---------------------------------------------------------------- MCP stdio 协议

def send(msg: dict):
    sys.stdout.write(json.dumps(msg, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def reply(id_, result):
    send({"jsonrpc": "2.0", "id": id_, "result": result})


def reply_err(id_, code, message):
    send({"jsonrpc": "2.0", "id": id_, "error": {"code": code, "message": message}})


def handle(req: dict):
    method = req.get("method")
    id_ = req.get("id")

    if method == "initialize":
        reply(id_, {"protocolVersion": PROTOCOL, "capabilities": CAPABILITIES, "serverInfo": SERVER_INFO})
    elif method == "notifications/initialized":
        pass  # 通知无响应
    elif method == "ping":
        reply(id_, {})
    elif method == "tools/list":
        reply(id_, {"tools": TOOLS_SPEC})
    elif method == "tools/call":
        name = req["params"]["name"]
        if name not in TOOL_IMPLS:
            reply_err(id_, -32602, f"未知工具: {name}")
            return
        try:
            text = TOOL_IMPLS[name](req["params"].get("arguments") or {})
            reply(id_, {"content": [{"type": "text", "text": text}], "isError": False})
        except Exception as e:  # noqa: BLE001 —— 把业务错误回传给 AI 而不是崩掉
            reply(id_, {"content": [{"type": "text", "text": f"错误: {e}"}], "isError": True})
    elif id_ is not None:
        reply_err(id_, -32601, f"未实现: {method}")


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            handle(json.loads(line))
        except json.JSONDecodeError:
            continue  # 坏行丢弃


if __name__ == "__main__":
    main()

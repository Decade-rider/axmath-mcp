"""Print this checkout's MCP JSON for CC Switch or another local stdio client."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def build_config(root: Path, *, codex: bool = False) -> dict:
    python = root.resolve() / ".venv" / "Scripts" / "python.exe"
    if not python.is_file():
        raise ValueError("Project environment missing. Run uv sync in this checkout first.")
    server = {
        "command": str(python),
        "args": ["-m", "axmath_mcp.server"],
        "env": {"PYTHONUTF8": "1", "PYTHONUNBUFFERED": "1"},
    }
    if codex:
        server.update(startup_timeout_sec=30, tool_timeout_sec=300)
    return {"mcpServers": {"axmath": server}}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--codex",
        action="store_true",
        help="Include Codex-specific timeout fields; sync to Codex only.",
    )
    args = parser.parse_args()
    if sys.platform != "win32":
        parser.exit(1, "AxMath MCP requires Windows desktop Office and AxMath.\n")
    try:
        config = build_config(ROOT, codex=args.codex)
    except ValueError as exc:
        parser.exit(1, str(exc) + "\n")
    # ASCII escapes survive both Windows PowerShell 5 and PowerShell 7 pipelines,
    # including checkouts whose real absolute path contains non-ASCII characters.
    print(json.dumps(config, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()

"""Preview or install this Windows checkout's AxMath MCP entry in Codex."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile
import tomllib
from collections.abc import Mapping
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]


def build_server_config(project_root: Path, uv_command: str) -> dict[str, Any]:
    """Build an entry using this checkout and the installed package version."""
    from axmath_mcp import __version__

    project_root = Path(project_root)
    if not project_root.is_absolute():
        raise ValueError("The project directory must be an absolute path.")
    if not uv_command:
        raise ValueError("An installed uv executable is required.")
    return {
        "command": uv_command,
        "args": [
            "--directory",
            str(project_root),
            "run",
            "--no-sync",
            "python",
            "-m",
            "axmath_mcp.server",
        ],
        "startup_timeout_sec": 30,
        "tool_timeout_sec": 300,
        "env": {
            "UV_CACHE_DIR": str(project_root / ".uv-cache"),
            "PYTHONUTF8": "1",
            "PYTHONUNBUFFERED": "1",
            "AXMATH_MCP_VERSION": __version__,
        },
    }


def render_config_snippet(entry: Mapping[str, Any]) -> str:
    """Render only the new TOML tables, never the existing private config."""
    lines = ["[mcp_servers.axmath]"]
    for key in ("command", "args", "startup_timeout_sec", "tool_timeout_sec"):
        lines.append(f"{key} = {json.dumps(entry[key], ensure_ascii=False)}")
    lines.extend(["", "[mcp_servers.axmath.env]"])
    lines.extend(
        f"{key} = {json.dumps(value, ensure_ascii=False)}" for key, value in entry["env"].items()
    )
    return "\n".join(lines) + "\n"


def resolve_config_directory(
    environ: Mapping[str, str] | None = None, *, home: Path | None = None
) -> Path:
    """Respect an explicit absolute CODEX_HOME; otherwise use the user's home."""
    environ = os.environ if environ is None else environ
    directory = (
        Path(environ["CODEX_HOME"])
        if "CODEX_HOME" in environ
        else ((Path.home() if home is None else home) / ".codex")
    )
    if not directory.is_absolute():
        raise ValueError("CODEX_HOME must be an absolute directory path, not empty or relative.")
    return directory


def _read_state(path: Path) -> tuple[bool, bytes]:
    try:
        return True, path.read_bytes()
    except FileNotFoundError:
        return False, b""


def install_config(path: Path, entry: Mapping[str, Any]) -> dict[str, Any]:
    """Append one server, preserving existing bytes; never replace a conflict.

    Repeating an identical installation is a no-op. A final state comparison
    detects concurrent changes, including creation or deletion of config.toml.
    """
    path = Path(path)
    if not path.is_absolute():
        raise ValueError("The Codex config path must be absolute.")
    original_state = _read_state(path)
    existed, original = original_state
    existing = tomllib.loads(original.decode("utf-8-sig"))
    servers = existing.get("mcp_servers", {})
    if not isinstance(servers, dict):
        raise TypeError("Existing mcp_servers is not a TOML table; no configuration was changed.")
    if "axmath" in servers:
        if servers["axmath"] == entry:
            return {
                "installed": "axmath",
                "status": "already_configured",
                "changed": False,
                "config": str(path),
                "backup": None,
            }
        raise ValueError(
            "A different axmath entry already exists; it was left unchanged. "
            "Run this script without --install to preview the desired entry, then merge it manually."
        )
    snippet = render_config_snippet(entry)
    combined = original + (b"\n\n" if original else b"") + snippet.encode("utf-8")
    parsed = tomllib.loads(combined.decode("utf-8-sig"))
    if parsed["mcp_servers"]["axmath"] != entry:
        raise ValueError("The generated configuration did not round-trip; no file was changed.")
    path.parent.mkdir(parents=True, exist_ok=True)
    backup = None
    if existed:
        stamp = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%dT%H%M%S%f")
        backup = path.with_name(path.name + ".axmath-backup-" + stamp)
        shutil.copy2(path, backup)
    with tempfile.NamedTemporaryFile(
        dir=path.parent, prefix="axmath-config-", delete=False
    ) as temp:
        temporary = Path(temp.name)
        temp.write(combined)
    try:
        if _read_state(path) != original_state:
            raise RuntimeError("Codex config changed concurrently. No configuration was replaced.")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    return {
        "installed": "axmath",
        "status": "installed",
        "changed": True,
        "config": str(path),
        "backup": str(backup) if backup else None,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--install", action="store_true", help="Install the entry; otherwise print a preview only."
    )
    args = parser.parse_args()
    if sys.platform != "win32":
        parser.exit(
            1, "AxMath MCP requires Windows desktop Word and AxMath; run this on Windows.\n"
        )
    uv = shutil.which("uv")
    if not uv:
        parser.exit(1, "uv was not found on PATH. Install uv, then run uv sync in this checkout.\n")
    if not (ROOT / ".venv/Scripts/python.exe").is_file():
        parser.exit(1, "The project environment is missing. Run uv sync in this checkout first.\n")
    try:
        entry = build_server_config(ROOT, uv)
        if not args.install:
            print(render_config_snippet(entry))
            return
        result = install_config(resolve_config_directory() / "config.toml", entry)
    except ModuleNotFoundError:
        parser.exit(
            1,
            "The project package is not installed. Run uv sync, then uv run python scripts/register_codex.py.\n",
        )
    except (OSError, TypeError, ValueError, RuntimeError) as exc:
        parser.exit(1, str(exc) + "\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()

"""Exercise the real stdio MCP transport, optionally using a new Word document.

Default: initialize, list tools, inspect health (no document mutation).
Opt in: ``uv run python scripts/smoke_mcp.py --exercise-word``.
The saved smoke document remains open; this script never closes or quits Word.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import tempfile
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parents[1]
PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"
EXPECTED_TOOLS = {
    "axmath_health",
    "word_list_documents",
    "word_create_document",
    "word_open_document",
    "word_document_info",
    "word_list_equations",
    "word_insert_axmath",
    "word_update_axmath",
    "word_format_axmath",
    "word_refresh_equation_numbers",
    "word_get_equation_latex",
    "word_edit_axmath",
    "word_save_document",
    "ppt_list_presentations",
    "ppt_create_presentation",
    "ppt_open_presentation",
    "ppt_list_slides",
    "ppt_add_slide",
    "ppt_list_equations",
    "ppt_insert_axmath",
    "ppt_update_axmath",
    "ppt_position_axmath",
    "ppt_get_equation_latex",
    "ppt_edit_axmath",
    "ppt_save_presentation",
}


def payload(result: Any) -> dict[str, Any]:
    if result.isError:
        messages = [item.text for item in result.content if item.type == "text"]
        raise AssertionError("MCP tool failed: " + "\n".join(messages))
    if isinstance(result.structuredContent, dict):
        return result.structuredContent
    texts = [item.text for item in result.content if item.type == "text"]
    assert len(texts) == 1, "Expected one JSON result or structuredContent."
    value = json.loads(texts[0])
    assert isinstance(value, dict), "Expected a JSON object result."
    return value


async def exercise_word(session: ClientSession) -> dict[str, Any]:
    async def call(name: str, **arguments: Any) -> dict[str, Any]:
        return payload(await session.call_tool(name, arguments))

    created = await call("word_create_document", visible=False)
    document_id = created["document_id"]
    inserted = await call("word_insert_axmath", document_id=document_id, latex=r"\frac{a}{b}")
    equation_id = inserted["equation_id"]
    assert inserted["progid"] == "Equation.AxMath"
    assert inserted["width_pt"] > 0 and inserted["height_pt"] > 0

    exported = await call(
        "word_get_equation_latex", document_id=document_id, equation_id=equation_id
    )
    assert r"\frac" in exported["latex"], "Fraction did not survive AxMath source export."
    assert "a" in exported["latex"] and "b" in exported["latex"]

    updated = await call(
        "word_update_axmath",
        document_id=document_id,
        equation_id=equation_id,
        latex="x^2+y^2=z^2",
    )
    assert updated["equation_id"] == equation_id, "Replacement must preserve the equation ID."
    second_export = await call(
        "word_get_equation_latex", document_id=document_id, equation_id=equation_id
    )
    assert all(symbol in second_export["latex"] for symbol in ("x", "y", "z", "2", "="))

    equations = await call("word_list_equations", document_id=document_id)
    assert len(equations["equations"]) == 1
    assert equations["equations"][0]["equation_id"] == equation_id

    before = await call("word_document_info", document_id=document_id)
    observed = (before.get("selection") or {}).get("text", "")
    refused = await session.call_tool(
        "word_insert_axmath",
        {
            "document_id": document_id,
            "latex": "q+1",
            "position": "selection",
            "expected_selection": observed + "__AXMATH_SMOKE_MISMATCH__",
        },
    )
    assert refused.isError, "An uninspected/mismatched selection must be refused."
    after = await call("word_document_info", document_id=document_id)
    assert after["inline_shape_count"] == before["inline_shape_count"] == 1

    numbered = await call(
        "word_format_axmath", document_id=document_id, equation_id=equation_id, layout="numbered"
    )
    assert numbered["number_text"] == "(1)"
    refreshed = await call("word_refresh_equation_numbers", document_id=document_id)
    assert refreshed["numbers"] == ["1"]

    output_dir = ROOT / "artifacts"
    output_dir.mkdir(exist_ok=True)
    stamp = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%dT%H%M%S-%f")
    target = output_dir / f"mcp-smoke-{stamp}-{uuid.uuid4().hex[:8]}.docx"
    saved = await call("word_save_document", document_id=document_id, path=str(target))
    assert saved["saved"] is True
    assert target.is_file() and target.stat().st_size > 0
    from zipfile import ZipFile

    with ZipFile(target) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
    assert "AMMPlaceRM" in xml and "SEQ AMEqn" in xml
    assert "AxMathEquation" not in xml
    return {
        "document_id": document_id,
        "equation_id": equation_id,
        "inserted_latex_export": exported["latex"],
        "updated_latex_export": second_export["latex"],
        "native_ole_count": 1,
        "native_numbering_verified": True,
        "selection_mismatch_refused": True,
        "saved_document": str(target),
        "document_left_open": True,
    }


async def smoke(exercise: bool) -> dict[str, Any]:
    assert PYTHON.is_file(), "Run uv sync first to create the project's .venv."
    protocol_errors: list[str] = []

    async def on_message(message: Any) -> None:
        # The standard transport reports invalid stdout JSON as Exception messages.
        if isinstance(message, Exception):
            protocol_errors.append(str(message))

    parameters = StdioServerParameters(
        command=str(PYTHON),
        args=["-m", "axmath_mcp.server"],
        cwd=str(ROOT),
        env={"PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"},
    )
    # Server logs belong on stderr. Capture them so successful runs emit one JSON object.
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8") as server_log:
        try:
            async with (
                stdio_client(parameters, errlog=server_log) as (read, write),
                ClientSession(
                    read,
                    write,
                    read_timeout_seconds=timedelta(seconds=240),
                    message_handler=on_message,
                ) as session,
            ):
                initialized = await session.initialize()
                tools = await session.list_tools()
                names = {tool.name for tool in tools.tools}
                assert len(tools.tools) == len(EXPECTED_TOOLS) and names == EXPECTED_TOOLS, names
                health = payload(await session.call_tool("axmath_health", {}))
                assert health["platform"] == "win32"
                assert health["word_registered"], "Word COM is not registered."
                assert health["axmath_registered"], "AxMath OLE is not registered."
                assert health["axmath_executable_exists"], "AxMath executable is missing."
                assert health["startup_template_exists"], "AxMath Word template is missing."
                result: dict[str, Any] = {
                    "ok": True,
                    "mode": "exercise-word" if exercise else "read-only",
                    "server": initialized.serverInfo.name,
                    "protocol_version": initialized.protocolVersion,
                    "tool_count": len(tools.tools),
                    "health": health,
                }
                if exercise:
                    result["word"] = await exercise_word(session)
            assert not protocol_errors, "Nonprotocol stdout: " + "; ".join(protocol_errors)
            result["stdout_protocol_clean"] = True
            return result
        except BaseException:
            server_log.seek(0)
            log_text = server_log.read()
            if log_text:
                import sys

                print(log_text[-6000:], file=sys.stderr)
            raise


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--exercise-word",
        action="store_true",
        help="Create, insert, export, update and save a new smoke document; leave it open.",
    )
    args = parser.parse_args()
    print(json.dumps(asyncio.run(smoke(args.exercise_word)), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

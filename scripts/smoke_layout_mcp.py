"""Exercise layout tools through a fresh stdio MCP connection.

Default performs a read-only handshake. ``--exercise-word`` creates one new
document and tests automatic numbering, update, display and fixed numbering.
The saved final document remains open; no existing document is changed.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import math
import re
import sys
import tempfile
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta, timezone
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from smoke_mcp import EXPECTED_TOOLS, PYTHON, ROOT, payload


def progress(step: str, state: str, **values: Any) -> None:
    print(
        json.dumps({"step": step, "state": state, **values}, ensure_ascii=False),
        file=sys.stderr,
        flush=True,
    )


async def exercise_layouts(session: ClientSession, report: dict[str, Any]) -> None:
    async def call(step: str, name: str, **arguments: Any) -> dict[str, Any]:
        progress(step, "started")
        started = time.monotonic()
        value = payload(await session.call_tool(name, arguments))
        elapsed = round(time.monotonic() - started, 3)
        report.setdefault("steps", []).append(
            {"step": step, "elapsed_seconds": elapsed, "result": value}
        )
        progress(step, "passed", elapsed_seconds=elapsed)
        return value

    created = await call("create", "word_create_document", visible=False)
    document_id = created["document_id"]
    report["document_id"] = document_id

    async def check_current(
        equation_id: str, layout: str, number: str | None, auto_count: int
    ) -> None:
        listed = await call("list_" + layout, "word_list_equations", document_id=document_id)
        assert len(listed["equations"]) == 1, "Reformatting duplicated the native equation."
        current = listed["equations"][0]
        assert current["equation_id"] == equation_id, "Equation ID changed after reformatting."
        assert current["layout"] == layout, current
        assert current["number_text"] == number, current
        assert math.isclose(current["font_size_pt"], 14, abs_tol=0.01), current
        info = await call("info_" + layout, "word_document_info", document_id=document_id)
        assert info["inline_shape_count"] == 1
        refreshed = await call(
            "refresh_" + layout, "word_refresh_equation_numbers", document_id=document_id
        )
        assert refreshed["updated_count"] == auto_count, refreshed
        assert refreshed["numbers"] == (["1"] if auto_count else []), refreshed

    inserted = await call(
        "insert_auto_number_14pt",
        "word_insert_axmath",
        document_id=document_id,
        latex="x^2+y^2=z^2",
        layout="numbered",
        font_size_pt=14,
    )
    equation_id = inserted["equation_id"]
    report["equation_id"] = equation_id
    assert inserted["progid"] == "Equation.AxMath"
    assert inserted["layout"] == "numbered" and inserted["number_text"] == "(1)"
    assert math.isclose(inserted["font_size_pt"], 14, abs_tol=0.01)
    await check_current(equation_id, "numbered", "(1)", 1)

    updated = await call(
        "update_preserves_number_and_14pt",
        "word_update_axmath",
        document_id=document_id,
        equation_id=equation_id,
        latex="a^2+b^2=c^2",
    )
    assert updated["equation_id"] == equation_id
    assert updated["number_text"] == "(1)" and updated["layout"] == "numbered"
    assert math.isclose(updated["font_size_pt"], 14, abs_tol=0.01)
    compact_source = re.sub(r"\s+", "", updated["latex_exported"]).replace("{", "").replace("}", "")
    assert "a^2+b^2=c^2" in compact_source, updated["latex_exported"]
    await check_current(equation_id, "numbered", "(1)", 1)

    displayed = await call(
        "reformat_display",
        "word_format_axmath",
        document_id=document_id,
        equation_id=equation_id,
        layout="display",
    )
    assert displayed["equation_id"] == equation_id
    assert displayed["layout"] == "display" and displayed["number_text"] is None
    assert math.isclose(displayed["font_size_pt"], 14, abs_tol=0.01)
    await check_current(equation_id, "display", None, 0)

    numbered = await call(
        "reformat_native_number",
        "word_format_axmath",
        document_id=document_id,
        equation_id=equation_id,
        layout="numbered",
    )
    assert numbered["equation_id"] == equation_id
    assert numbered["number_label"] is None and numbered["number_text"] == "(1)"
    assert math.isclose(numbered["font_size_pt"], 14, abs_tol=0.01)
    await check_current(equation_id, "numbered", "(1)", 1)

    target = ROOT / "artifacts/layout-validation" / (report["run_id"] + ".docx")
    saved = await call("save", "word_save_document", document_id=document_id, path=str(target))
    assert saved["saved"] and target.is_file()
    with zipfile.ZipFile(target) as package:
        document_xml = ET.fromstring(package.read("word/document.xml"))
    ns_word = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    ns_ole = "{urn:schemas-microsoft-com:office:office}"
    field_codes = " ".join(node.text or "" for node in document_xml.iter(ns_word + "instrText"))
    objects = list(document_xml.iter(ns_ole + "OLEObject"))
    assert len(objects) == 1 and objects[0].attrib["ProgID"] == "Equation.AxMath"
    assert "AMMPlaceRM" in field_codes and "SEQ AMEqn" in field_codes
    from axmath_mcp.font_metrics import equation_run_metrics

    assert (
        equation_run_metrics(ET.tostring(document_xml, encoding="unicode"))["baseline_offset_pt"]
        < 0
    )
    assert not re.search(r"\bSEQ\s+AxMathEquation\b", field_codes, re.IGNORECASE)
    report["summary"] = {
        "saved_document": str(target),
        "final_layout": "numbered",
        "final_number": "(1)",
        "font_size_pt": 14,
        "native_ole_count": 1,
        "stale_auto_numbers": 0,
        "document_left_open": True,
    }


async def smoke(exercise: bool) -> dict[str, Any]:
    assert PYTHON.is_file(), "Run uv sync first."
    stamp = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%dT%H%M%S-%f")
    report: dict[str, Any] = {
        "run_id": "mcp-layout-regression-" + stamp + "-" + uuid.uuid4().hex[:8],
        "mode": "exercise-word" if exercise else "read-only",
        "ok": False,
    }
    report_path = None
    if exercise:
        output = ROOT / "artifacts/layout-validation"
        output.mkdir(parents=True, exist_ok=True)
        report_path = output / (report["run_id"] + ".json")
    protocol_errors: list[str] = []

    async def on_message(message: Any) -> None:
        if isinstance(message, Exception):
            protocol_errors.append(str(message))

    parameters = StdioServerParameters(
        command=str(PYTHON),
        args=["-m", "axmath_mcp.server"],
        cwd=str(ROOT),
        env={"PYTHONUTF8": "1"},
    )
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8") as server_log:
        try:
            async with (
                stdio_client(parameters, errlog=server_log) as (read, write),
                ClientSession(
                    read,
                    write,
                    read_timeout_seconds=timedelta(seconds=300),
                    message_handler=on_message,
                ) as session,
            ):
                initialized = await session.initialize()
                listed = await session.list_tools()
                assert {tool.name for tool in listed.tools} == EXPECTED_TOOLS
                assert len(listed.tools) == len(EXPECTED_TOOLS)
                health = payload(await session.call_tool("axmath_health", {}))
                assert health["word_registered"] and health["axmath_registered"]
                report.update(
                    {
                        "server": initialized.serverInfo.name,
                        "protocol_version": initialized.protocolVersion,
                        "tool_count": len(listed.tools),
                        "health": health,
                    }
                )
                if exercise:
                    await exercise_layouts(session, report)
            assert not protocol_errors, "Nonprotocol stdout: " + "; ".join(protocol_errors)
            report["ok"] = True
            report["stdout_protocol_clean"] = True
        except BaseException as exc:
            report["error"] = repr(exc)
            server_log.seek(0)
            logs = server_log.read()
            if logs:
                print(logs[-6000:], file=sys.stderr)
            raise
        finally:
            if report_path:
                report["report_path"] = str(report_path)
                report_path.write_text(
                    json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
                )
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--exercise-word",
        action="store_true",
        help="Create a new document and test layout mutations; save it and leave it open.",
    )
    args = parser.parse_args()
    print(json.dumps(asyncio.run(smoke(args.exercise_word)), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

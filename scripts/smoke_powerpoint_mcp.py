"""Exercise PowerPoint tools through a real stdio MCP connection.

Default checks registration and the tool list without opening Office. Explicit
``--exercise-powerpoint`` creates a two-slide presentation and leaves it open
after saving. This script does not modify an existing presentation.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import math
import sys
import tempfile
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from smoke_mcp import EXPECTED_TOOLS, PYTHON, ROOT, payload

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
EMU_PER_PT = 12700


def progress(step: str, state: str, **values: Any) -> None:
    print(
        json.dumps({"step": step, "state": state, **values}, ensure_ascii=False),
        file=sys.stderr,
        flush=True,
    )


def close_to(actual: float, expected: float) -> bool:
    return math.isclose(actual, expected, abs_tol=0.1)


def _geometry(node: ET.Element) -> dict[str, float]:
    transform = node.find("p:xfrm", NS)
    if transform is None:
        transform = node.find("p:spPr/a:xfrm", NS)
    assert transform is not None, "Saved shape has no explicit position."
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)
    assert offset is not None and extent is not None
    return {
        "left_pt": int(offset.attrib["x"]) / EMU_PER_PT,
        "top_pt": int(offset.attrib["y"]) / EMU_PER_PT,
        "width_pt": int(extent.attrib["cx"]) / EMU_PER_PT,
        "height_pt": int(extent.attrib["cy"]) / EMU_PER_PT,
    }


def inspect_saved_package(path: Path, moved: dict[str, Any]) -> dict[str, Any]:
    """Check real embedded equations and label geometry, not only screen previews."""
    with zipfile.ZipFile(path) as package:
        slide_names = sorted(
            name
            for name in package.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )
        assert len(slide_names) == 2, slide_names
        slides = [ET.fromstring(package.read(name)) for name in slide_names]
        native_shapes = [
            node
            for slide in slides
            for node in slide.findall(".//p:graphicFrame", NS)
            if node.find(".//p:oleObj", NS) is not None
        ]
        assert len(native_shapes) == 2, "Saved presentation must contain two native OLE shapes."
        # PowerPoint emits Choice/Fallback representations of one logical object.
        # Count its graphicFrame once and require both representations to agree.
        ole_objects = [obj for node in native_shapes for obj in node.findall(".//p:oleObj", NS)]
        assert all(obj.get("progId") == "Equation.AxMath" for obj in ole_objects), (
            "A formula was wrapped as another Office object instead of native AxMath."
        )
        relationship_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        for node in native_shapes:
            refs = {obj.get(relationship_id) for obj in node.findall(".//p:oleObj", NS)}
            assert len(refs) == 1 and None not in refs, (
                "OLE compatibility representations disagree."
            )
        embeddings = [name for name in package.namelist() if name.startswith("ppt/embeddings/")]
        assert len(embeddings) == 2, embeddings
        for name in embeddings:
            assert package.read(name).startswith(bytes.fromhex("d0cf11e0a1b11ae1")), name
        texts = [
            "".join(node.text or "" for node in slide.findall(".//a:t", NS)) for slide in slides
        ]
        assert "中文公式与固定编号" in texts[0]
        assert "Editable matrix in PowerPoint" in texts[1]
        assert texts[0].count("(1)") == 1 and "(1)" not in texts[1]

        # Word wrapping a clipboard object would give Word.Document.* as ProgID;
        # the checks above reject that case before inspecting exact placement.
        native_nodes = [
            node
            for node in slides[0].findall(".//p:graphicFrame", NS)
            if node.find(".//p:oleObj", NS) is not None
        ]
        assert len(native_nodes) == 1
        equation = _geometry(native_nodes[0])
        labels = [
            node
            for node in slides[0].findall(".//p:sp", NS)
            if "".join(t.text or "" for t in node.findall(".//a:t", NS)) == "(1)"
        ]
        assert len(labels) == 1
        label = _geometry(labels[0])
        for key in ("left_pt", "top_pt", "width_pt", "height_pt"):
            assert close_to(equation[key], moved[key]), (key, equation, moved)
        assert close_to(label["left_pt"], equation["left_pt"] + equation["width_pt"] + 12), (
            equation,
            label,
        )
        assert close_to(
            label["top_pt"] + label["height_pt"] / 2,
            equation["top_pt"] + equation["height_pt"] / 2,
        ), (equation, label)
    return {
        "slide_count": 2,
        "native_ole_count": 2,
        "embedding_count": 2,
        "progids": [node.find(".//p:oleObj", NS).get("progId") for node in native_shapes],
        "xml_ole_representations": len(ole_objects),
        "label_count": 1,
        "equation_geometry": equation,
        "label_geometry": label,
        "label_moves_with_equation": True,
    }


async def exercise_powerpoint(session: ClientSession, report: dict[str, Any]) -> None:
    async def call(step: str, name: str, **arguments: Any) -> dict[str, Any]:
        progress(step, "started")
        started = time.monotonic()
        result = payload(await session.call_tool(name, arguments))
        elapsed = round(time.monotonic() - started, 3)
        report.setdefault("steps", []).append(
            {"step": step, "elapsed_seconds": elapsed, "result": result}
        )
        progress(step, "passed", elapsed_seconds=elapsed)
        return result

    created = await call("create", "ppt_create_presentation", visible=False)
    pid = created["presentation_id"]
    report["presentation_id"] = pid
    first = await call(
        "add_chinese_slide", "ppt_add_slide", presentation_id=pid, title="中文公式与固定编号"
    )
    second = await call(
        "add_english_slide",
        "ppt_add_slide",
        presentation_id=pid,
        title="Editable matrix in PowerPoint",
    )
    first_id, second_id = first["slide_id"], second["slide_id"]
    assert first_id != second_id
    slides = await call("list_slides", "ppt_list_slides", presentation_id=pid)
    assert [slide["slide_id"] for slide in slides["slides"]] == [first_id, second_id]
    report["slide_ids"] = [first_id, second_id]

    inserted = await call(
        "insert_fraction",
        "ppt_insert_axmath",
        presentation_id=pid,
        slide_id=first_id,
        latex=r"\frac{a+b}{c+d}",
        left_pt=72,
        top_pt=144,
        font_size_pt=24,
        number_label="(1)",
    )
    equation_id = inserted["equation_id"]
    assert inserted["progid"] == "Equation.AxMath" and inserted["number_label"] == "(1)"
    assert close_to(inserted["left_pt"], 72) and close_to(inserted["top_pt"], 144)
    assert close_to(inserted["font_size_pt"], 24)
    exported = await call(
        "read_native_fraction",
        "ppt_get_equation_latex",
        presentation_id=pid,
        slide_id=first_id,
        equation_id=equation_id,
    )
    assert r"\frac" in exported["latex"]
    assert all(symbol in exported["latex"] for symbol in ("a", "b", "c", "d"))

    updated = await call(
        "update_integral",
        "ppt_update_axmath",
        presentation_id=pid,
        slide_id=first_id,
        equation_id=equation_id,
        latex=r"\int_0^1 x^2\,dx=\frac{1}{3}",
    )
    assert updated["equation_id"] == equation_id
    assert updated["number_label"] == "(1)" and close_to(updated["font_size_pt"], 24)
    assert close_to(updated["left_pt"], 72) and close_to(updated["top_pt"], 144)
    assert r"\int" in updated["latex_exported"] and "3" in updated["latex_exported"]

    moved = await call(
        "move_equation_and_label",
        "ppt_position_axmath",
        presentation_id=pid,
        slide_id=first_id,
        equation_id=equation_id,
        left_pt=180,
        top_pt=180,
    )
    assert moved["equation_id"] == equation_id
    assert close_to(moved["left_pt"], 180) and close_to(moved["top_pt"], 180)
    assert moved["number_label"] == "(1)"

    for step, arguments in (
        (
            "reject_off_slide",
            {"slide_id": first_id, "left_pt": slides["slide_width_pt"] + 1, "top_pt": 180},
        ),
        ("reject_wrong_slide", {"slide_id": second_id, "left_pt": 72, "top_pt": 144}),
    ):
        progress(step, "started")
        refused = await session.call_tool(
            "ppt_position_axmath", {"presentation_id": pid, "equation_id": equation_id, **arguments}
        )
        assert refused.isError, step
        report.setdefault("steps", []).append({"step": step, "expected_error": True})
        progress(step, "passed")
    after_refusal = await call(
        "check_unchanged_after_refusals", "ppt_list_equations", presentation_id=pid
    )
    assert len(after_refusal["equations"]) == 1
    preserved = after_refusal["equations"][0]
    assert preserved["equation_id"] == equation_id and preserved["slide_id"] == first_id
    assert close_to(preserved["left_pt"], 180) and close_to(preserved["top_pt"], 180)

    matrix = await call(
        "insert_matrix",
        "ppt_insert_axmath",
        presentation_id=pid,
        slide_id=second_id,
        latex=r"\mathbf{p}_{t+1}=\begin{pmatrix}0.8&0.3\\0.2&0.7\end{pmatrix}\mathbf{p}_t",
        left_pt=72,
        top_pt=144,
        font_size_pt=24,
    )
    assert matrix["progid"] == "Equation.AxMath" and matrix["number_label"] is None
    assert all(value in matrix["latex_exported"] for value in ("0.8", "0.3", "0.2", "0.7"))
    for slide_id, expected_id in ((first_id, equation_id), (second_id, matrix["equation_id"])):
        listed = await call(
            "list_slide_equations", "ppt_list_equations", presentation_id=pid, slide_id=slide_id
        )
        assert len(listed["equations"]) == 1
        assert listed["equations"][0]["equation_id"] == expected_id
    listed = await call("list_all_equations", "ppt_list_equations", presentation_id=pid)
    assert len(listed["equations"]) == 2

    target = ROOT / "artifacts/powerpoint-validation" / (report["run_id"] + ".pptx")
    saved = await call("save", "ppt_save_presentation", presentation_id=pid, path=str(target))
    assert saved["saved"] and target.is_file()
    package = inspect_saved_package(target, moved)
    report["summary"] = {
        "saved_presentation": str(target),
        "presentation_left_open": True,
        "font_size_pt": 24,
        "equation_ids": [equation_id, matrix["equation_id"]],
        "off_slide_move_refused": True,
        "wrong_slide_refused": True,
        **package,
    }


async def smoke(exercise: bool) -> dict[str, Any]:
    assert PYTHON.is_file(), "Run uv sync first."
    stamp = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%dT%H%M%S-%f")
    report: dict[str, Any] = {
        "run_id": "mcp-powerpoint-" + stamp + "-" + uuid.uuid4().hex[:8],
        "mode": "exercise-powerpoint" if exercise else "read-only",
        "ok": False,
    }
    report_path = None
    if exercise:
        output = ROOT / "artifacts/powerpoint-validation"
        output.mkdir(parents=True, exist_ok=True)
        report_path = output / (report["run_id"] + ".json")
    protocol_errors = []

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
                assert len(listed.tools) == 25
                assert {tool.name for tool in listed.tools} == EXPECTED_TOOLS
                health = payload(await session.call_tool("axmath_health", {}))
                assert health["word_registered"] and health["axmath_registered"]
                assert health["powerpoint_registered"]
                report.update(
                    {
                        "server": initialized.serverInfo.name,
                        "tool_count": len(listed.tools),
                        "protocol_version": initialized.protocolVersion,
                        "health": health,
                    }
                )
                if exercise:
                    await exercise_powerpoint(session, report)
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
        "--exercise-powerpoint",
        action="store_true",
        help="Create, edit, position and save a new two-slide presentation.",
    )
    args = parser.parse_args()
    print(json.dumps(asyncio.run(smoke(args.exercise_powerpoint)), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

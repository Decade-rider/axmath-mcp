"""Generate native AxMath assets used only by the typography verification suite.

Run this on Windows with licensed AxMath and its Word add-in installed. Each
distinct (LaTeX, inline/display mode, font size) is converted and exported once.
The companion JSON records actual exported sources and object dimensions. Nothing
is published to the requested paths unless all conversions and saved-package
checks succeed. Existing asset pairs are never overwritten; choose a new --output
path to regenerate them.

This reuse is limited to verification. It does not add a production formula cache.
Baseline verification belongs to verify_layouts.py and its OOXML/render checks:
Range.Font.Position over a mixed OLE range may be Word's wdUndefined sentinel.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import time
import zipfile
from contextlib import ExitStack
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from layout_cases import CASES, LayoutCase

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / ".research" / "layout-assets.docx"


def asset_key(case: LayoutCase) -> tuple[str, str, float]:
    """Numbered and unnumbered display cases share the same equation object."""
    mode = "inline" if case.layout == "inline" else "display"
    return case.latex, mode, case.font_size_pt


def unique_cases() -> list[LayoutCase]:
    result: list[LayoutCase] = []
    seen: set[tuple[str, str, float]] = set()
    for case in CASES:
        key = asset_key(case)
        if key not in seen:
            seen.add(key)
            result.append(case)
    return result


def generate_staged_assets(path: Path) -> list[dict[str, Any]]:
    """Own and close one isolated Word application, including on failed setup."""
    import pythoncom
    import win32com.client

    from axmath_mcp.backend import WordBackend
    from axmath_mcp.validation import latex_for_layout
    from axmath_mcp.word_transfer import copy_native_range

    records: list[dict[str, Any]] = []
    cases = unique_cases()
    with ExitStack() as cleanup:
        pythoncom.CoInitialize()
        cleanup.callback(pythoncom.CoUninitialize)
        app = win32com.client.DispatchEx("Word.Application")
        cleanup.callback(app.Quit, 0)
        app.Visible = False
        backend = WordBackend()
        cleanup.callback(backend.shutdown)
        doc = app.Documents.Add(Visible=False)
        cleanup.callback(doc.Close, 0)

        for case in cases:
            latex, mode, font_size = asset_key(case)
            started = time.monotonic()
            print("generating " + case.case_id, flush=True)
            with (
                backend._conversion() as converter,
                backend._generated(converter, latex_for_layout(latex, mode), font_size) as (
                    shape,
                    source,
                ),
            ):
                doc.Activate()
                destination = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                copy_native_range(destination, shape.Range)
                if destination.InlineShapes.Count != 1:
                    raise RuntimeError(f"Expected one copied native object for {case.case_id}.")
                inserted = destination.InlineShapes(1)
                if (
                    not backend._is_equation(inserted)
                    or inserted.Width <= 0
                    or inserted.Height <= 0
                ):
                    raise RuntimeError(f"Invalid native AxMath asset for {case.case_id}.")
                inserted.AlternativeText = case.case_id
                records.append(
                    {
                        "key": case.case_id,
                        "latex": latex,
                        "mode": mode,
                        "font_size_pt": font_size,
                        "exported": source,
                        "width": float(inserted.Width),
                        "height": float(inserted.Height),
                        "reused_by": [c.case_id for c in CASES if asset_key(c) == asset_key(case)],
                    }
                )
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
            print(
                json.dumps(
                    {
                        "generated": case.case_id,
                        "seconds": round(time.monotonic() - started, 2),
                    }
                ),
                flush=True,
            )

        if int(doc.InlineShapes.Count) != len(cases) or len(records) != len(cases):
            raise RuntimeError("Asset count does not match the deduplicated case set.")
        doc.SaveAs2(FileName=str(path), FileFormat=16, AddToRecentFiles=False)
        if not doc.Saved:
            raise RuntimeError("Word did not confirm saving the asset document.")

    # Inspect the file only after closing our Word document and application.
    with zipfile.ZipFile(path) as package:
        tree = ET.fromstring(package.read("word/document.xml"))
        count = len(tree.findall(".//{urn:schemas-microsoft-com:office:office}OLEObject"))
        if count != len(records):
            raise RuntimeError(f"Saved DOCX contains {count} OLE objects, expected {len(records)}.")
    return records


def build(output: Path) -> dict[str, Any]:
    if sys.platform != "win32":
        raise RuntimeError("Asset generation requires Windows desktop Word and AxMath.")
    output = output if output.is_absolute() else ROOT / output
    output = output.resolve()
    if output.suffix.lower() != ".docx":
        raise ValueError("--output must name a .docx file.")
    companion = output.with_suffix(".json")
    if output.exists() or companion.exists():
        raise FileExistsError(
            f"An asset output already exists: {output} or {companion}. "
            "Existing files were left unchanged; choose a new --output path."
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".axmath-assets-", dir=output.parent) as temporary:
        staged_docx = Path(temporary) / "assets.docx"
        staged_json = Path(temporary) / "assets.json"
        records = generate_staged_assets(staged_docx)
        staged_json.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
        # On Windows os.rename fails if the destination was created concurrently.
        # Publish the JSON last: an incomplete pair must never look like a cache hit.
        os.rename(staged_docx, output)
        try:
            os.rename(staged_json, companion)
        except BaseException:
            output.unlink(missing_ok=True)
            raise
    return {
        "ok": True,
        "docx": str(output),
        "json": str(companion),
        "native_assets": len(records),
        "layout_cases": len(CASES),
        "reuse_scope": "typography verification only",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help="New asset DOCX path; writes a JSON file beside it. Default: .research/layout-assets.docx",
    )
    args = parser.parse_args()
    try:
        result = build(args.output)
    except Exception as exc:
        raise SystemExit(
            f"Asset generation failed; no successful new pair published: {exc}"
        ) from exc
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()

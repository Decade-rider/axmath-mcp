"""Build and verify bilingual Word typography with real editable AxMath objects.

Run on Windows with Word/AxMath installed. Cached test assets reuse seven native
conversions; production calls still convert each request. --assets points at the
asset DOCX, paired with a JSON file of exported sources.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from contextlib import contextmanager
from pathlib import Path
from xml.etree import ElementTree as ET

import pythoncom
import win32com.client
from layout_cases import CASES, PAGE_INTRODUCTIONS, PAGE_TITLES

from axmath_mcp.backend import WordBackend
from axmath_mcp.layout import is_number_field
from axmath_mcp.validation import latex_for_layout

ROOT = Path(__file__).resolve().parents[1]
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
}


class CachedTestBackend(WordBackend):
    def __init__(self, assets, records):
        super().__init__()
        shapes = {str(s.AlternativeText): s for s in assets.InlineShapes}
        self.assets = {
            latex_for_layout(r["latex"], r["mode"]): (shapes[r["key"]], r["exported"])
            for r in records
        }

    @contextmanager
    def _generated(self, app, source, font_size=12):
        if font_size != 12:
            raise ValueError("Test assets are 12 pt only; use real conversion for other sizes.")
        shape, latex = self.assets[source]
        shape.Range.Font.Size = font_size
        yield shape, latex


def append_paragraph(doc, text, *, size=12, bold=False, page_break=False):
    start = int(doc.Content.End - 1)
    doc.Range(start, start).InsertAfter(text + "\r")
    rng = doc.Range(start, doc.Content.End - 1)
    rng.Font.NameFarEast = "宋体"
    rng.Font.NameAscii = "Times New Roman"
    rng.Font.NameOther = "Times New Roman"
    rng.Font.Size = size
    rng.Font.Bold = bold
    rng.Font.Color = 0
    rng.ParagraphFormat.SpaceAfter = 8
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.LineSpacingRule = 3
    rng.ParagraphFormat.LineSpacing = 18
    rng.ParagraphFormat.PageBreakBefore = page_break
    rng.ParagraphFormat.DisableLineHeightGrid = True
    rng.ParagraphFormat.Alignment = 0
    return start


def inspect_package(path):
    with zipfile.ZipFile(path) as z:
        tree = ET.fromstring(z.read("word/document.xml"))
        style_tree = ET.fromstring(z.read("word/styles.xml"))
        styles = {s.get("{" + NS["w"] + "}styleId"): s for s in style_tree.findall("w:style", NS)}
        defaults = style_tree.find("w:docDefaults/w:pPrDefault/w:pPr/w:snapToGrid", NS)

        def snap_value(paragraph):
            grid = paragraph.find("w:pPr/w:snapToGrid", NS)
            style = paragraph.find("w:pPr/w:pStyle", NS)
            sid = (
                style.get("{" + NS["w"] + "}val")
                if style is not None
                else next(
                    (
                        key
                        for key, s in styles.items()
                        if s.get("{" + NS["w"] + "}default") == "1"
                        and s.get("{" + NS["w"] + "}type") == "paragraph"
                    ),
                    None,
                )
            )
            seen = set()
            while grid is None and sid in styles and sid not in seen:
                seen.add(sid)
                node = styles[sid]
                grid = node.find("w:pPr/w:snapToGrid", NS)
                base = node.find("w:basedOn", NS)
                sid = base.get("{" + NS["w"] + "}val") if base is not None else None
            grid = defaults if grid is None else grid
            return grid.get("{" + NS["w"] + "}val", "1") if grid is not None else "1"

        equations = []
        for paragraph in tree.findall(".//w:p", NS):
            for run in paragraph.findall("w:r", NS):
                if run.find("w:object", NS) is None:
                    continue
                offset = run.find("w:rPr/w:position", NS)
                equations.append(
                    {
                        "baseline_half_pt": None
                        if offset is None
                        else offset.get("{" + NS["w"] + "}val"),
                        "snap_to_grid": snap_value(paragraph),
                    }
                )
        fields = ["".join(n.itertext()) for n in tree.findall(".//w:instrText", NS)]
        return {
            "equations": equations,
            "sequence_fields": [s for s in fields if "SEQ AMEqn" in s],
            "ole_count": len(tree.findall(".//o:OLEObject", NS)),
        }


def build(assets_path, output):
    output.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    assets = doc = backend = None
    records = json.loads(assets_path.with_suffix(".json").read_text(encoding="utf-8"))
    checks = []
    try:
        assets = app.Documents.Open(
            FileName=str(assets_path), ReadOnly=True, AddToRecentFiles=False, Visible=False
        )
        doc = app.Documents.Add(Visible=False)
        doc.PageSetup.PageWidth = 595.3
        doc.PageSetup.PageHeight = 841.9
        doc.PageSetup.TopMargin = 60
        doc.PageSetup.BottomMargin = 60
        doc.PageSetup.LeftMargin = 72
        doc.PageSetup.RightMargin = 72
        normal = doc.Styles(-1)
        normal.Font.NameFarEast = "宋体"
        normal.Font.NameAscii = "Times New Roman"
        normal.Font.NameOther = "Times New Roman"
        normal.Font.Size = 12
        normal.ParagraphFormat.LineSpacingRule = 3
        normal.ParagraphFormat.LineSpacing = 18
        normal.ParagraphFormat.DisableLineHeightGrid = True
        backend = CachedTestBackend(assets, records)
        did = backend._remember_doc(doc)
        current_page = 0
        for case in CASES:
            if case.page != current_page:
                start = append_paragraph(
                    doc, PAGE_TITLES[case.page], size=16, bold=True, page_break=case.page > 1
                )
                doc.Range(start, start + len(PAGE_TITLES[case.page])).Style = doc.Styles(-63)
                # Explicit title formatting, preserving a plain black readable hierarchy.
                title = doc.Range(start, start + len(PAGE_TITLES[case.page]))
                title.Font.Color = 0
                title.Font.Size = 16
                title.Font.NameFarEast = "黑体"
                title.ParagraphFormat.PageBreakBefore = case.page > 1
                title.ParagraphFormat.KeepWithNext = True
                append_paragraph(doc, PAGE_INTRODUCTIONS[case.page], size=10)
                current_page = case.page
            append_paragraph(doc, case.title, size=11, bold=True)
            token = "[[" + case.case_id + "]]"
            start = append_paragraph(doc, case.before + token + case.after)
            position = start + len(case.before)
            target = doc.Range(position, position + len(token))
            # Explicitly stress the fixed-line-height clipping repair in one case.
            if case.case_id == "zh_inline_fraction":
                target.ParagraphFormat.LineSpacingRule = 4
                target.ParagraphFormat.LineSpacing = 12
            doc.Activate()
            target.Select()
            result = backend.insert_equation(
                did,
                case.latex,
                position="selection",
                expected_selection=token,
                layout=case.layout,
                font_size_pt=12,
            )
            shape = backend._eq(did, result["equation_id"])
            assert case.before in doc.Content.Text and case.after in doc.Content.Text
            assert result["layout"] == case.layout
            assert shape.Range.ParagraphFormat.LineSpacingRule != 4
            assert bool(shape.Range.ParagraphFormat.DisableLineHeightGrid)
            checks.append({"case_id": case.case_id, **result})
            print("checked " + case.case_id, flush=True)
        numbers = backend.refresh_equation_numbers(did)
        assert numbers["numbers"] == ["1", "2", "3", "4"], numbers
        assert sum(is_number_field(f) for f in doc.Fields) == 4
        assert doc.InlineShapes.Count == 14
        doc.Repaginate()
        file = output / "axmath-layout-validation.docx"
        doc.SaveAs2(FileName=str(file), FileFormat=16, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(str(file.with_suffix(".pdf")), 17)
        package = inspect_package(file)
        assert package["ole_count"] == 14
        assert len(package["sequence_fields"]) == 4
        assert all(e["snap_to_grid"] == "0" for e in package["equations"]), package
        originals = inspect_package(assets_path)["equations"]
        native_offsets = {
            latex_for_layout(r["latex"], r["mode"]): m["baseline_half_pt"]
            for r, m in zip(records, originals, strict=True)
        }
        for case, actual in zip(CASES, package["equations"], strict=True):
            expected_offset = native_offsets[latex_for_layout(case.latex, case.layout)]
            assert actual["baseline_half_pt"] == expected_offset, (
                case.case_id,
                actual,
                expected_offset,
            )
        result = {
            "cases": checks,
            "numbering": numbers,
            "package": package,
            "pages": int(doc.ComputeStatistics(2)),
            "docx": str(file),
            "pdf": str(file.with_suffix(".pdf")),
            "asset_reuse": "Seven real vendor conversions reused across 14 paragraph contexts; not production caching.",
        }
        (output / "verification.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(
            json.dumps(
                {
                    "docx": str(file),
                    "pages": result["pages"],
                    "checks": len(checks),
                    "numbering": numbers,
                }
            ),
            flush=True,
        )
    finally:
        if doc is not None:
            doc.Close(0)
        if assets is not None:
            assets.Close(0)
        if backend is not None:
            backend.shutdown()
        app.Quit(0)
        pythoncom.CoUninitialize()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--assets", type=Path, default=ROOT / ".research/layout-assets.docx")
    parser.add_argument("--output", type=Path, default=ROOT / "artifacts/layout-validation")
    args = parser.parse_args()
    if not args.assets.is_file():
        sys.exit("Generate real AxMath assets first; see scripts/build_layout_assets.py.")
    build(args.assets.resolve(), args.output.resolve())


if __name__ == "__main__":
    main()

"""Reopen a saved PowerPoint copy, export slides and check editable AxMath objects.

The original file stays untouched. A separate copy avoids file-in-use dialogs.
PNG exports support visual inspection; successful structural checks do not replace
viewing those images. --read-latex also uses desktop Word/AxMath and the clipboard.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import json
import math
import shutil
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client

from axmath_mcp.backend import WordBackend
from axmath_mcp.powerpoint import LABEL_GAP, LABEL_TAG, OWNER_TAG, PowerPointBackend
from axmath_mcp.validation import validate_local_path


def _progress(step: str, **values: Any) -> None:
    print(json.dumps({"step": step, **values}, ensure_ascii=False), file=sys.stderr, flush=True)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _shapes(slide: Any) -> list[Any]:
    return [slide.Shapes.Item(index) for index in range(1, int(slide.Shapes.Count) + 1)]


def _geometry(shape: Any) -> dict[str, Any]:
    return {
        "shape_id": int(shape.Id),
        "name": str(shape.Name),
        "type": int(shape.Type),
        "left_pt": float(shape.Left),
        "top_pt": float(shape.Top),
        "width_pt": float(shape.Width),
        "height_pt": float(shape.Height),
        "rotation_deg": float(shape.Rotation),
    }


def _check_bounds(geometry: dict[str, Any], width: float, height: float) -> None:
    left, top = geometry["left_pt"], geometry["top_pt"]
    object_width, object_height = geometry["width_pt"], geometry["height_pt"]
    tolerance = 0.25
    if (
        not all(math.isfinite(value) for value in (left, top, object_width, object_height))
        or object_width < 0
        or object_height < 0
        or left < -tolerance
        or top < -tolerance
        or left + object_width > width + tolerance
        or top + object_height > height + tolerance
    ):
        raise AssertionError(f"Shape falls outside slide bounds: {geometry}")


def verify(
    source: Path,
    output: Path,
    *,
    read_latex: bool = False,
    expected_slides: int | None = None,
    expected_equations: int | None = None,
) -> dict[str, Any]:
    source = validate_local_path(str(source.resolve()), {".pptx"})
    if not source.is_file():
        raise ValueError("Source PowerPoint file does not exist.")
    output = output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    report: dict[str, Any] = {
        "ok": False,
        "source": str(source),
        "source_sha256": _sha256(source),
        "output": str(output),
        "read_latex": read_latex,
        "slides": [],
        "visual_review_required": True,
        "scope": "Top-level shapes; COM geometry checked, followed by manual PNG inspection.",
    }
    app = presentation = word = ppt = None
    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
        with tempfile.TemporaryDirectory(prefix="axmath-ppt-reopen-") as directory:
            copy = Path(directory) / source.name
            shutil.copy2(source, copy)
            assert _sha256(copy) == report["source_sha256"]
            try:
                _progress("open_saved_copy", source=str(source))
                app = win32com.client.DispatchEx("PowerPoint.Application")
                security = app.AutomationSecurity
                try:
                    app.AutomationSecurity = 3
                    presentation = app.Presentations.Open(
                        FileName=str(copy), ReadOnly=True, WithWindow=False
                    )
                finally:
                    app.AutomationSecurity = security
                word = WordBackend()
                ppt = PowerPointBackend(word)
                pid = ppt._remember(presentation)
                slide_info = ppt.ppt_list_slides(pid)
                width = float(slide_info["slide_width_pt"])
                height = float(slide_info["slide_height_pt"])
                count = len(slide_info["slides"])
                if expected_slides is not None:
                    assert count == expected_slides, (count, expected_slides)
                image_width = 1600
                image_height = max(1, round(image_width * height / width))
                report.update(
                    {
                        "slide_count": count,
                        "slide_width_pt": width,
                        "slide_height_pt": height,
                        "image_size_px": [image_width, image_height],
                    }
                )
                total = 0
                for item in slide_info["slides"]:
                    slide_id = item["slide_id"]
                    slide = presentation.Slides.FindBySlideID(slide_id)
                    page: dict[str, Any] = {**item, "shapes": [], "equations": []}
                    report["slides"].append(page)
                    shapes = _shapes(slide)
                    for shape in shapes:
                        geometry = _geometry(shape)
                        _check_bounds(geometry, width, height)
                        page["shapes"].append(geometry)
                    png = output / f"slide-{item['index']:03d}.png"
                    _progress("export_slide", slide_id=slide_id, path=str(png))
                    slide.Export(str(png), "PNG", image_width, image_height)
                    assert png.is_file() and png.stat().st_size > 0
                    page["png"] = str(png)
                    equations = ppt.ppt_list_equations(pid, slide_id)["equations"]
                    total += len(equations)
                    native_ids = {entry["shape_id"] for entry in equations}
                    for entry in equations:
                        equation = dict(entry)
                        page["equations"].append(equation)
                        _, _, shape = ppt._eq(pid, slide_id, entry["equation_id"])
                        assert entry["progid"] == "Equation.AxMath"
                        label = ppt._label(slide, shape)
                        marker = str(shape.Tags.Item(LABEL_TAG))
                        if marker:
                            assert label is not None, "Saved equation lost its label association."
                        if label is not None:
                            number = _geometry(label)
                            _check_bounds(number, width, height)
                            assert math.isclose(
                                number["left_pt"],
                                entry["left_pt"] + entry["width_pt"] + LABEL_GAP,
                                abs_tol=0.25,
                            ), "Number no longer follows the equation horizontally."
                            assert math.isclose(
                                number["top_pt"] + number["height_pt"] / 2,
                                entry["top_pt"] + entry["height_pt"] / 2,
                                abs_tol=0.25,
                            ), "Number no longer aligns with the equation vertically."
                            equation["label_geometry"] = number
                            equation["label_association_verified"] = True
                        if read_latex:
                            _progress(
                                "read_native_latex",
                                slide_id=slide_id,
                                equation_id=entry["equation_id"],
                            )
                            exported = ppt.ppt_get_equation_latex(
                                pid, slide_id, entry["equation_id"]
                            )
                            assert exported["latex"].strip()
                            equation["latex"] = exported["latex"]
                            equation["latex_source"] = exported["source"]
                    for shape in shapes:
                        owner = str(shape.Tags.Item(OWNER_TAG))
                        if owner:
                            assert owner.isdecimal() and int(owner) in native_ids, (
                                "An equation label has no native equation owner."
                            )
                            owned_shape = next(s for s in shapes if int(s.Id) == int(owner))
                            assert str(owned_shape.Tags.Item(LABEL_TAG)) == str(shape.Id)
                if expected_equations is not None:
                    assert total == expected_equations, (total, expected_equations)
                assert _sha256(source) == report["source_sha256"], (
                    "Original file changed during verification."
                )
                report.update({"native_ole_count": total, "source_unchanged": True, "ok": True})
            finally:
                cleanup = []
                if presentation is not None:
                    try:
                        presentation.Close()
                    except pywintypes.com_error as exc:
                        cleanup.append("Could not close verification copy: " + repr(exc))
                if ppt is not None:
                    ppt.shutdown()
                if word is not None:
                    word.shutdown()
                if app is not None:
                    try:
                        # PowerPoint may reuse a running instance. Only quit when
                        # no presentation remains; never close another document.
                        if app.Presentations.Count == 0:
                            app.Quit()
                    except pywintypes.com_error as exc:
                        cleanup.append("PowerPoint cleanup: " + repr(exc))
                presentation = ppt = word = app = None
                gc.collect()
                if cleanup:
                    report["cleanup_notes"] = cleanup
    except BaseException as exc:
        report["ok"] = False
        report["error"] = repr(exc)
        raise
    finally:
        if initialized:
            pythoncom.CoUninitialize()
        (output / "verification.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "path", type=Path, help="Saved .pptx file to reopen through a temporary copy."
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="New output directory; existing directories are not overwritten.",
    )
    parser.add_argument(
        "--read-latex", action="store_true", help="Read each native equation via Word/AxMath."
    )
    parser.add_argument("--expected-slides", type=int)
    parser.add_argument("--expected-equations", type=int)
    args = parser.parse_args()
    for name in ("expected_slides", "expected_equations"):
        value = getattr(args, name)
        if value is not None and value < 0:
            parser.error(f"--{name.replace('_', '-')} must be nonnegative.")
    source = args.path.resolve()
    output = args.output or source.parent / (source.stem + "-verification-" + uuid.uuid4().hex[:8])
    report = verify(
        source,
        output,
        read_latex=args.read_latex,
        expected_slides=args.expected_slides,
        expected_equations=args.expected_equations,
    )
    print(
        json.dumps(
            {
                "ok": report["ok"],
                "output": report["output"],
                "slide_count": report["slide_count"],
                "native_ole_count": report["native_ole_count"],
                "visual_review_required": True,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()

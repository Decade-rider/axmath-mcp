"""Word paragraph layout, native baseline preservation and equation numbering."""

from __future__ import annotations

import re
import uuid
from functools import lru_cache
from typing import Any

from axmath_mcp.font_metrics import equation_font_size, read_equation_metrics
from axmath_mcp.native_numbering import insert_native_number, preflight_native_numbering
from axmath_mcp.validation import LayoutOptions

DISPLAY_STYLE = "AxMathMCP_Display"
NUMBERED_STYLE = "AxMathMCP_Numbered"
NUMBER_BOOKMARK = "AxMathMCP_N_"
NUMBER_GAP_PT = 8.0
_NATIVE_SEQ = re.compile(r"^\s*SEQ\s+(AMEqn|AMChap|AMSec)(?:\s|$)", re.IGNORECASE)


def is_number_field(field: Any) -> bool:
    code = str(field.Code.Text)
    return int(field.Type) == 12 and bool(_NATIVE_SEQ.match(code))


def is_visible_number(field: Any) -> bool:
    code = str(field.Code.Text)
    return is_number_field(field) and (
        bool(re.match(r"\s*SEQ\s+AMEqn\s+\\c(?:\s|$)", code, re.IGNORECASE))
    )


def preflight_numbers(doc: Any) -> None:
    for field in doc.Fields:
        if is_number_field(field) and field.Locked:
            raise ValueError("An AxMath equation number field is locked; unlock it before refresh.")


def refresh_numbers(doc: Any) -> dict[str, Any]:
    updated = []
    # Never update arbitrary document fields: they may include external links.
    preflight_numbers(doc)
    for field in doc.Fields:
        if is_number_field(field):
            if not field.Update():
                raise RuntimeError("Word could not update an AxMath equation number.")
            if is_visible_number(field):
                updated.append(str(field.Result.Text))
    return {"updated_count": len(updated), "numbers": updated}


def describe_layout(shape: Any) -> dict[str, Any]:
    paragraph = shape.Range.Paragraphs(1)
    style = str(paragraph.Range.Style.NameLocal)
    mode = {DISPLAY_STYLE: "display", NUMBERED_STYLE: "numbered"}.get(style, "inline")
    result = {"layout": mode, "number_label": None, "number_text": None}
    if mode == "numbered":
        markers = [b for b in paragraph.Range.Bookmarks if str(b.Name).startswith(NUMBER_BOOKMARK)]
        if len(markers) == 1:
            number = markers[0].Range
            native = [f for f in number.Fields if is_visible_number(f)]
            result["number_text"] = (
                "(" + str(native[0].Result.Text) + ")" if native else str(number.Text)
            )
            if not any(is_number_field(f) for f in number.Fields):
                result["number_label"] = str(number.Text)
    return result


def managed_target(doc: Any, shape: Any) -> Any:
    """Expand to our own block only when it contains no unrelated user content."""
    state = describe_layout(shape)
    if state["layout"] == "inline":
        return shape.Range.Duplicate
    paragraph = shape.Range.Paragraphs(1).Range
    if paragraph.InlineShapes.Count != 1:
        raise ValueError(
            "The formula paragraph contains other objects; isolate it before reformatting."
        )
    before = str(doc.Range(paragraph.Start, shape.Range.Start).Text)
    after = str(doc.Range(shape.Range.End, paragraph.End - 1).Text)
    if state["layout"] == "display":
        safe = not before.strip() and not after.strip()
    else:
        markers = [b for b in paragraph.Bookmarks if str(b.Name).startswith(NUMBER_BOOKMARK)]
        safe = len(markers) == 1
        if safe:
            marker = markers[0].Range
            safe = (
                marker.Start >= shape.Range.End
                and not before.strip()
                and not str(doc.Range(shape.Range.End, marker.Start).Text).strip()
                and not str(doc.Range(marker.End, paragraph.End - 1).Text).strip()
            )
    if not safe:
        raise ValueError(
            "The managed formula paragraph has additional text; separate it before reformatting."
        )
    return doc.Range(paragraph.Start, paragraph.End - 1)


def available_width(target: Any) -> float:
    if target.Information(12):  # wdWithInTable
        raise ValueError("Display/numbered layout inside tables is not supported yet.")
    setup = target.Sections(1).PageSetup
    if setup.TextColumns.Count != 1:
        raise ValueError("Display/numbered layout currently requires a single-column section.")
    gutter = float(setup.Gutter) if int(setup.GutterPos) != 1 else 0.0  # top gutter uses no width
    return float(setup.PageWidth - setup.LeftMargin - setup.RightMargin - gutter)


@lru_cache(maxsize=256)
def _number_text_width_em(text: str) -> float:
    """Measure the actual number typeface using a printer-independent memory DC.

    A large negative font height requests a 1024-unit character em. Scaling the
    measured width back to points avoids screen-DPI/font-size rounding. GDI uses
    the selected Times New Roman font; no printer or Office application is opened.
    """
    import ctypes
    from ctypes import wintypes

    if not hasattr(ctypes, "WinDLL"):
        raise RuntimeError("Equation-number measurement requires Windows GDI.")
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.CreateCompatibleDC.argtypes = [wintypes.HDC]
    gdi.CreateCompatibleDC.restype = wintypes.HDC
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    gdi.DeleteDC.restype = wintypes.BOOL
    gdi.CreateFontW.argtypes = [ctypes.c_int] * 5 + [wintypes.DWORD] * 8 + [wintypes.LPCWSTR]
    gdi.CreateFontW.restype = wintypes.HFONT
    gdi.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
    gdi.SelectObject.restype = wintypes.HGDIOBJ
    gdi.DeleteObject.argtypes = [wintypes.HGDIOBJ]
    gdi.DeleteObject.restype = wintypes.BOOL
    gdi.GetTextExtentPoint32W.argtypes = [
        wintypes.HDC,
        wintypes.LPCWSTR,
        ctypes.c_int,
        ctypes.POINTER(wintypes.SIZE),
    ]
    gdi.GetTextExtentPoint32W.restype = wintypes.BOOL
    dc = gdi.CreateCompatibleDC(None)
    if not dc:
        raise RuntimeError("Could not create the equation-number measurement context.")
    font = old_font = None
    try:
        font = gdi.CreateFontW(-1024, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, "Times New Roman")
        if not font:
            raise RuntimeError("Could not load the equation-number font.")
        old_font = gdi.SelectObject(dc, font)
        if not old_font or old_font == ctypes.c_void_p(-1).value:
            old_font = None
            raise RuntimeError("Could not select the equation-number font.")
        extent = wintypes.SIZE()
        units = len(text.encode("utf-16-le")) // 2
        if not gdi.GetTextExtentPoint32W(dc, text, units, ctypes.byref(extent)):
            raise RuntimeError("Could not measure the equation-number label.")
        return float(extent.cx) / 1024.0
    finally:
        if old_font:
            gdi.SelectObject(dc, old_font)
        if font:
            gdi.DeleteObject(font)
        gdi.DeleteDC(dc)


def number_reserve_pt(target: Any, options: LayoutOptions, size: float) -> float:
    """Reserve rendered label width plus a gap, including the next SEQ digit count."""
    if options.number_label is not None:
        labels = [options.number_label]
    else:
        fields = [field for field in target.Document.Fields if is_visible_number(field)]
        # A format/update may not add a field, but allowing one prospective field
        # also covers insertion across 9->10, 99->100, etc. Existing restart values
        # can be larger than the field count, so include their digit lengths too.
        digits = len(str(len(fields) + 1))
        labels = []
        for field in fields:
            value = str(field.Result.Text).strip()
            if value.isascii() and value.isdecimal():
                digits = max(digits, len(value))
            elif value:
                labels.append("(" + value + ")")
        labels.append("(" + "8" * digits + ")")
    return max(_number_text_width_em(label) * size for label in labels) + NUMBER_GAP_PT


def check_block_fits(target: Any, shape: Any, options: LayoutOptions, size: float) -> float:
    width = available_width(target)
    # Keep the formula centered in the entire text width. Therefore the real
    # right-side number width + gap must also fit in the left-side margin.
    reserve = number_reserve_pt(target, options, size) if options.layout == "numbered" else 0.0
    allowed = width - 2 * reserve
    if allowed <= 0 or float(shape.Width) > allowed:
        raise ValueError(
            f"Equation is {float(shape.Width):.1f} pt wide, but {allowed:.1f} pt is available "
            "for this layout. Split the formula, use a smaller font, or use unnumbered display."
        )
    return width


def _style(doc: Any, name: str) -> Any:
    if not any(str(s.NameLocal) == name for s in doc.Styles):
        style = doc.Styles.Add(Name=name, Type=1)
        style.BaseStyle = doc.Styles(-1)  # wdStyleNormal; no hardcoded font override
    return doc.Styles(name)


def inline_paragraph(shape: Any) -> None:
    """Preserve AxMath's half-point run offset; only adjust paragraph constraints."""
    paragraph = shape.Range.Paragraphs(1).Format
    paragraph.BaseLineAlignment = 2  # wdBaselineAlignBaseline
    paragraph.DisableLineHeightGrid = True  # verified as snapToGrid=0 by integration QA
    if paragraph.LineSpacingRule == 4:  # wdLineSpaceExactly clips tall OLE objects
        paragraph.LineSpacingRule = 3  # wdLineSpaceAtLeast, retaining prior minimum


def place_equation(
    doc: Any, target: Any, generated: Any, options: LayoutOptions, size: float, replace_object: Any
) -> Any:
    """Insert only after conversion/width validation; caller owns one Undo record."""
    if options.layout == "inline":
        old_style = str(target.Paragraphs(1).Range.Style.NameLocal)
        if old_style in (DISPLAY_STYLE, NUMBERED_STYLE):
            paragraph = target.Paragraphs(1)
            if target.Start != paragraph.Range.Start or target.End != paragraph.Range.End - 1:
                raise ValueError(
                    "Choose a body paragraph for inline insertion; use word_format_axmath "
                    "to change an existing managed display/numbered paragraph."
                )
            paragraph.Range.Style = doc.Styles(-1)
            paragraph.Format.Alignment = 0
            paragraph.Format.TabStops.ClearAll()
        inserted = replace_object(doc, target, generated)
        inline_paragraph(inserted)
        return inserted

    if options.layout == "numbered":
        preflight_numbers(doc)
        if options.number_label is None:
            preflight_native_numbering(doc)
    width = check_block_fits(target, generated, options, size)
    para = target.Paragraphs(1).Range
    start = int(target.Start)
    prefix = "\r" if start > para.Start else ""
    suffix = "\r" if target.End < para.End - 1 else ""
    target.Text = prefix + suffix
    location = start + len(prefix)
    # A separate paragraph is now available; surrounding body text is retained.
    # Set a paragraph style BEFORE copying the OLE: applying a linked Word style
    # afterwards can erase the object's native half-point baseline position.
    paragraph = doc.Range(location, location).Paragraphs(1)
    paragraph.Range.Style = _style(
        doc, NUMBERED_STYLE if options.layout == "numbered" else DISPLAY_STYLE
    )
    fmt = paragraph.Format
    fmt.LeftIndent = 0
    fmt.RightIndent = 0
    fmt.FirstLineIndent = 0
    fmt.SpaceBefore = 6
    fmt.SpaceAfter = 6
    fmt.KeepTogether = True
    fmt.KeepWithNext = False
    fmt.LineSpacingRule = 3
    fmt.LineSpacing = max(12.0, size * 1.2)
    fmt.DisableLineHeightGrid = True
    fmt.BaseLineAlignment = 2
    fmt.TabStops.ClearAll()
    if options.layout == "numbered":
        doc.Range(location, location).InsertAfter("\t")
        location += 1
    inserted = replace_object(doc, doc.Range(location, location), generated)
    # Cross-instance XML insertion may replace the paragraph object. Formatting
    # proxies captured before insertion can then refer to the old paragraph.
    paragraph = inserted.Range.Paragraphs(1)
    fmt = paragraph.Format
    if options.layout == "display":
        fmt.Alignment = 1  # wdAlignParagraphCenter
        return inserted

    fmt.Alignment = 0
    fmt.TabStops.ClearAll()
    fmt.TabStops.Add(Position=width / 2, Alignment=1, Leader=0)
    fmt.TabStops.Add(Position=width, Alignment=2, Leader=0)
    end = int(inserted.Range.End)
    suffix_range = doc.Range(end, end)
    suffix_range.InsertAfter("\t" + (options.number_label or ""))
    if options.number_label is None:
        baseline = read_equation_metrics(generated.Range)["baseline_offset_pt"]
        if baseline is None:
            raise RuntimeError("Cannot determine the native equation baseline for numbering.")
        original_size = equation_font_size(generated.Range)
        original_font = generated.Range.Font.Name
        original_format = inserted.Range.ParagraphFormat.Duplicate
        number = insert_native_number(doc, end + 1)
        # The vendor applies its display style. Restore our managed paragraph and
        # the OLE baseline explicitly before reapplying center/right tabs.
        paragraph = inserted.Range.Paragraphs(1)
        paragraph.Range.Style = _style(doc, NUMBERED_STYLE)
        paragraph.Format = original_format
        inserted.Range.Font.Name = original_font
        inserted.Range.Font.Size = original_size
        inserted.Range.Font.Position = baseline
        fmt = paragraph.Format
        fmt.Alignment = 0
        fmt.TabStops.ClearAll()
        fmt.TabStops.Add(Position=width / 2, Alignment=1, Leader=0)
        fmt.TabStops.Add(Position=width, Alignment=2, Leader=0)
    # Formatting the suffix only avoids erasing the OLE's native baseline offset.
    number = doc.Range(end + 1, paragraph.Range.End - 1)
    number.Font.Name = "Times New Roman"
    number.Font.Size = size
    number.Font.Position = 0
    number.Font.Bold = False
    number.Font.Italic = False
    number.Font.Spacing = 0
    number.Font.Scaling = 100
    number.Font.AllCaps = False
    number.Font.SmallCaps = False
    number = doc.Range(end + 1, paragraph.Range.End - 1)
    doc.Bookmarks.Add(Name=NUMBER_BOOKMARK + uuid.uuid4().hex[:20], Range=number)
    refresh_numbers(doc)
    if options.number_label is None:
        # Style changes and field updates can coalesce character runs. Restore the
        # native OLE run last, using the untouched converter object's formatting.
        inserted.Range.Font.Name = original_font
        inserted.Range.Font.Size = original_size
        inserted.Range.Font.Position = baseline
    return inserted

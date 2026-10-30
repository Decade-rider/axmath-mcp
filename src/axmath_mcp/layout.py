"""Word paragraph layout, native baseline preservation and equation numbering."""

from __future__ import annotations

import re
import uuid
from typing import Any

from axmath_mcp.validation import LayoutOptions

DISPLAY_STYLE = "AxMathMCP_Display"
NUMBERED_STYLE = "AxMathMCP_Numbered"
NUMBER_BOOKMARK = "AxMathMCP_N_"
SEQUENCE = "AxMathEquation"
_SEQ_CODE = re.compile(r"^\s*SEQ\s+AxMathEquation(?:\s|$)", re.IGNORECASE)


def is_number_field(field: Any) -> bool:
    return int(field.Type) == 12 and bool(_SEQ_CODE.match(str(field.Code.Text)))


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
            result["number_text"] = str(number.Text)
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


def check_block_fits(target: Any, shape: Any, options: LayoutOptions, size: float) -> float:
    width = available_width(target)
    # Conservative reserve, preventing center/right tabs from colliding. Long
    # equations are rejected before any document edit instead of silently shrunk.
    reserve = 0.0
    if options.layout == "numbered":
        reserve = max(54.0, size * len(options.number_label or "(999999)"))
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
    if options.layout == "display":
        fmt.Alignment = 1  # wdAlignParagraphCenter
        return inserted

    fmt.Alignment = 0
    fmt.TabStops.Add(Position=width / 2, Alignment=1, Leader=0)
    fmt.TabStops.Add(Position=width, Alignment=2, Leader=0)
    end = int(inserted.Range.End)
    suffix_range = doc.Range(end, end)
    suffix_range.InsertAfter("\t" + (options.number_label or "()"))
    # Formatting the suffix only avoids erasing the OLE's native baseline offset.
    number = doc.Range(end + 1, suffix_range.End)
    number.Font.Name = "Times New Roman"
    number.Font.Size = size
    number.Font.Position = 0
    number.Font.Italic = False
    if options.number_label is None:
        field_at = doc.Range(end + 2, end + 2)
        doc.Fields.Add(
            Range=field_at, Type=12, Text=SEQUENCE + r" \* ARABIC", PreserveFormatting=True
        )
    number = doc.Range(end + 1, paragraph.Range.End - 1)
    doc.Bookmarks.Add(Name=NUMBER_BOOKMARK + uuid.uuid4().hex[:20], Range=number)
    refresh_numbers(doc)
    return inserted

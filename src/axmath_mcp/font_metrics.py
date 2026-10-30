"""Read equation run metrics without rewriting native AxMath baseline offsets.

Word's OLE range includes hidden field characters, so Range.Font.Size/Position
can be wdUndefined (9999999). The actual object run in WordOpenXML is unambiguous
and records half-point values. These metrics describe the Word run, not a parser
for AxMath's private embedded equation format.
"""

from __future__ import annotations

import math
from typing import Any
from xml.etree import ElementTree as ET

try:
    from pywintypes import com_error
except ImportError:  # Pure XML readers and tests can also run without Windows COM.
    _RANGE_ERRORS = (AttributeError, TypeError, ValueError, OSError)
else:
    _RANGE_ERRORS = (AttributeError, TypeError, ValueError, OSError, com_error)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
O = "{urn:schemas-microsoft-com:office:office}"
PKG = "{http://schemas.microsoft.com/office/2006/xmlPackage}"


def _empty_metrics(count: int = 0) -> dict[str, Any]:
    return {
        "object_count": count,
        "font_size_pt": None,
        "font_size_source": None,
        "baseline_offset_pt": None,
        "baseline_source": None,
    }


def _half_points(properties: ET.Element | None, name: str) -> float | None:
    if properties is None:
        return None
    element = properties.find(W + name)
    if element is None:
        return None
    try:
        value = float(element.get(W + "val", "")) / 2
    except ValueError:
        return None
    if not math.isfinite(value):
        return None
    if name == "sz" and not 1 <= value <= 1638:
        return None
    if name == "position" and not -1638 <= value <= 1638:
        return None
    return value


def _style_chain(styles: dict[str, ET.Element], style_id: str | None):
    seen = set()
    while style_id and style_id not in seen and style_id in styles:
        seen.add(style_id)
        style = styles[style_id]
        yield style.find(W + "rPr"), "style:" + style_id
        based_on = style.find(W + "basedOn")
        style_id = based_on.get(W + "val") if based_on is not None else None


def equation_run_metrics(xml: str) -> dict[str, Any]:
    """Read one native AxMath object run from document XML or Word's flat OPC XML.

    Inherited character/paragraph styles and document defaults are resolved when
    included. Multiple objects are deliberately ambiguous: no font guess is made.
    Invalid XML yields unknown metrics so callers can fall back to COM properties.
    """
    try:
        root = ET.fromstring(xml)
    except (ET.ParseError, TypeError, ValueError):
        return _empty_metrics()
    document = root if root.tag == W + "document" else None
    styles_root = None
    if root.tag == PKG + "package":
        for part in root.findall(PKG + "part"):
            name = part.get(PKG + "name")
            if name == "/word/document.xml":
                document = part.find(PKG + "xmlData/" + W + "document")
            elif name == "/word/styles.xml":
                styles_root = part.find(PKG + "xmlData/" + W + "styles")
    if document is None:
        return _empty_metrics()

    parents = {child: parent for parent in document.iter() for child in parent}
    objects = []
    for run in document.iter(W + "r"):
        if any(
            str(obj.get("ProgID", "")).lower() == "equation.axmath"
            for obj in run.iter(O + "OLEObject")
        ):
            objects.append(run)
    result = _empty_metrics(len(objects))
    if len(objects) != 1:
        return result

    run = objects[0]
    properties = run.find(W + "rPr")
    candidates = [(properties, "object_run")]
    styles = {}
    default_paragraph = None
    if styles_root is not None:
        for style in styles_root.findall(W + "style"):
            style_id = style.get(W + "styleId")
            if style_id:
                styles[style_id] = style
            if style.get(W + "type") == "paragraph" and style.get(W + "default") in {
                "1",
                "true",
                "on",
            }:
                default_paragraph = style_id
    if properties is not None:
        character_style = properties.find(W + "rStyle")
        if character_style is not None:
            candidates.extend(_style_chain(styles, character_style.get(W + "val")))
    paragraph = parents.get(run)
    while paragraph is not None and paragraph.tag != W + "p":
        paragraph = parents.get(paragraph)
    paragraph_style = None if paragraph is None else paragraph.find(W + "pPr/" + W + "pStyle")
    paragraph_id = (
        paragraph_style.get(W + "val") if paragraph_style is not None else default_paragraph
    )
    candidates.extend(_style_chain(styles, paragraph_id))
    if styles_root is not None:
        candidates.append(
            (
                styles_root.find(W + "docDefaults/" + W + "rPrDefault/" + W + "rPr"),
                "document_defaults",
            )
        )
    for properties, source in candidates:
        if result["font_size_pt"] is None:
            value = _half_points(properties, "sz")
            if value is not None:
                result["font_size_pt"] = value
                result["font_size_source"] = source
        if result["baseline_offset_pt"] is None:
            value = _half_points(properties, "position")
            if value is not None:
                result["baseline_offset_pt"] = value
                result["baseline_source"] = source
    if result["baseline_offset_pt"] is None:
        result["baseline_offset_pt"] = 0.0
        result["baseline_source"] = "default"
    return result


def read_equation_metrics(rng: Any) -> dict[str, Any]:
    """Read only; no character formatting or selection is changed."""
    try:
        return equation_run_metrics(str(rng.WordOpenXML))
    except _RANGE_ERRORS:  # COM may reject WordOpenXML for a stale/disconnected range.
        return _empty_metrics()


def _valid_size(value: Any) -> float | None:
    try:
        size = float(value)
    except (TypeError, ValueError):
        return None
    return size if math.isfinite(size) and 1 <= size <= 1638 else None


def equation_font_size(rng: Any, default: float = 12.0) -> float:
    """Use the object run, then uniform COM size, then inherited recorded size.

    A native equation's internal base size cannot be reconstructed if neither its
    Word run nor styles record it; generation should persist the requested run size.
    """
    metrics = read_equation_metrics(rng)
    if metrics["font_size_source"] == "object_run":
        return float(metrics["font_size_pt"])
    try:
        size = _valid_size(rng.Font.Size)
    except _RANGE_ERRORS:  # A mixed/invalid COM range is not evidence for a font size.
        size = None
    if size is not None:
        return size
    if metrics["font_size_pt"] is not None:
        return float(metrics["font_size_pt"])
    try:
        size = _valid_size(rng.Paragraphs(1).Range.Style.Font.Size)
    except _RANGE_ERRORS:
        size = None
    return size if size is not None else default

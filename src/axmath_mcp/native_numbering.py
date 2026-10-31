"""Invoke installed AxMath numbering macros; never synthesize vendor field codes."""

from contextlib import contextmanager
from typing import Any

FORMAT_PROPERTY = "AMEquationNumber2"
SIMPLE_FORMAT = "(#E1)"


def property_value(doc: Any, name: str) -> Any:
    for prop in doc.CustomDocumentProperties:
        if str(prop.Name) == name:
            return prop.Value
    return None


def set_property(doc: Any, name: str, value: Any) -> None:
    for prop in doc.CustomDocumentProperties:
        if str(prop.Name) == name:
            prop.Value = value
            return
    doc.CustomDocumentProperties.Add(name, False, 4, str(value))


@contextmanager
def temporary_property(doc: Any, name: str, value: str):
    previous = property_value(doc, name)
    set_property(doc, name, value)
    try:
        yield
    finally:
        if previous is None:
            doc.CustomDocumentProperties(name).Delete()
        else:
            set_property(doc, name, previous)


def preflight_native_numbering(doc: Any) -> None:
    fmt = property_value(doc, FORMAT_PROPERTY)
    if fmt is not None and str(fmt) != SIMPLE_FORMAT:
        raise ValueError(
            "This document uses a custom/chapter AxMath number format. "
            "MCP native insertion currently supports (1), (2), ... only; "
            "use AxMath interactively for that format. Existing settings are preserved."
        )
    if not any(str(t.Name).lower() == "axmath.dotm" for t in doc.Application.Templates):
        raise RuntimeError("AxMath.dotm must be loaded in the target Word instance for numbering.")


def insert_native_number(doc: Any, at: int) -> Any:
    """Insert a vendor-generated number at a caret and return its complete range.

    Plain sequence numbering needs no chapter initializer. Suppress that initializer
    only during this call, and defer the vendor's document-wide field update so that
    external links and unrelated fields are never refreshed as a side effect.
    """
    preflight_native_numbering(doc)
    if property_value(doc, FORMAT_PROPERTY) is None:
        set_property(doc, FORMAT_PROPERTY, SIMPLE_FORMAT)
    doc.Activate()
    doc.Range(at, at).Select()
    with (
        temporary_property(doc, "AMEquationSection", "1"),
        temporary_property(doc, "AMDeferFieldUpdate", "1"),
    ):
        doc.Application.Run("AMCEqnNum.InsertEqnNum")
    paragraph = doc.Range(at, at).Paragraphs(1).Range
    number = doc.Range(at, paragraph.End - 1)
    if not any("MACROBUTTON AMMPlaceRM " in str(f.Code.Text) for f in number.Fields):
        raise RuntimeError("AxMath did not create its native equation-number field.")
    # AxMath uses a MACROBUTTON with nested fields in its display text. Its Result
    # can be empty through COM, so always show codes for this specific wrapper.
    for field in number.Fields:
        if "MACROBUTTON AMMPlaceRM " in str(field.Code.Text):
            field.ShowCodes = True
    return number

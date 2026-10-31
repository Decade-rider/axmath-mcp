"""Copy native equation ranges within Word or across isolated Word instances.

Cross-instance FormattedText assignment is not supported by the tested Word COM
interfaces. Flat OPC transfers the embedded storage and run formatting instead.
"""

from __future__ import annotations

import re
from typing import Any
from xml.dom import Node, minidom

PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OFFICE = "urn:schemas-microsoft-com:office:office"
XMLNS = "http://www.w3.org/2000/xmlns/"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_DOCTYPE = re.compile(r"<!\s*DOCTYPE\b", re.IGNORECASE)


def _parse_package(xml: str) -> minidom.Document:
    if not isinstance(xml, str) or _DOCTYPE.search(xml):
        raise ValueError("Expected Word-generated XML without a DOCTYPE declaration.")
    dom = minidom.parseString(xml)
    root = dom.documentElement
    if root.namespaceURI != PKG or root.localName != "package":
        dom.unlink()
        raise ValueError("Expected a WordOpenXML flat OPC package.")
    return dom


def _part(root: Any, name: str) -> Any:
    parts = [
        node
        for node in root.childNodes
        if node.nodeType == Node.ELEMENT_NODE
        and node.namespaceURI == PKG
        and node.localName == "part"
        and node.getAttributeNS(PKG, "name") == name
    ]
    if len(parts) > 1:
        raise ValueError(f"Duplicate Word package part: {name}")
    return parts[0] if parts else None


def _fragment(root: Any) -> Any:
    part = _part(root, "/word/document.xml")
    fragments = [] if part is None else part.getElementsByTagNameNS(WORD, "document")
    if len(fragments) != 1:
        raise ValueError("Expected exactly one Word document fragment.")
    return fragments[0]


def _bindings(element: Any) -> dict[str, str]:
    ancestors = []
    while element is not None and element.nodeType == Node.ELEMENT_NODE:
        ancestors.append(element)
        element = element.parentNode
    result = {}
    for ancestor in reversed(ancestors):
        for index in range(ancestor.attributes.length):
            attribute = ancestor.attributes.item(index)
            if attribute.namespaceURI == XMLNS:
                result[attribute.name] = attribute.value
    return result


def _merge_namespace_context(target: Any, origin: Any) -> None:
    existing = _bindings(target)
    for name, value in _bindings(origin).items():
        if name in existing and existing[name] != value:
            raise ValueError(f"Incompatible namespace binding in Word fragments: {name}")
        if name not in existing:
            target.setAttributeNS(XMLNS, name, value)
    original_ignore = origin.getAttributeNodeNS(MC, "Ignorable")
    if original_ignore is not None:
        target_ignore = target.getAttributeNodeNS(MC, "Ignorable")
        names = target_ignore.value.split() if target_ignore is not None else []
        names.extend(original_ignore.value.split())
        target.setAttributeNS(
            MC,
            target_ignore.name if target_ignore is not None else original_ignore.name,
            " ".join(dict.fromkeys(names)),
        )


def _merge_destination_styles(source_dom: Any, destination_root: Any) -> set[str]:
    source_root = source_dom.documentElement
    source_part = _part(source_root, "/word/styles.xml")
    destination_part = _part(destination_root, "/word/styles.xml")
    if destination_part is None:
        return {
            style.getAttributeNS(WORD, "styleId")
            for style in (
                [] if source_part is None else source_part.getElementsByTagNameNS(WORD, "style")
            )
        }
    merged_part = source_dom.importNode(destination_part, deep=True)
    _merge_namespace_context(merged_part, destination_part)
    roots = merged_part.getElementsByTagNameNS(WORD, "styles")
    if len(roots) != 1:
        raise ValueError("Expected one destination style definition part.")
    merged = roots[0]
    # Importing a subtree does not copy namespace declarations from its ancestors.
    destination_styles = destination_part.getElementsByTagNameNS(WORD, "styles")[0]
    _merge_namespace_context(merged, destination_styles)
    ids = {
        style.getAttributeNS(WORD, "styleId")
        for style in merged.getElementsByTagNameNS(WORD, "style")
    }
    if source_part is not None:
        source_roots = source_part.getElementsByTagNameNS(WORD, "styles")
        if len(source_roots) != 1:
            raise ValueError("Expected one source style definition part.")
        _merge_namespace_context(merged, source_roots[0])
        for style in source_roots[0].getElementsByTagNameNS(WORD, "style"):
            style_id = style.getAttributeNS(WORD, "styleId")
            if style_id not in ids:
                merged.appendChild(source_dom.importNode(style, deep=True))
                ids.add(style_id)
        source_root.replaceChild(merged_part, source_part)
        source_part.unlink()
    else:
        source_root.appendChild(merged_part)
    return ids


def equation_fragment_xml(word_open_xml: str, *, destination_xml: str | None = None) -> str:
    """Transfer one equation using the destination's paragraph style and layout.

    Without a destination, source paragraph/page properties are removed. With a
    destination, its paragraph properties and style definitions take precedence;
    source-only styles remain available to equation runs. Native run properties,
    relationships, media and binary OLE storage remain intact. Existing namespace
    prefixes are retained, including those named only in mc:Ignorable values.
    """
    dom = _parse_package(word_open_xml)
    destination = None
    try:
        root = dom.documentElement
        fragment = _fragment(root)
        paragraphs = fragment.getElementsByTagNameNS(WORD, "p")
        objects = fragment.getElementsByTagNameNS(OFFICE, "OLEObject")
        if (
            len(paragraphs) != 1
            or len(objects) != 1
            or objects[0].getAttribute("ProgID").casefold() != "equation.axmath"
        ):
            raise ValueError(
                "Expected exactly one paragraph containing one native AxMath OLE object."
            )
        # Paragraph properties include alignment, spacing, style, tabs and grid;
        # section properties include paper size/margins. The destination owns all
        # of those. Named bookmarks must not be duplicated into the target file.
        for local_name in ("pPr", "sectPr", "bookmarkStart", "bookmarkEnd"):
            for node in list(fragment.getElementsByTagNameNS(WORD, local_name)):
                node.parentNode.removeChild(node)
                node.unlink()
        if destination_xml is not None:
            destination = _parse_package(destination_xml)
            destination_fragment = _fragment(destination.documentElement)
            destination_paragraphs = destination_fragment.getElementsByTagNameNS(WORD, "p")
            if not destination_paragraphs:
                raise ValueError("Destination paragraph XML contains no paragraph.")
            destination_paragraph = destination_paragraphs[0]
            properties = next(
                (
                    child
                    for child in destination_paragraph.childNodes
                    if child.nodeType == Node.ELEMENT_NODE
                    and child.namespaceURI == WORD
                    and child.localName == "pPr"
                ),
                None,
            )
            _merge_namespace_context(fragment, destination_fragment)
            ids = _merge_destination_styles(dom, destination.documentElement)
            if properties is not None:
                copied = dom.importNode(properties, deep=True)
                _merge_namespace_context(copied, properties)
                for section in list(copied.getElementsByTagNameNS(WORD, "sectPr")):
                    section.parentNode.removeChild(section)
                    section.unlink()
                paragraph_style = copied.getElementsByTagNameNS(WORD, "pStyle")
                if paragraph_style and paragraph_style[0].getAttributeNS(WORD, "val") not in ids:
                    raise ValueError(
                        "Destination paragraph style has no supplied style definition."
                    )
                paragraphs[0].insertBefore(copied, paragraphs[0].firstChild)
        return dom.toxml()
    finally:
        if destination is not None:
            destination.unlink()
        dom.unlink()


def _same_application(target: Any, source: Any) -> bool:
    first, second = target.Application, source.Application
    if first is second:
        return True
    try:
        return first._oleobj_ == second._oleobj_
    except AttributeError:
        return False


def copy_native_range(target: Any, source: Any) -> None:
    """Copy a native equation and keep the destination Range around the result.

    Both ranges are expected to contain only the caller-approved replacement or
    equation. No selection, clipboard, save, or process-lifecycle operation occurs.
    """
    if _same_application(target, source):
        target.FormattedText = source.FormattedText
        return

    fragment = equation_fragment_xml(
        str(source.WordOpenXML),
        destination_xml=str(target.Paragraphs(1).Range.WordOpenXML),
    )
    source_text = str(source.Text)
    if "\r" in source_text or "\r" in str(target.Text):
        raise ValueError(
            "Native equation transfer ranges must exclude paragraph marks; no equation was inserted."
        )
    document = target.Document
    start, end = int(target.Start), int(target.End)
    before = int(document.Content.End)
    if not 0 <= start <= end <= before:
        raise ValueError("The destination Word range is invalid; no equation was inserted.")
    target.InsertXML(fragment)
    after = int(document.Content.End)
    new_end = end + after - before
    if not 0 <= start <= new_end <= after:
        raise RuntimeError(
            "Word XML insertion changed the range unexpectedly. Inspect the document; "
            "do not retry the insertion blindly."
        )
    # InsertXML may leave Range collapsed even after inserting a 28-position OLE
    # field, unlike FormattedText. Do not assume one visible U+0001 is one position.
    target.SetRange(start, new_end)
    inserted_text = str(target.Text)
    if int(target.InlineShapes.Count) != 1 or inserted_text not in {
        source_text,
        source_text + "\r",
    }:
        raise RuntimeError(
            "Word XML insertion produced unexpected text, paragraphs or objects. "
            "Inspect the document; do not retry blindly."
        )
    if inserted_text == source_text + "\r":
        # Only a mark added inside the rebound insertion span is eligible. Both
        # original ranges excluded paragraph marks, so the original paragraph's
        # final mark is outside this range and cannot be removed here.
        trailing = document.Range(new_end - 1, new_end)
        if str(trailing.Text) != "\r":
            raise RuntimeError(
                "Could not identify the inserted paragraph mark; inspect the document."
            )
        trailing.Delete()
        target.SetRange(start, new_end - 1)
        if (
            int(document.Content.End) != after - 1
            or str(target.Text) != source_text
            or int(target.InlineShapes.Count) != 1
        ):
            raise RuntimeError(
                "Word's inserted paragraph mark could not be removed cleanly. "
                "Inspect the document before another insertion."
            )

"""Extract one Word-generated native AxMath OLE storage, without vendor binaries."""

from __future__ import annotations

import base64
import re
import zipfile
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET

PKG = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
OFFICE = "{urn:schemas-microsoft-com:office:office}"
CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
PPT = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
REL_TYPES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/",
)


def extract_axmath_storage(word_open_xml: str) -> bytes:
    root = ET.fromstring(word_open_xml)
    objects = list(root.iter(OFFICE + "OLEObject"))
    if len(objects) != 1 or objects[0].get("ProgID", "").lower() != "equation.axmath":
        raise ValueError("Expected exactly one native AxMath OLE object in the temporary range.")
    payloads = []
    for part in root.findall(PKG + "part"):
        name = part.get(PKG + "name", "")
        if name.startswith("/word/embeddings/") and name.endswith(".bin"):
            data = part.find(PKG + "binaryData")
            if data is None or not data.text:
                raise ValueError("The embedded AxMath storage is missing.")
            payloads.append(base64.b64decode("".join(data.text.split()), validate=True))
    if len(payloads) != 1 or not payloads[0].startswith(CFB_SIGNATURE):
        raise ValueError(
            "Expected one valid OLE compound storage, not a picture or package wrapper."
        )
    return payloads[0]


def _read_package_part(package: zipfile.ZipFile, name: str) -> bytes:
    if sum(info.filename == name for info in package.infolist()) != 1:
        raise ValueError(f"Missing or ambiguous PowerPoint package part: {name}")
    return package.read(name)


def _read_package_xml(package: zipfile.ZipFile, name: str) -> ET.Element:
    data = _read_package_part(package, name)
    if re.search(rb"<!\s*DOCTYPE\b", data, re.IGNORECASE):
        raise ValueError("PowerPoint package XML must not contain a DOCTYPE declaration.")
    return ET.fromstring(data)


def _internal_target(owner: str, target: str) -> str:
    """Resolve normal OPC ../ paths without permitting filesystem/root escape."""
    decoded = unquote(target)
    parts = urlsplit(decoded)
    if (
        not decoded
        or parts.scheme
        or parts.netloc
        or parts.query
        or parts.fragment
        or decoded.startswith(("/", "\\"))
        or "\\" in decoded
        or ":" in decoded
        or any(ord(char) < 32 for char in decoded)
    ):
        raise ValueError("Expected a relative internal PowerPoint relationship target.")
    stack = list(PurePosixPath(owner).parent.parts)
    for component in decoded.split("/"):
        if component in ("", "."):
            continue
        if component == "..":
            if not stack:
                raise ValueError("PowerPoint relationship escapes the package root.")
            stack.pop()
        else:
            stack.append(component)
    if not stack:
        raise ValueError("PowerPoint relationship has no package-part target.")
    return "/".join(stack)


def _related_part(package: zipfile.ZipFile, owner: str, relationship_id: str, kind: str) -> str:
    location = PurePosixPath(owner)
    relations_path = str(location.parent / "_rels" / (location.name + ".rels"))
    root = _read_package_xml(package, relations_path)
    matches = [node for node in root.findall(REL) if node.get("Id") == relationship_id]
    if not relationship_id or len(matches) != 1:
        raise ValueError("PowerPoint relationship is missing or ambiguous.")
    relationship = matches[0]
    if relationship.get("TargetMode", "Internal").casefold() != "internal":
        raise ValueError("External PowerPoint OLE/slide relationships are not supported.")
    if relationship.get("Type") not in {prefix + kind for prefix in REL_TYPES}:
        raise ValueError(f"PowerPoint relationship is not of type {kind}.")
    return _internal_target(owner, relationship.get("Target", ""))


def extract_powerpoint_storage(path: str | Path, slide_id: int, shape_id: int) -> bytes:
    """Read a specific native equation from a saved PPTX/PPTM copy.

    SlideID and shape Id are stable Office identifiers, not collection indices or
    filenames. Choice/Fallback OLE representations must agree on one embedded
    relationship. Nothing is extracted to disk and no Office process is invoked.
    """
    if any(
        isinstance(value, bool) or not isinstance(value, int) or value <= 0
        for value in (slide_id, shape_id)
    ):
        raise ValueError("slide_id and shape_id must be positive integer Office IDs.")
    with zipfile.ZipFile(path) as package:
        presentation_part = "ppt/presentation.xml"
        presentation = _read_package_xml(package, presentation_part)
        slides = [
            node
            for node in presentation.findall(PPT + "sldIdLst/" + PPT + "sldId")
            if node.get("id") == str(slide_id)
        ]
        if len(slides) != 1:
            raise ValueError("SlideID was not found uniquely in the saved presentation.")
        slide_part = _related_part(package, presentation_part, slides[0].get(REL_ID, ""), "slide")
        if not slide_part.startswith("ppt/slides/") or not slide_part.endswith(".xml"):
            raise ValueError("Slide relationship points outside PowerPoint slide parts.")
        slide = _read_package_xml(package, slide_part)
        tree = slide.find(PPT + "cSld/" + PPT + "spTree")
        if tree is None:
            raise ValueError("Saved slide has no shape tree.")
        frames = []
        for frame in tree.findall(PPT + "graphicFrame"):
            identity = frame.find(PPT + "nvGraphicFramePr/" + PPT + "cNvPr")
            if identity is not None and identity.get("id") == str(shape_id):
                frames.append(frame)
        if len(frames) != 1:
            raise ValueError("Shape ID was not found uniquely as a top-level OLE frame.")
        objects = list(frames[0].iter(PPT + "oleObj"))
        if not objects or any(
            node.get("progId", "").casefold() != "equation.axmath"
            or node.find(PPT + "embed") is None
            or node.find(PPT + "link") is not None
            for node in objects
        ):
            raise ValueError("Saved shape is not an embedded native Equation.AxMath object.")
        references = {node.get(REL_ID, "") for node in objects}
        if len(references) != 1 or "" in references:
            raise ValueError("PowerPoint OLE compatibility representations disagree.")
        storage_part = _related_part(package, slide_part, references.pop(), "oleObject")
        if not storage_part.startswith("ppt/embeddings/") or not storage_part.endswith(".bin"):
            raise ValueError("OLE relationship points outside PowerPoint embedded storage.")
        storage = _read_package_part(package, storage_part)
        if not storage.startswith(CFB_SIGNATURE):
            raise ValueError("PowerPoint equation does not contain an OLE compound storage.")
        return storage

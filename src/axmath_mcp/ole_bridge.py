"""Extract one Word-generated native AxMath OLE storage, without vendor binaries."""

from __future__ import annotations

import base64
from xml.etree import ElementTree as ET

PKG = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
OFFICE = "{urn:schemas-microsoft-com:office:office}"
CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


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

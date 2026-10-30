"""The native storage bridge must reject pictures, wrappers and ambiguity."""

import base64
from xml.etree import ElementTree as ET

import pytest

from axmath_mcp.ole_bridge import CFB_SIGNATURE, extract_axmath_storage


def package(*, progids=("Equation.AxMath",), storages=None):
    if storages is None:
        storages = [
            (
                "/word/embeddings/oleObject1.bin",
                base64.b64encode(CFB_SIGNATURE + bytes(504)).decode(),
            )
        ]
    objects = "".join(f'<o:OLEObject ProgID="{p}"/>' for p in progids)
    parts = "".join(
        f'<pkg:part pkg:name="{name}">'
        + (f"<pkg:binaryData>{data}</pkg:binaryData>" if data is not None else "")
        + "</pkg:part>"
        for name, data in storages
    )
    return (
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" '
        'xmlns:o="urn:schemas-microsoft-com:office:office">'
        f'<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>{objects}</pkg:xmlData></pkg:part>'
        f"{parts}</pkg:package>"
    )


def test_extracts_one_declared_axmath_compound_storage():
    assert extract_axmath_storage(package()) == CFB_SIGNATURE + bytes(504)


def test_object_progid_is_case_insensitive():
    assert extract_axmath_storage(package(progids=("equation.axmath",))).startswith(CFB_SIGNATURE)


@pytest.mark.parametrize(
    "progids", [(), ("Equation.AxMath", "Equation.AxMath"), ("Word.Document.12",), ("Package",)]
)
def test_rejects_missing_multiple_and_container_wrapper_objects(progids):
    with pytest.raises(ValueError, match="exactly one native AxMath"):
        extract_axmath_storage(package(progids=progids))


@pytest.mark.parametrize("data", [None, ""])
def test_rejects_missing_embedded_payload(data):
    with pytest.raises(ValueError, match="storage is missing"):
        extract_axmath_storage(package(storages=[("/word/embeddings/oleObject1.bin", data)]))


def test_rejects_multiple_storages_even_with_one_declared_object():
    encoded = base64.b64encode(CFB_SIGNATURE + bytes(504)).decode()
    with pytest.raises(ValueError, match="one valid OLE"):
        extract_axmath_storage(
            package(
                storages=[
                    ("/word/embeddings/oleObject1.bin", encoded),
                    ("/word/embeddings/oleObject2.bin", encoded),
                ]
            )
        )


@pytest.mark.parametrize("payload", [b"PNG image", b"", b"{\\rtf1 plain text}"])
def test_picture_or_text_payload_is_not_reported_as_native_ole(payload):
    encoded = base64.b64encode(payload).decode()
    with pytest.raises(ValueError):
        extract_axmath_storage(package(storages=[("/word/embeddings/oleObject1.bin", encoded)]))


def test_rejects_invalid_base64():
    with pytest.raises(ValueError):
        extract_axmath_storage(
            package(storages=[("/word/embeddings/oleObject1.bin", "not!!!base64")])
        )


def test_base64_line_wrapping_does_not_change_native_bytes():
    data = CFB_SIGNATURE + bytes(range(256))
    encoded = base64.b64encode(data).decode()
    wrapped = "\n ".join(encoded[i : i + 20] for i in range(0, len(encoded), 20))
    assert (
        extract_axmath_storage(package(storages=[("/word/embeddings/oleObject1.bin", wrapped)]))
        == data
    )


def test_ignores_non_embedding_binary_parts():
    native = base64.b64encode(CFB_SIGNATURE + bytes(504)).decode()
    result = extract_axmath_storage(
        package(
            storages=[
                ("/word/media/image1.png", base64.b64encode(b"PNG preview").decode()),
                ("/word/embeddings/oleObject1.bin", native),
            ]
        )
    )
    assert result == CFB_SIGNATURE + bytes(504)


def test_rejects_malformed_xml():
    with pytest.raises(ET.ParseError):
        extract_axmath_storage("<not-closed>")

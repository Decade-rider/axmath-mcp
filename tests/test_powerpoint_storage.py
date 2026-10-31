"""Resolve native equation storage using real Office IDs and OPC relationships."""

import zipfile
from pathlib import Path

import pytest

from axmath_mcp.ole_bridge import CFB_SIGNATURE, extract_powerpoint_storage

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
PAYLOAD = CFB_SIGNATURE + bytes(range(256))


def fixture_parts(
    *,
    ole_target="../embeddings/equation42.bin",
    target_mode="",
    progid="Equation.AxMath",
    fallback_ref="rEquation",
    grouped=False,
):
    objects = (
        f'<mc:AlternateContent><mc:Choice Requires="v"><p:oleObj progId="{progid}" r:id="rEquation"><p:embed/></p:oleObj></mc:Choice>'
        f'<mc:Fallback><p:oleObj progId="{progid}" r:id="{fallback_ref}"><p:embed/></p:oleObj></mc:Fallback></mc:AlternateContent>'
    )
    frame = f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="17" name="formula"/></p:nvGraphicFramePr>{objects}</p:graphicFrame>'
    if grouped:
        frame = "<p:grpSp>" + frame + "</p:grpSp>"
    mode = f' TargetMode="{target_mode}"' if target_mode else ""
    return {
        "ppt/presentation.xml": f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="902" r:id="rSlide"/><p:sldId id="256" r:id="rOther"/></p:sldIdLst></p:presentation>',
        "ppt/_rels/presentation.xml.rels": f'<Relationships xmlns="{PACKAGE_RELS}"><Relationship Id="rSlide" Type="{R}/slide" Target="slides/custom77.xml"/><Relationship Id="rOther" Type="{R}/slide" Target="slides/slide1.xml"/></Relationships>',
        "ppt/slides/custom77.xml": f'<p:sld xmlns:p="{P}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree>{frame}</p:spTree></p:cSld></p:sld>',
        "ppt/slides/slide1.xml": f'<p:sld xmlns:p="{P}"><p:cSld><p:spTree/></p:cSld></p:sld>',
        "ppt/slides/_rels/custom77.xml.rels": f'<Relationships xmlns="{PACKAGE_RELS}"><Relationship Id="rEquation" Type="{R}/oleObject" Target="{ole_target}"{mode}/></Relationships>',
        "ppt/embeddings/equation42.bin": PAYLOAD,
    }


def write_package(tmp_path: Path, parts):
    path = tmp_path / "copy.pptx"
    with zipfile.ZipFile(path, "w") as package:
        for name, value in parts.items():
            package.writestr(name, value)
    return path


def test_reads_stable_slide_and_shape_ids_independent_of_order_and_filename(tmp_path):
    path = write_package(tmp_path, fixture_parts())
    assert extract_powerpoint_storage(path, 902, 17) == PAYLOAD


def test_does_not_treat_slide_index_as_slide_id(tmp_path):
    path = write_package(tmp_path, fixture_parts())
    with pytest.raises(ValueError, match="SlideID"):
        extract_powerpoint_storage(path, 1, 17)
    with pytest.raises(ValueError, match="top-level"):
        extract_powerpoint_storage(path, 256, 17)


@pytest.mark.parametrize(
    "slide_id,shape_id", [(True, 17), (902, False), (0, 17), (902, -1), ("902", 17)]
)
def test_rejects_invalid_office_ids(tmp_path, slide_id, shape_id):
    with pytest.raises(ValueError, match="positive integer"):
        extract_powerpoint_storage(tmp_path / "unused.pptx", slide_id, shape_id)


def test_grouped_equation_is_not_mistaken_for_top_level_shape(tmp_path):
    path = write_package(tmp_path, fixture_parts(grouped=True))
    with pytest.raises(ValueError, match="top-level"):
        extract_powerpoint_storage(path, 902, 17)


@pytest.mark.parametrize("progid", ["Word.Document.12", "Package", "Equation.3"])
def test_rejects_non_axmath_wrappers(tmp_path, progid):
    path = write_package(tmp_path, fixture_parts(progid=progid))
    with pytest.raises(ValueError, match="native Equation.AxMath"):
        extract_powerpoint_storage(path, 902, 17)


def test_choice_fallback_must_reference_identical_storage(tmp_path):
    path = write_package(tmp_path, fixture_parts(fallback_ref="rDifferent"))
    with pytest.raises(ValueError, match="representations disagree"):
        extract_powerpoint_storage(path, 902, 17)


def test_linked_object_is_rejected_even_with_internal_target(tmp_path):
    parts = fixture_parts()
    parts["ppt/slides/custom77.xml"] = parts["ppt/slides/custom77.xml"].replace(
        "<p:embed/>", "<p:link/>"
    )
    path = write_package(tmp_path, parts)
    with pytest.raises(ValueError, match="embedded native"):
        extract_powerpoint_storage(path, 902, 17)


def test_external_relationship_is_never_followed(tmp_path):
    path = write_package(tmp_path, fixture_parts(target_mode="External"))
    with pytest.raises(ValueError, match="External"):
        extract_powerpoint_storage(path, 902, 17)


@pytest.mark.parametrize(
    "target",
    [
        "https://example.com/formula.bin",
        "file:///C:/formula.bin",
        "//server/formula.bin",
        "C:/formula.bin",
        "..\\embeddings\\equation42.bin",
        "../../../outside.bin",
        "%2e%2e/%2e%2e/%2e%2e/outside.bin",
        "../media/equation42.bin",
        "../embeddings/equation42.bin#fragment",
    ],
)
def test_rejects_external_or_escaping_targets(tmp_path, target):
    path = write_package(tmp_path, fixture_parts(ole_target=target))
    with pytest.raises(ValueError):
        extract_powerpoint_storage(path, 902, 17)


def test_missing_storage_is_reported_without_path_extraction(tmp_path):
    parts = fixture_parts()
    del parts["ppt/embeddings/equation42.bin"]
    with pytest.raises(ValueError, match="Missing or ambiguous"):
        extract_powerpoint_storage(write_package(tmp_path, parts), 902, 17)


def test_storage_must_have_compound_file_signature(tmp_path):
    parts = fixture_parts()
    parts["ppt/embeddings/equation42.bin"] = b"PNG preview is not editable storage"
    with pytest.raises(ValueError, match="compound storage"):
        extract_powerpoint_storage(write_package(tmp_path, parts), 902, 17)


def test_duplicate_relationship_ids_are_ambiguous(tmp_path):
    parts = fixture_parts()
    key = "ppt/slides/_rels/custom77.xml.rels"
    parts[key] = parts[key].replace(
        "</Relationships>",
        f'<Relationship Id="rEquation" Type="{R}/oleObject" Target="../embeddings/other.bin"/></Relationships>',
    )
    with pytest.raises(ValueError, match="ambiguous"):
        extract_powerpoint_storage(write_package(tmp_path, parts), 902, 17)


def test_wrong_relationship_type_is_rejected(tmp_path):
    parts = fixture_parts()
    key = "ppt/slides/_rels/custom77.xml.rels"
    parts[key] = parts[key].replace(R + "/oleObject", R + "/image")
    with pytest.raises(ValueError, match="type oleObject"):
        extract_powerpoint_storage(write_package(tmp_path, parts), 902, 17)


def test_duplicate_matching_graphic_frames_are_rejected(tmp_path):
    parts = fixture_parts()
    key = "ppt/slides/custom77.xml"
    extra = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="17"/></p:nvGraphicFramePr></p:graphicFrame>'
    parts[key] = parts[key].replace("</p:spTree>", extra + "</p:spTree>")
    with pytest.raises(ValueError, match="uniquely"):
        extract_powerpoint_storage(write_package(tmp_path, parts), 902, 17)

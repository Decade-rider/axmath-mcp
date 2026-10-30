"""Native OLE fields must not turn mixed COM font values into a 12 pt reset."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from axmath_mcp.font_metrics import equation_font_size, equation_run_metrics

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"
PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"


def _document(properties="", *, prefix="", suffix="", paragraph_properties=""):
    return (
        f'<w:document xmlns:w="{W}" xmlns:o="{O}"><w:body><w:p>'
        f"<w:pPr>{paragraph_properties}</w:pPr>{prefix}"
        f"<w:r><w:rPr>{properties}</w:rPr><w:object>"
        '<o:OLEObject ProgID="Equation.AxMath"/></w:object></w:r>'
        f"{suffix}</w:p></w:body></w:document>"
    )


def _package(document, styles):
    return (
        f'<pkg:package xmlns:pkg="{PKG}">'
        f'<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>{document}</pkg:xmlData></pkg:part>'
        f'<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData>'
        f'<w:styles xmlns:w="{W}">{styles}</w:styles></pkg:xmlData></pkg:part></pkg:package>'
    )


def _range(xml, size=9999999, paragraph_size=9999999):
    return SimpleNamespace(
        WordOpenXML=xml,
        Font=SimpleNamespace(Size=size),
        Paragraphs=Mock(
            return_value=SimpleNamespace(
                Range=SimpleNamespace(
                    Style=SimpleNamespace(Font=SimpleNamespace(Size=paragraph_size))
                )
            )
        ),
    )


def test_object_run_font_wins_over_hidden_field_and_undefined_com_values():
    hidden = '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
    xml = _document('<w:sz w:val="28"/><w:position w:val="-21"/>', prefix=hidden)
    assert equation_font_size(_range(xml)) == 14.0
    metrics = equation_run_metrics(xml)
    assert metrics["baseline_offset_pt"] == -10.5
    assert metrics["font_size_source"] == "object_run"


def test_native_half_point_size_is_not_rounded_to_integer():
    xml = _document('<w:sz w:val="25"/><w:position w:val="-11"/>')
    assert equation_font_size(_range(xml, size=12)) == 12.5
    assert equation_run_metrics(xml)["baseline_offset_pt"] == -5.5


def test_object_run_does_not_accidentally_read_a_number_suffix_size():
    suffix = '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>(1)</w:t></w:r>'
    xml = _document('<w:sz w:val="32"/>', suffix=suffix)
    assert equation_font_size(_range(xml)) == 16


def test_inherited_style_chain_is_resolved_with_cycle_protection():
    document = _document(paragraph_properties='<w:pStyle w:val="Formula"/>')
    styles = (
        '<w:style w:type="paragraph" w:styleId="Formula"><w:basedOn w:val="Body"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Formula"/>'
        '<w:rPr><w:sz w:val="28"/></w:rPr></w:style>'
    )
    assert equation_font_size(_range(_package(document, styles))) == 14


def test_character_style_overrides_paragraph_style():
    document = _document(
        '<w:rStyle w:val="Math"/>', paragraph_properties='<w:pStyle w:val="Body"/>'
    )
    styles = (
        '<w:style w:type="paragraph" w:styleId="Body"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="Math"><w:rPr><w:sz w:val="28"/></w:rPr></w:style>'
    )
    assert equation_font_size(_range(_package(document, styles))) == 14


def test_localized_normal_style_id_is_not_assumed_to_be_normal():
    styles = '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:rPr><w:sz w:val="28"/></w:rPr></w:style>'
    assert equation_font_size(_range(_package(_document(), styles))) == 14


def test_uniform_com_size_is_preferred_to_inherited_xml_when_object_size_is_absent():
    defaults = '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
    xml = _package(_document(), defaults)
    assert equation_font_size(_range(xml, size=14)) == 14
    assert equation_font_size(_range(xml)) == 10.5


def test_multiple_objects_do_not_produce_an_arbitrary_first_object_font():
    other = '<w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:object><o:OLEObject ProgID="Equation.AxMath"/></w:object></w:r>'
    xml = _document('<w:sz w:val="28"/>', suffix=other)
    assert equation_run_metrics(xml)["object_count"] == 2
    assert equation_font_size(_range(xml, paragraph_size=11)) == 11


@pytest.mark.parametrize("size", ["9999999", "NaN", "-1", "0", "3278"])
def test_invalid_recorded_font_sizes_are_not_used(size):
    assert equation_font_size(_range(_document(f'<w:sz w:val="{size}"/>'), size=13)) == 13


def test_malformed_xml_and_com_undefined_fall_back_to_paragraph_style():
    assert equation_font_size(_range("<broken", paragraph_size=15)) == 15


def test_plain_text_range_uses_uniform_font_size_without_needing_an_equation():
    xml = f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body></w:document>'
    assert equation_font_size(_range(xml, size=11.5)) == 11.5


def test_missing_position_on_a_known_equation_means_default_baseline():
    assert equation_run_metrics(_document('<w:sz w:val="28"/>'))["baseline_offset_pt"] == 0

"""Protect the OLE payload, paragraph boundaries and range rebinding on transfer."""

from types import SimpleNamespace
from unittest.mock import Mock
from xml.dom import minidom

import pytest

from axmath_mcp.word_transfer import OFFICE, PKG, WORD, copy_native_range, equation_fragment_xml

XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="{PKG}">
 <pkg:part pkg:name="/word/document.xml"><pkg:xmlData>
  <w:document xmlns:w="{WORD}" xmlns:o="{OFFICE}"
   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
   xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
   xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">
   <w:body><w:p w14:paraId="12345678">
    <w:pPr><w:pStyle w:val="SourceDisplay"/><w:jc w:val="center"/><w:tabs>
     <w:tab w:val="right" w:pos="9000"/></w:tabs><w:sectPr><w:pgMar w:left="360"/></w:sectPr></w:pPr>
    <w:bookmarkStart w:id="7" w:name="DoNotDuplicate"/>
    <w:r><w:rPr><w:sz w:val="28"/><w:position w:val="-21"/></w:rPr>
     <w:object><o:OLEObject ProgID="Equation.AxMath" r:id="rId1"/></w:object></w:r>
    <w:bookmarkEnd w:id="7"/>
   </w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body>
  </w:document>
 </pkg:xmlData></pkg:part>
 <pkg:part pkg:name="/word/styles.xml"><pkg:xmlData>
  <w:styles xmlns:w="{WORD}"><w:style w:styleId="KeepStyle">
   <w:pPr><w:jc w:val="right"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr>
  </w:style></w:styles>
 </pkg:xmlData></pkg:part>
 <pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
   <Relationship Id="rId1" Type="oleObject" Target="embeddings/oleObject1.bin"/>
  </Relationships>
 </pkg:xmlData></pkg:part>
 <pkg:part pkg:name="/word/embeddings/oleObject1.bin"><pkg:binaryData>0M8R4KGxGuE=
  AAECAw==</pkg:binaryData></pkg:part>
</pkg:package>'''

DESTINATION_XML = f'''<pkg:package xmlns:pkg="{PKG}">
 <pkg:part pkg:name="/word/document.xml"><pkg:xmlData>
  <w:document xmlns:w="{WORD}"
   xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
   xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15">
   <w:body><w:p><w:pPr><w:pStyle w:val="TargetNumbered"/>
    <w:jc w:val="left"/><w:ind w:left="120" w:firstLine="0"/>
    <w:tabs><w:tab w:val="center" w:pos="4500"/><w:tab w:val="right" w:pos="9000"/></w:tabs>
    <w:spacing w:before="120" w:after="120"/><w:snapToGrid w:val="0"/>
    <w:sectPr><w:pgMar w:left="1440"/></w:sectPr></w:pPr>
    <w:r><w:t>Existing destination text</w:t></w:r>
   </w:p><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:body>
  </w:document>
 </pkg:xmlData></pkg:part>
 <pkg:part pkg:name="/word/styles.xml"><pkg:xmlData>
  <w:styles xmlns:w="{WORD}">
   <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="28"/></w:rPr></w:rPrDefault></w:docDefaults>
   <w:style w:styleId="TargetNumbered"><w:name w:val="AxMathMCP_Numbered"/>
    <w:pPr><w:jc w:val="left"/></w:pPr></w:style>
   <w:style w:styleId="KeepStyle"><w:rPr><w:sz w:val="32"/></w:rPr></w:style>
  </w:styles>
 </pkg:xmlData></pkg:part>
</pkg:package>'''


def _fragment(dom):
    part = next(
        p
        for p in dom.getElementsByTagNameNS(PKG, "part")
        if p.getAttributeNS(PKG, "name") == "/word/document.xml"
    )
    return part.getElementsByTagNameNS(WORD, "document")[0]


def test_transfer_strips_only_source_document_layout_and_bookmarks():
    result = equation_fragment_xml(XML)
    with minidom.parseString(result) as dom:
        fragment = _fragment(dom)
        for name in ("pPr", "sectPr", "bookmarkStart", "bookmarkEnd"):
            assert not fragment.getElementsByTagNameNS(WORD, name)
        assert len(fragment.getElementsByTagNameNS(WORD, "p")) == 1
        assert len(fragment.getElementsByTagNameNS(OFFICE, "OLEObject")) == 1
        # Style definitions are separate package content, not target layout.
        assert len(dom.getElementsByTagNameNS(WORD, "pPr")) == 1


def test_run_size_half_point_offset_relationships_and_binary_text_survive():
    before = minidom.parseString(XML)
    after = minidom.parseString(equation_fragment_xml(XML))
    try:
        for tag in ("sz", "position"):
            source = _fragment(before).getElementsByTagNameNS(WORD, tag)[0]
            target = _fragment(after).getElementsByTagNameNS(WORD, tag)[0]
            assert target.toxml() == source.toxml()
        for name in (
            "/word/styles.xml",
            "/word/_rels/document.xml.rels",
            "/word/embeddings/oleObject1.bin",
        ):
            old = next(
                p
                for p in before.getElementsByTagNameNS(PKG, "part")
                if p.getAttributeNS(PKG, "name") == name
            )
            new = next(
                p
                for p in after.getElementsByTagNameNS(PKG, "part")
                if p.getAttributeNS(PKG, "name") == name
            )
            assert new.toxml() == old.toxml()
    finally:
        before.unlink()
        after.unlink()


def test_mc_ignorable_prefixes_are_not_renamed_or_dropped():
    result = equation_fragment_xml(XML)
    assert 'mc:Ignorable="w14"' in result
    assert 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' in result
    assert 'w14:paraId="12345678"' in result
    assert '<o:OLEObject ProgID="Equation.AxMath" r:id="rId1"' in result
    assert '<?mso-application progid="Word.Document"?>' in result


def test_destination_paragraph_properties_are_supplied_before_insert():
    result = equation_fragment_xml(XML, destination_xml=DESTINATION_XML)
    with minidom.parseString(result) as dom:
        fragment = _fragment(dom)
        properties = fragment.getElementsByTagNameNS(WORD, "pPr")[0]
        assert (
            properties.getElementsByTagNameNS(WORD, "pStyle")[0].getAttributeNS(WORD, "val")
            == "TargetNumbered"
        )
        assert (
            properties.getElementsByTagNameNS(WORD, "jc")[0].getAttributeNS(WORD, "val") == "left"
        )
        assert (
            properties.getElementsByTagNameNS(WORD, "ind")[0].getAttributeNS(WORD, "left") == "120"
        )
        tabs = properties.getElementsByTagNameNS(WORD, "tab")
        assert [(t.getAttributeNS(WORD, "val"), t.getAttributeNS(WORD, "pos")) for t in tabs] == [
            ("center", "4500"),
            ("right", "9000"),
        ]
        assert not fragment.getElementsByTagNameNS(WORD, "sectPr")
        assert not fragment.getElementsByTagNameNS(WORD, "bookmarkStart")
        assert (
            fragment.getElementsByTagNameNS(WORD, "position")[0].getAttributeNS(WORD, "val")
            == "-21"
        )
        assert fragment.getElementsByTagNameNS(WORD, "sz")[0].getAttributeNS(WORD, "val") == "28"
        assert "Existing destination text" not in fragment.toxml()


def test_destination_style_definitions_win_but_source_only_styles_survive():
    source = XML.replace(
        "</w:styles>",
        '<w:style w:styleId="SourceOnly"><w:rPr><w:i/></w:rPr></w:style></w:styles>',
    )
    result = equation_fragment_xml(source, destination_xml=DESTINATION_XML)
    with minidom.parseString(result) as dom:
        part = next(
            p
            for p in dom.getElementsByTagNameNS(PKG, "part")
            if p.getAttributeNS(PKG, "name") == "/word/styles.xml"
        )
        definitions = {
            s.getAttributeNS(WORD, "styleId"): s for s in part.getElementsByTagNameNS(WORD, "style")
        }
        assert set(definitions) == {"TargetNumbered", "KeepStyle", "SourceOnly"}
        assert (
            definitions["KeepStyle"]
            .getElementsByTagNameNS(WORD, "sz")[0]
            .getAttributeNS(WORD, "val")
            == "32"
        )
        assert definitions["SourceOnly"].getElementsByTagNameNS(WORD, "i")
        assert len(part.getElementsByTagNameNS(WORD, "docDefaults")) == 1
        assert (
            len(
                [
                    p
                    for p in dom.getElementsByTagNameNS(PKG, "part")
                    if p.getAttributeNS(PKG, "name") == "/word/styles.xml"
                ]
            )
            == 1
        )


def test_destination_merge_keeps_original_native_relationship_and_binary_payload():
    result = equation_fragment_xml(XML, destination_xml=DESTINATION_XML)
    with minidom.parseString(XML) as before, minidom.parseString(result) as after:
        for name in ("/word/_rels/document.xml.rels", "/word/embeddings/oleObject1.bin"):
            old = next(
                p
                for p in before.getElementsByTagNameNS(PKG, "part")
                if p.getAttributeNS(PKG, "name") == name
            )
            new = next(
                p
                for p in after.getElementsByTagNameNS(PKG, "part")
                if p.getAttributeNS(PKG, "name") == name
            )
            assert old.toxml() == new.toxml()


def test_destination_namespace_prefixes_and_ignorable_values_are_merged():
    result = equation_fragment_xml(XML, destination_xml=DESTINATION_XML)
    with minidom.parseString(result) as dom:
        fragment = _fragment(dom)
        mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
        assert set(fragment.getAttributeNS(mc, "Ignorable").split()) == {"w14", "w15"}
        assert fragment.getAttribute("xmlns:w14")
        assert fragment.getAttribute("xmlns:w15")
        assert len(dom.getElementsByTagNameNS(OFFICE, "OLEObject")) == 1


def test_missing_destination_style_is_rejected_before_copy():
    bad = DESTINATION_XML.replace('w:styleId="TargetNumbered"', 'w:styleId="OtherStyle"')
    with pytest.raises(ValueError, match="no supplied style definition"):
        equation_fragment_xml(XML, destination_xml=bad)


def test_destination_doctype_is_rejected():
    with pytest.raises(ValueError, match="DOCTYPE"):
        equation_fragment_xml(XML, destination_xml="<!DOCTYPE x>" + DESTINATION_XML)


@pytest.mark.parametrize(
    "source",
    [
        XML.replace('ProgID="Equation.AxMath"', 'ProgID="Word.Document.12"'),
        XML.replace("</w:p>", "</w:p><w:p/>"),
        XML.replace("</w:object>", '<o:OLEObject ProgID="Equation.AxMath"/></w:object>'),
        XML.replace('/word/document.xml"', '/word/other.xml"'),
        '<w:document xmlns:w="' + WORD + '"/>',
    ],
)
def test_ambiguous_or_non_native_fragments_are_rejected(source):
    with pytest.raises(ValueError):
        equation_fragment_xml(source)


def test_doctype_is_rejected_before_entities_can_be_used():
    with pytest.raises(ValueError, match="DOCTYPE"):
        equation_fragment_xml('<!DOCTYPE x [<!ENTITY y "expanded">]>' + XML)


class Target:
    def __init__(
        self,
        *,
        start=10,
        end=10,
        before=100,
        after=128,
        inserted_text="\x01",
        original_text="",
        object_count=1,
        tail_text="\r",
    ):
        self.Start, self.End = start, end
        self.Application = SimpleNamespace(_oleobj_=object())
        self.Paragraphs = Mock(
            return_value=SimpleNamespace(Range=SimpleNamespace(WordOpenXML=DESTINATION_XML))
        )
        self.Document = SimpleNamespace(
            Content=SimpleNamespace(End=before), Range=Mock(side_effect=self._tail)
        )
        self.after = after
        self.original_text = original_text
        self.inserted_text = inserted_text
        self.tail_text = tail_text
        self.inserted = False
        self.InlineShapes = SimpleNamespace(Count=object_count)
        self.InsertXML = Mock(side_effect=self._insert)
        self.SetRange = Mock(side_effect=self._range)
        self.Delete = Mock(side_effect=self._delete_tail)

    @property
    def Text(self):
        return self.inserted_text if self.inserted else self.original_text

    def _insert(self, xml):
        self.inserted = True
        self.Document.Content.End = self.after
        self.Start = self.End = 90  # Word may collapse/move the live Range.

    def _range(self, start, end):
        self.Start, self.End = start, end

    def _tail(self, start, end):
        return SimpleNamespace(Start=start, End=end, Text=self.tail_text, Delete=self.Delete)

    def _delete_tail(self):
        self.inserted_text = self.inserted_text[:-1]
        self.Document.Content.End -= 1


def _source():
    return SimpleNamespace(
        Application=SimpleNamespace(_oleobj_=object()), WordOpenXML=XML, Text="\x01"
    )


def test_same_application_uses_formatted_text_without_serializing():
    app_identity = object()

    class Source:
        Application = SimpleNamespace(_oleobj_=app_identity)
        FormattedText = object()

        @property
        def WordOpenXML(self):
            raise AssertionError("Same-instance transfer must not serialize XML")

    target = Target()
    target.Application = SimpleNamespace(_oleobj_=app_identity)
    source = Source()
    copy_native_range(target, source)
    assert target.FormattedText is source.FormattedText
    target.InsertXML.assert_not_called()
    target.SetRange.assert_not_called()


def test_cross_instance_insert_rebinds_full_hidden_ole_span():
    target = Target(start=10, end=10, before=100, after=128)
    copy_native_range(target, _source())
    target.InsertXML.assert_called_once()
    target.SetRange.assert_called_once_with(10, 38)
    assert (target.Start, target.End) == (10, 38)
    inserted = target.InsertXML.call_args.args[0]
    assert 'w:pStyle w:val="TargetNumbered"' in inserted
    assert 'w:pStyle w:val="SourceDisplay"' not in inserted
    target.Document.Range.assert_not_called()
    target.Delete.assert_not_called()


def test_only_added_terminal_paragraph_mark_is_removed_from_cross_instance_insert():
    target = Target(start=10, end=10, before=100, after=129, inserted_text="\x01\r")
    copy_native_range(target, _source())
    target.Document.Range.assert_called_once_with(38, 39)
    target.Delete.assert_called_once_with()
    assert [call.args for call in target.SetRange.call_args_list] == [(10, 39), (10, 38)]
    assert target.Document.Content.End == 128
    assert target.Text == "\x01"


def test_added_mark_is_removed_after_replacing_text_without_deleting_original_mark():
    target = Target(
        start=10, end=15, before=100, after=124, inserted_text="\x01\r", original_text="TOKEN"
    )
    copy_native_range(target, _source())
    target.Document.Range.assert_called_once_with(38, 39)
    assert (target.Start, target.End) == (10, 38)
    assert target.Document.Content.End == 123


@pytest.mark.parametrize("text", ["\x01\r\r", "\x01 unexpected", "", "\r\x01"])
def test_unexpected_inserted_text_is_not_deleted_speculatively(text):
    target = Target(inserted_text=text)
    with pytest.raises(RuntimeError, match="unexpected text"):
        copy_native_range(target, _source())
    target.Delete.assert_not_called()


@pytest.mark.parametrize("count", [0, 2])
def test_unexpected_inserted_object_count_is_reported_without_trimming(count):
    target = Target(inserted_text="\x01\r", object_count=count)
    with pytest.raises(RuntimeError, match="unexpected text"):
        copy_native_range(target, _source())
    target.Delete.assert_not_called()


def test_paragraph_mark_position_must_agree_before_deleting_any_character():
    target = Target(inserted_text="\x01\r", tail_text="X")
    with pytest.raises(RuntimeError, match="identify the inserted paragraph mark"):
        copy_native_range(target, _source())
    target.Delete.assert_not_called()


@pytest.mark.parametrize("which", ["source", "target"])
def test_original_paragraph_marks_are_never_candidates_for_removal(which):
    source = _source()
    target = Target()
    if which == "source":
        source.Text = "\x01\r"
    else:
        target.original_text = "TOKEN\r"
    with pytest.raises(ValueError, match="exclude paragraph marks"):
        copy_native_range(target, source)
    target.InsertXML.assert_not_called()
    target.Delete.assert_not_called()


@pytest.mark.parametrize(
    ("end", "after", "expected_end"), [(20, 118, 38), (38, 100, 38), (45, 93, 38)]
)
def test_replacement_rebinding_accounts_for_removed_range_length(end, after, expected_end):
    target = Target(start=10, end=end, before=100, after=after)
    copy_native_range(target, _source())
    target.SetRange.assert_called_once_with(10, expected_end)


def test_invalid_source_is_rejected_before_target_mutation():
    target = Target()
    source = _source()
    source.WordOpenXML = XML.replace('ProgID="Equation.AxMath"', 'ProgID="Package"')
    with pytest.raises(ValueError):
        copy_native_range(target, source)
    target.InsertXML.assert_not_called()


def test_invalid_initial_range_is_rejected_before_target_mutation():
    target = Target(start=20, end=10)
    with pytest.raises(ValueError, match="destination Word range"):
        copy_native_range(target, _source())
    target.InsertXML.assert_not_called()


def test_unexpected_negative_delta_is_reported_without_blind_retry():
    target = Target(start=10, end=10, before=100, after=1)
    with pytest.raises(RuntimeError, match="Inspect the document"):
        copy_native_range(target, _source())
    target.InsertXML.assert_called_once()
    target.SetRange.assert_not_called()

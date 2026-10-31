"""Layout safety checks independent of a running Word instance."""

import sys
from types import SimpleNamespace
from unittest.mock import Mock, call

import pytest

from axmath_mcp import layout as layout_module
from axmath_mcp.layout import (
    DISPLAY_STYLE,
    NUMBER_BOOKMARK,
    NUMBERED_STYLE,
    available_width,
    check_block_fits,
    describe_layout,
    inline_paragraph,
    managed_target,
    number_reserve_pt,
    place_equation,
    refresh_numbers,
)
from axmath_mcp.validation import validate_layout_options


def _field(code, *, field_type=12, locked=False, success=True, result="1"):
    return SimpleNamespace(
        Type=field_type,
        Code=SimpleNamespace(Text=code),
        Locked=locked,
        Update=Mock(return_value=success),
        Result=SimpleNamespace(Text=result),
    )


def _page_target(*, table=False, columns=1, gutter=0, gutter_position=0, text_width=468, fields=()):
    setup = SimpleNamespace(
        PageWidth=text_width + 144,
        LeftMargin=72,
        RightMargin=72,
        Gutter=gutter,
        GutterPos=gutter_position,
        TextColumns=SimpleNamespace(Count=columns),
    )
    return SimpleNamespace(
        Information=Mock(return_value=table),
        Sections=Mock(return_value=SimpleNamespace(PageSetup=setup)),
        Document=SimpleNamespace(Fields=list(fields)),
    )


def test_sequence_refresh_does_not_update_other_numbering_or_external_fields():
    owned = _field(r" SEQ AMEqn \c \* ARABIC ", result="3")
    other = [
        _field("SEQ Figure", locked=True),
        _field("SEQ UnrelatedEquationSupplement"),
        _field("INCLUDETEXT external.docx", field_type=68),
        _field("REF AxMathEquation", field_type=3),
    ]

    result = refresh_numbers(SimpleNamespace(Fields=[other[0], owned, *other[1:]]))

    assert result == {"updated_count": 1, "numbers": ["3"]}
    owned.Update.assert_called_once_with()
    for field in other:
        field.Update.assert_not_called()


def test_locked_owned_sequence_is_detected_before_any_number_is_changed():
    first = _field(r"SEQ AMEqn \c", result="2")
    locked = _field(r"SEQ ameqn \c", locked=True)

    with pytest.raises(ValueError, match="locked"):
        refresh_numbers(SimpleNamespace(Fields=[first, locked]))

    first.Update.assert_not_called()
    locked.Update.assert_not_called()


def test_failed_number_update_is_reported_and_later_numbers_are_not_updated():
    broken = _field(r"SEQ AMEqn \c", success=False)
    following = _field(r"SEQ AMEqn \c")

    with pytest.raises(RuntimeError, match="could not update"):
        refresh_numbers(SimpleNamespace(Fields=[broken, following]))

    following.Update.assert_not_called()


@pytest.mark.parametrize(("gutter_position", "expected"), [(0, 444), (1, 468), (2, 444)])
def test_available_width_handles_horizontal_and_top_gutters(gutter_position, expected):
    target = _page_target(gutter=24, gutter_position=gutter_position)
    assert available_width(target) == expected


@pytest.mark.parametrize(
    ("kwargs", "message"), [({"table": True}, "tables"), ({"columns": 2}, "single-column")]
)
def test_unsupported_block_containers_are_rejected(kwargs, message):
    with pytest.raises(ValueError, match=message):
        available_width(_page_target(**kwargs))


def test_numbered_formula_reserves_measured_label_and_gap_equally(monkeypatch):
    monkeypatch.setattr(layout_module, "_number_text_width_em", lambda _: 14 / 12)
    target = _page_target()
    options = validate_layout_options("numbered")
    # Measured 14 pt label + 8 pt gap on EACH side, not 96 pt per side.
    assert check_block_fits(target, SimpleNamespace(Width=424), options, 12) == 468
    with pytest.raises(ValueError, match="424.0 pt is available"):
        check_block_fits(target, SimpleNamespace(Width=424.5), options, 12)


def test_reported_markov_formula_is_not_rejected_by_six_digit_reserve(monkeypatch):
    measure = Mock(return_value=14 / 12)
    monkeypatch.setattr(layout_module, "_number_text_width_em", measure)
    target = _page_target(
        text_width=447.9, fields=[_field(r"SEQ AMEqn \c", result=str(i)) for i in range(1, 4)]
    )
    options = validate_layout_options("numbered")
    assert check_block_fits(target, SimpleNamespace(Width=343.8), options, 12) == pytest.approx(
        447.9
    )
    measure.assert_called_once_with("(8)")
    with pytest.raises(ValueError, match="403.9 pt is available"):
        check_block_fits(target, SimpleNamespace(Width=404), options, 12)


@pytest.mark.parametrize("count,expected", [(0, "(8)"), (8, "(8)"), (9, "(88)"), (99, "(888)")])
def test_automatic_width_covers_next_sequence_digit_count(monkeypatch, count, expected):
    measure = Mock(return_value=1.5)
    monkeypatch.setattr(layout_module, "_number_text_width_em", measure)
    target = _page_target(fields=[_field(r"SEQ AMEqn \c", result=str(i + 1)) for i in range(count)])
    assert number_reserve_pt(target, validate_layout_options("numbered"), 12) == 26
    measure.assert_called_once_with(expected)


def test_existing_sequence_restart_value_is_not_undercounted(monkeypatch):
    measure = Mock(return_value=2.5)
    monkeypatch.setattr(layout_module, "_number_text_width_em", measure)
    target = _page_target(fields=[_field(r"SEQ AMEqn \c \\r 1000", result="1000")])
    number_reserve_pt(target, validate_layout_options("numbered"), 12)
    measure.assert_called_once_with("(8888)")


def test_manual_unicode_label_is_measured_as_its_exact_text(monkeypatch):
    measure = Mock(return_value=5.0)
    monkeypatch.setattr(layout_module, "_number_text_width_em", measure)
    options = validate_layout_options("numbered", number_label="（附录甲）")
    assert number_reserve_pt(_page_target(), options, 12) == 68
    measure.assert_called_once_with("（附录甲）")


@pytest.mark.skipif(sys.platform != "win32", reason="Windows GDI font metrics")
def test_gdi_measures_real_number_width_without_the_old_full_em_estimate():
    measured = layout_module._number_text_width_em("(1)") * 12
    assert 12 < measured < 17
    assert layout_module._number_text_width_em("(888)") > layout_module._number_text_width_em("(8)")
    assert layout_module._number_text_width_em("（附录甲）") > 0
    assert layout_module._number_text_width_em("(😀)") > 0


def test_unnumbered_display_uses_full_text_width_without_scaling_object():
    shape = SimpleNamespace(Width=460)
    assert check_block_fits(_page_target(), shape, validate_layout_options("display"), 12) == 468
    assert shape.Width == 460


def test_long_explicit_label_cannot_silently_overlap_the_formula():
    options = validate_layout_options("numbered", number_label="(" + "W" * 60 + ")")
    with pytest.raises(ValueError, match="available"):
        check_block_fits(_page_target(), SimpleNamespace(Width=100), options, 12)


@pytest.mark.parametrize("rule", [0, 3, 4, 5])
def test_inline_layout_preserves_native_half_point_baseline_and_nonfixed_spacing(rule):
    class EquationRange:
        @property
        def Font(self):
            raise AssertionError("Do not rewrite AxMath's native half-point baseline offset")

    fmt = SimpleNamespace(LineSpacingRule=rule, LineSpacing=18)
    rng = EquationRange()
    rng.Paragraphs = Mock(return_value=SimpleNamespace(Format=fmt))

    inline_paragraph(SimpleNamespace(Range=rng))

    assert fmt.LineSpacingRule == (3 if rule == 4 else rule)
    assert fmt.LineSpacing == 18
    assert fmt.BaseLineAlignment == 2


def _managed_document(text, style, *, marker=None, object_count=1, automatic=False):
    object_start = text.index("\x01")
    bookmarks = []
    if marker:
        a, b = marker
        fields = [_field(r"SEQ AMEqn \c")] if automatic else []
        bookmarks.append(
            SimpleNamespace(
                Name=NUMBER_BOOKMARK + "test",
                Range=SimpleNamespace(Start=a, End=b, Text=text[a:b], Fields=fields),
            )
        )
    paragraph = SimpleNamespace(
        Start=0,
        End=len(text),
        Style=SimpleNamespace(NameLocal=style),
        InlineShapes=SimpleNamespace(Count=object_count),
        Bookmarks=bookmarks,
    )
    rng = SimpleNamespace(
        Start=object_start,
        End=object_start + 1,
        Paragraphs=Mock(return_value=SimpleNamespace(Range=paragraph)),
        Duplicate=object(),
    )
    shape = SimpleNamespace(Range=rng)
    doc = SimpleNamespace(Range=lambda a, b: SimpleNamespace(Start=a, End=b, Text=text[a:b]))
    return doc, shape


def test_managed_display_target_excludes_the_paragraph_mark():
    doc, shape = _managed_document("\t\x01 \r", DISPLAY_STYLE)
    target = managed_target(doc, shape)
    assert target.Text == "\t\x01 "
    assert target.End == 3


@pytest.mark.parametrize("text", ["explanation \x01\r", "\x01 user annotation\r"])
def test_display_reformat_refuses_unrelated_text(text):
    doc, shape = _managed_document(text, DISPLAY_STYLE)
    with pytest.raises(ValueError, match="additional text"):
        managed_target(doc, shape)


def test_multiple_objects_are_not_deleted_as_one_managed_formula():
    doc, shape = _managed_document("\x01\x01\r", DISPLAY_STYLE, object_count=2)
    with pytest.raises(ValueError, match="other objects"):
        managed_target(doc, shape)


def test_numbered_target_recognizes_only_its_marked_number():
    doc, shape = _managed_document("\t\x01\t(1)\r", NUMBERED_STYLE, marker=(3, 6), automatic=True)
    assert describe_layout(shape) == {
        "layout": "numbered",
        "number_label": None,
        "number_text": "(1)",
    }
    assert managed_target(doc, shape).Text == "\t\x01\t(1)"


def test_numbered_reformat_refuses_text_after_the_number_bookmark():
    doc, shape = _managed_document("\t\x01\t(1) user annotation\r", NUMBERED_STYLE, marker=(3, 6))
    with pytest.raises(ValueError, match="additional text"):
        managed_target(doc, shape)


def test_numbered_reformat_requires_its_number_marker():
    doc, shape = _managed_document("\t\x01\t(1)\r", NUMBERED_STYLE)
    with pytest.raises(ValueError, match="additional text"):
        managed_target(doc, shape)


def test_plain_inline_formula_never_expands_to_surrounding_prose():
    doc, shape = _managed_document("Chinese 中文 \x01 English prose\r", "Normal")
    assert managed_target(doc, shape) is shape.Range.Duplicate


@pytest.mark.parametrize(
    ("style", "start", "end"),
    [(DISPLAY_STYLE, 28, 28), (NUMBERED_STYLE, 28, 28), (DISPLAY_STYLE, 0, 27)],
)
def test_inline_insertion_cannot_restyle_an_existing_managed_block(style, start, end):
    paragraph_range = SimpleNamespace(Start=0, End=29, Style=SimpleNamespace(NameLocal=style))
    target = SimpleNamespace(
        Start=start,
        End=end,
        Text="untouched",
        Paragraphs=Mock(return_value=SimpleNamespace(Range=paragraph_range)),
    )
    replace = Mock()

    with pytest.raises(ValueError, match="word_format_axmath"):
        place_equation(
            SimpleNamespace(), target, object(), validate_layout_options("inline"), 12, replace
        )

    assert target.Text == "untouched"
    assert paragraph_range.Style.NameLocal == style
    replace.assert_not_called()


@pytest.mark.parametrize("mode", ["display", "numbered", "inline"])
def test_style_is_applied_before_copy_so_it_cannot_clear_native_baseline(mode, monkeypatch):
    """Real Word linked styles cleared w:position when applied after object copy."""
    state = {"position": None, "events": []}
    original_style = DISPLAY_STYLE if mode == "inline" else "Normal"
    original_end = 29 if mode == "inline" else 1

    class ParagraphRange:
        Start = 0
        End = original_end
        _style = SimpleNamespace(NameLocal=original_style)

        @property
        def Style(self):
            return self._style

        @Style.setter
        def Style(self, value):
            state["events"].append("style")
            self._style = value
            # Emulate the observed destructive formatting side effect of Word.
            if state["position"] is not None:
                state["position"] = 0

    paragraph = SimpleNamespace(
        Range=ParagraphRange(),
        Format=SimpleNamespace(LineSpacingRule=4, LineSpacing=18, TabStops=Mock()),
    )

    class Range:
        def __init__(self, start, end):
            self.Start, self.End = start, end
            self.Text = ""
            self.Font = SimpleNamespace()
            self.Paragraphs = Mock(return_value=paragraph)
            page = _page_target()
            self.Information, self.Sections = page.Information, page.Sections
            self.Document = page.Document

        def InsertAfter(self, text):
            self.End += len(text)
            paragraph.Range.End += len(text)

    normal = SimpleNamespace(NameLocal="Normal")
    doc = SimpleNamespace(
        Range=Range,
        Styles=Mock(return_value=normal),
        Fields=[],
        Bookmarks=SimpleNamespace(Add=Mock()),
    )
    monkeypatch.setattr(layout_module, "_style", lambda _, name: SimpleNamespace(NameLocal=name))

    def copy_object(_, target, generated):
        state["events"].append("copy")
        state["position"] = generated.baseline
        target.End = target.Start + 28
        paragraph.Range.End = target.End + 1
        return SimpleNamespace(Range=target)

    target = Range(0, original_end - 1)
    options = validate_layout_options(mode, number_label="(A)" if mode == "numbered" else None)
    replace = Mock(side_effect=copy_object)
    generated = SimpleNamespace(Width=80, baseline=-14.5)

    place_equation(doc, target, generated, options, 12, replace)

    assert state["events"] == ["style", "copy"]
    assert state["position"] == -14.5
    replace.assert_called_once()


def test_locked_sequence_blocks_numbered_insertion_before_target_changes():
    doc = SimpleNamespace(Fields=[_field(r"SEQ AMEqn \c", locked=True)])
    target = SimpleNamespace(Text="existing target text")
    replace = Mock()

    with pytest.raises(ValueError, match="locked"):
        place_equation(doc, target, object(), validate_layout_options("numbered"), 12, replace)

    assert target.Text == "existing target text"
    replace.assert_not_called()


@pytest.mark.parametrize("mode", ["display", "numbered"])
def test_replaced_paragraph_receives_final_alignment_and_number_tabs(mode, monkeypatch):
    """InsertXML can invalidate the paragraph/Format proxies captured before copy."""
    state = {"inserted": False}
    old_format = SimpleNamespace(Alignment=99, TabStops=Mock())
    new_format = SimpleNamespace(Alignment=99, TabStops=Mock())
    old_paragraph = SimpleNamespace(
        Range=SimpleNamespace(Start=0, End=1, Style=SimpleNamespace(NameLocal="Normal")),
        Format=old_format,
    )
    new_paragraph = SimpleNamespace(
        Range=SimpleNamespace(Start=0, End=1, Style=SimpleNamespace(NameLocal="Normal")),
        Format=new_format,
    )

    class Range:
        def __init__(self, start, end):
            self.Start, self.End = start, end
            self.Text = ""
            self.Font = SimpleNamespace()
            # This proxy is bound to the paragraph existing when Range was made.
            self.paragraph = new_paragraph if state["inserted"] else old_paragraph
            self.Paragraphs = Mock(return_value=self.paragraph)

        def InsertAfter(self, text):
            self.End += len(text)
            self.paragraph.Range.End += len(text)

    def replace(_doc, location, _generated):
        state["inserted"] = True
        new_paragraph.Range.End = location.Start + 28 + 1
        return SimpleNamespace(Range=Range(location.Start, location.Start + 28))

    doc = SimpleNamespace(Range=Range, Fields=[], Bookmarks=SimpleNamespace(Add=Mock()))
    monkeypatch.setattr(layout_module, "check_block_fits", lambda *_: 447.9)
    monkeypatch.setattr(layout_module, "_style", lambda _, name: SimpleNamespace(NameLocal=name))
    options = validate_layout_options(mode, number_label="(1)" if mode == "numbered" else None)

    inserted = place_equation(doc, Range(0, 0), object(), options, 12, replace)

    assert inserted.Range.Paragraphs(1) is new_paragraph
    assert old_format.Alignment == 99, "Post-insertion alignment must not target the old proxy."
    assert new_format.Alignment == (0 if mode == "numbered" else 1)
    if mode == "numbered":
        new_format.TabStops.ClearAll.assert_called_once_with()
        assert new_format.TabStops.Add.call_args_list == [
            call(Position=223.95, Alignment=1, Leader=0),
            call(Position=447.9, Alignment=2, Leader=0),
        ]
        old_format.TabStops.Add.assert_not_called()
        assert doc.Bookmarks.Add.call_args.kwargs["Range"].End == new_paragraph.Range.End - 1
    else:
        new_format.TabStops.Add.assert_not_called()

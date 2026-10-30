"""PowerPoint preservation and cleanup checks without launching Office."""

from contextlib import nullcontext
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

pytest.importorskip("pythoncom")

from axmath_mcp import powerpoint as ppt


class Collection:
    def __init__(self, values=()):
        self.values = list(values)

    @property
    def Count(self):
        return len(self.values)

    def Item(self, index):
        return self.values[index - 1]


class Tags:
    def __init__(self, values=None):
        self.values = values or {}

    def Item(self, name):
        return self.values.get(name, "")

    def Add(self, name, value):
        self.values[name] = str(value).upper()


def equation(shape_id=1, **overrides):
    values = {
        "Id": shape_id,
        "Type": 7,
        "OLEFormat": SimpleNamespace(ProgID="Equation.AxMath"),
        "Left": 72.0,
        "Top": 144.0,
        "Width": 120.0,
        "Height": 48.0,
        "Name": "Original formula",
        "ZOrderPosition": 1,
        "Delete": Mock(),
        "Tags": Tags({ppt.FONT_TAG: "24", ppt.WIDTH_TAG: "120", ppt.HEIGHT_TAG: "48"}),
        "ActionSettings": SimpleNamespace(Item=lambda _: SimpleNamespace(Action=0)),
        "Rotation": 0,
        "HorizontalFlip": 0,
        "VerticalFlip": 0,
        "Connector": 0,
    }
    values.update(overrides)
    return SimpleNamespace(**values)


@pytest.fixture
def deck(tmp_path):
    original = equation()
    slide = SimpleNamespace(
        SlideID=256,
        SlideIndex=1,
        Shapes=Collection([original]),
        TimeLine=SimpleNamespace(MainSequence=Collection(), InteractiveSequences=Collection()),
    )
    slides = Collection([slide])

    def find(sid):
        if sid != 256:
            raise ppt.pywintypes.com_error(-2147352567, "Slide not found", None, None)
        return slide

    slides.FindBySlideID = Mock(side_effect=find)
    app = SimpleNamespace(DisplayAlerts=2)
    pres = SimpleNamespace(
        Name="original.pptx",
        Path=str(tmp_path),
        FullName=str(tmp_path / "original.pptx"),
        ReadOnly=False,
        HasVBProject=False,
        Slides=slides,
        PageSetup=SimpleNamespace(SlideWidth=960, SlideHeight=540),
        Application=app,
        Save=Mock(),
        SaveAs=Mock(),
    )
    word = Mock()
    word._conversion.side_effect = lambda _: nullcontext()
    word._generated.side_effect = lambda *_: nullcontext((object(), "$x$"))
    backend = ppt.PowerPointBackend(word)
    backend._presentations["pres_test"] = pres
    backend._equations["eq_test"] = ("pres_test", 256, 1)
    backend._info = Mock(return_value={"presentation_id": "pres_test"})
    backend.ppt_list_presentations = Mock(return_value={"presentations": []})
    return backend, pres, slide, original


@pytest.mark.parametrize("operation", ["insert", "update", "move", "save", "slide"])
def test_read_only_presentation_rejects_mutations_before_converter_or_edit(deck, operation):
    backend, pres, slide, original = deck
    pres.ReadOnly = True
    backend._generate = Mock()
    slide.Shapes.AddOLEObject = Mock()
    slide.Slides = Mock()
    calls = {
        "insert": lambda: backend.ppt_insert_equation("pres_test", 256, "x"),
        "update": lambda: backend.ppt_update_equation("pres_test", 256, "eq_test", "y"),
        "move": lambda: backend.ppt_position_equation("pres_test", 256, "eq_test", 10, 10),
        "save": lambda: backend.ppt_save_presentation("pres_test"),
        "slide": lambda: backend.ppt_add_slide("pres_test"),
    }
    with pytest.raises(ValueError, match="read-only"):
        calls[operation]()
    backend._generate.assert_not_called()
    backend.word._active_app.assert_not_called()
    pres.Save.assert_not_called()
    original.Delete.assert_not_called()


def test_out_of_bounds_move_does_not_touch_equation_or_its_number(deck):
    backend, _, slide, original = deck
    label = SimpleNamespace(
        Id=2,
        Tags=Tags({ppt.OWNER_TAG: "1"}),
        HasTextFrame=True,
        Left=204.0,
        Top=156.0,
        Width=40.0,
        Height=24.0,
    )
    original.Tags.Add(ppt.LABEL_TAG, "2")
    slide.Shapes.values.append(label)
    with pytest.raises(ValueError, match="does not fit"):
        backend.ppt_position_equation("pres_test", 256, "eq_test", 900, 144)
    assert (original.Left, original.Top, label.Left, label.Top) == (72, 144, 204, 156)


def test_taller_number_must_fit_above_the_equation_before_move(deck):
    backend, pres, _, original = deck
    label = SimpleNamespace(Width=30, Height=70)
    with pytest.raises(ValueError, match="nonnegative"):
        backend._check_bounds(pres, original, label, 10, 0)


@pytest.mark.parametrize("reference", [None, ("other_pres", 256, 1), ("pres_test", 257, 1)])
def test_equation_id_cannot_cross_presentation_or_slide(deck, reference):
    backend, _, _, _ = deck
    backend._equations = {} if reference is None else {"eq_test": reference}
    with pytest.raises(ValueError, match="Unknown equation_id|wrong presentation/slide"):
        backend._eq("pres_test", 256, "eq_test")


@pytest.mark.parametrize("replacement", [None, equation(Type=13)])
def test_deleted_or_non_ole_replacement_cannot_reuse_equation_id(deck, replacement):
    backend, _, slide, _ = deck
    slide.Shapes.values = [] if replacement is None else [replacement]
    with pytest.raises(ValueError, match="removed/replaced"):
        backend._eq("pres_test", 256, "eq_test")


def test_wrong_slide_index_is_not_treated_as_stable_slide_id(deck):
    backend, _, _, _ = deck
    with pytest.raises(ValueError, match="SlideID not found"):
        backend._slide("pres_test", 1)


def test_equation_identity_survives_slide_reordering(deck):
    backend, _, slide, original = deck
    slide.SlideIndex = 9
    assert backend._eq("pres_test", 256, "eq_test")[2] is original


@pytest.mark.parametrize("name", ["legacy.ppt", "macro.pptm", "Presentation1"])
def test_save_as_pptx_never_drops_vba(deck, tmp_path, name):
    backend, pres, _, _ = deck
    pres.Name = name
    pres.HasVBProject = True
    with pytest.raises(ValueError, match="preserve VBA"):
        backend.ppt_save_presentation("pres_test", str(tmp_path / "new.pptx"))
    pres.SaveAs.assert_not_called()
    assert pres.Application.DisplayAlerts == 2


def test_explicit_in_place_save_preserves_macro_format(deck):
    backend, pres, _, _ = deck
    pres.Name = "macro.pptm"
    pres.HasVBProject = True
    backend.ppt_save_presentation("pres_test")
    pres.Save.assert_called_once_with()
    pres.SaveAs.assert_not_called()


def test_existing_destination_is_not_overwritten_by_default(deck, tmp_path):
    backend, pres, _, _ = deck
    target = tmp_path / "existing.pptx"
    target.write_bytes(b"existing presentation")
    with pytest.raises(ValueError, match="Destination exists"):
        backend.ppt_save_presentation("pres_test", str(target))
    pres.SaveAs.assert_not_called()
    assert target.read_bytes() == b"existing presentation"


def test_conversion_failure_leaves_original_and_stable_id_intact(deck):
    backend, _, _, original = deck
    backend.word._generated.side_effect = RuntimeError("vendor conversion failed")
    backend._import_native = Mock()
    with pytest.raises(RuntimeError, match="vendor conversion failed"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    original.Delete.assert_not_called()
    backend._import_native.assert_not_called()
    assert backend._equations["eq_test"] == ("pres_test", 256, 1)


def test_new_equation_is_removed_if_number_creation_fails(deck):
    backend, pres, slide, original = deck
    candidate = equation(2)
    backend._import_native = Mock(return_value=candidate)
    backend._new_label = Mock(side_effect=RuntimeError("number creation failed"))
    with pytest.raises(RuntimeError, match="number creation failed"):
        backend._generate(pres, slide, "x", 72, 144, 24, "(1)")
    candidate.Delete.assert_called_once_with()
    original.Delete.assert_not_called()


def test_number_and_equation_candidates_are_removed_if_they_do_not_fit(deck):
    backend, pres, slide, original = deck
    candidate = equation(2, Width=1000)
    label = SimpleNamespace(Width=40.0, Height=24.0, Delete=Mock())
    backend._import_native = Mock(return_value=candidate)
    backend._new_label = Mock(return_value=label)
    with pytest.raises(ValueError, match="does not fit"):
        backend._generate(pres, slide, "x", 72, 144, 24, "(1)")
    candidate.Delete.assert_called_once_with()
    label.Delete.assert_called_once_with()
    original.Delete.assert_not_called()


def test_half_created_number_textbox_is_cleaned_up_on_format_failure(deck):
    backend, _, slide, _ = deck

    class BrokenLabel:
        Delete = Mock()

        @property
        def TextFrame(self):
            raise RuntimeError("font/frame unavailable")

    label = BrokenLabel()
    slide.Shapes.AddTextbox = Mock(return_value=label)
    with pytest.raises(RuntimeError, match="font/frame unavailable"):
        backend._new_label(slide, "(1)", 24)
    label.Delete.assert_called_once_with()


def test_manually_resized_equation_requires_explicit_font_before_conversion(deck):
    backend, _, _, original = deck
    original.Width = 150
    backend._generate = Mock()
    with pytest.raises(ValueError, match="resized manually"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()
    original.Delete.assert_not_called()


def test_external_equation_requires_explicit_font_size(deck):
    backend, _, _, original = deck
    original.Tags = Tags()
    backend._generate = Mock()
    with pytest.raises(ValueError, match="no recorded font size"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()
    original.Delete.assert_not_called()


def test_uncertain_old_delete_retains_validated_recovery_copy(deck):
    backend, _, _, original = deck
    original.Delete.side_effect = ppt.pywintypes.com_error(
        -2147352567, "Delete reply lost", None, None
    )
    candidate = equation(3)
    label = SimpleNamespace(Delete=Mock())
    backend._generate = Mock(return_value=(candidate, label, "$y$"))
    with pytest.raises(RuntimeError, match="validated replacement"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    candidate.Delete.assert_not_called()
    label.Delete.assert_not_called()
    assert any(ref == ("pres_test", 256, 3) for ref in backend._equations.values())


def test_action_bound_equation_is_not_replaced_silently(deck):
    backend, _, _, original = deck
    original.ActionSettings.Item = lambda _: SimpleNamespace(Action=7)
    backend._generate = Mock()
    with pytest.raises(ValueError, match="actions"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()


def test_animation_bound_equation_is_not_replaced_silently(deck):
    backend, _, slide, original = deck
    slide.TimeLine.MainSequence = Collection([SimpleNamespace(Shape=original)])
    backend._generate = Mock()
    with pytest.raises(ValueError, match="animations"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()


@pytest.mark.parametrize(
    "attribute,value", [("Rotation", 45), ("HorizontalFlip", -1), ("VerticalFlip", -1)]
)
def test_oriented_equation_replacement_is_refused_before_conversion(deck, attribute, value):
    backend, _, _, original = deck
    setattr(original, attribute, value)
    backend._generate = Mock()
    with pytest.raises(ValueError, match="rotated or flipped"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()
    original.Delete.assert_not_called()


def test_associated_number_orientation_is_protected_too(deck):
    backend, _, slide, original = deck
    label = equation(
        2,
        Type=17,
        Rotation=90,
        HasTextFrame=True,
        TextFrame=SimpleNamespace(TextRange=SimpleNamespace(Text="(1)")),
        Tags=Tags({ppt.OWNER_TAG: "1"}),
    )
    original.Tags.Add(ppt.LABEL_TAG, "2")
    slide.Shapes.values.append(label)
    backend._generate = Mock()
    with pytest.raises(ValueError, match="rotated or flipped"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()
    original.Delete.assert_not_called()
    label.Delete.assert_not_called()


@pytest.mark.parametrize("endpoint", ["Begin", "End"])
@pytest.mark.parametrize("owner", ["equation", "number"])
def test_connector_attachments_are_preserved_by_refusing_replacement(deck, endpoint, owner):
    backend, _, slide, original = deck
    label = equation(
        2,
        Type=17,
        HasTextFrame=True,
        TextFrame=SimpleNamespace(TextRange=SimpleNamespace(Text="(1)")),
        Tags=Tags({ppt.OWNER_TAG: "1"}),
    )
    original.Tags.Add(ppt.LABEL_TAG, "2")
    connected_shape = original if owner == "equation" else label
    settings = {"BeginConnected": False, "EndConnected": False}
    settings[endpoint + "Connected"] = True
    settings[endpoint + "ConnectedShape"] = connected_shape
    connector = SimpleNamespace(Id=3, Connector=-1, ConnectorFormat=SimpleNamespace(**settings))
    slide.Shapes.values.extend([label, connector])
    backend._generate = Mock()
    with pytest.raises(ValueError, match="attached connectors"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    backend._generate.assert_not_called()
    original.Delete.assert_not_called()
    label.Delete.assert_not_called()


def test_unrelated_or_disconnected_connectors_do_not_block_replacement(deck):
    backend, _, slide, original = deck
    slide.Shapes.values.extend(
        [
            SimpleNamespace(
                Id=3,
                Connector=-1,
                ConnectorFormat=SimpleNamespace(BeginConnected=False, EndConnected=False),
            ),
            SimpleNamespace(
                Id=4,
                Connector=-1,
                ConnectorFormat=SimpleNamespace(
                    BeginConnected=True,
                    BeginConnectedShape=SimpleNamespace(Id=99),
                    EndConnected=False,
                ),
            ),
        ]
    )
    backend._assert_replaceable(slide, [original])


def test_candidate_placement_failure_preserves_old_equation_and_removes_candidates(deck):
    backend, _, _, original = deck
    candidate = equation(
        3, ZOrderPosition=2, ZOrder=Mock(side_effect=RuntimeError("placement failed"))
    )
    label = SimpleNamespace(Delete=Mock())
    backend._generate = Mock(return_value=(candidate, label, "$y$"))
    with pytest.raises(RuntimeError, match="placement failed"):
        backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    original.Delete.assert_not_called()
    candidate.Delete.assert_called_once_with()
    label.Delete.assert_called_once_with()
    assert backend._equations["eq_test"] == ("pres_test", 256, 1)


def test_old_number_cleanup_failure_keeps_new_equation_id_and_reports_partial_cleanup(deck):
    backend, _, slide, original = deck
    old_label = equation(
        2,
        Type=17,
        HasTextFrame=True,
        TextFrame=SimpleNamespace(TextRange=SimpleNamespace(Text="(1)")),
        Tags=Tags({ppt.OWNER_TAG: "1"}),
        Delete=Mock(
            side_effect=ppt.pywintypes.com_error(-2147352567, "Cannot remove label", None, None)
        ),
    )
    original.Tags.Add(ppt.LABEL_TAG, "2")
    slide.Shapes.values.append(old_label)
    original.Delete.side_effect = lambda: slide.Shapes.values.remove(original)
    candidate = equation(3)
    new_label = equation(
        4,
        Type=17,
        HasTextFrame=True,
        TextFrame=SimpleNamespace(TextRange=SimpleNamespace(Text="(1)")),
        Tags=Tags({ppt.OWNER_TAG: "3"}),
    )
    candidate.Tags.Add(ppt.LABEL_TAG, "4")

    def generate(*_):
        slide.Shapes.values.extend([candidate, new_label])
        return candidate, new_label, "$y$"

    backend._generate = Mock(side_effect=generate)
    result = backend.ppt_update_equation("pres_test", 256, "eq_test", "y")
    assert result["equation_id"] == "eq_test"
    assert result["shape_id"] == 3
    assert result["number_label"] == "(1)"
    assert "do not repeat" in result["warnings"][0]
    assert backend._eq("pres_test", 256, "eq_test")[2] is candidate
    candidate.Delete.assert_not_called()
    new_label.Delete.assert_not_called()
    old_label.Delete.assert_called_once_with()

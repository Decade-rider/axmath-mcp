"""Exercise mutation boundaries without launching or changing Microsoft Word."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

pytest.importorskip("pythoncom")
pytest.importorskip("win32api")

from axmath_mcp import backend as backend_module
from axmath_mcp.backend import WordBackend


def _plain_paragraphs():
    return Mock(
        return_value=SimpleNamespace(
            Range=SimpleNamespace(Style=SimpleNamespace(NameLocal="Normal"))
        )
    )


@pytest.fixture
def word(tmp_path):
    doc = SimpleNamespace(
        Name="original.docx",
        Path=str(tmp_path),
        FullName=str(tmp_path / "original.docx"),
        ReadOnly=False,
        ProtectionType=-1,
        TrackRevisions=False,
        HasVBProject=False,
        Save=Mock(),
        SaveAs2=Mock(),
        _oleobj_=object(),
    )
    app = SimpleNamespace(DisplayAlerts=1, ActiveDocument=doc)
    doc.Application = app
    service = WordBackend.__new__(WordBackend)
    service._doc = Mock(return_value=doc)
    service._info = Mock(return_value={"document_id": "doc_test"})
    service.list_documents = Mock(return_value={"documents": []})
    return service, doc, app


@pytest.mark.parametrize("name", ["legacy.doc", "macro.docm", "Document1"])
def test_save_as_docx_never_discards_an_existing_vba_project(word, tmp_path, name):
    service, doc, app = word
    doc.Name = name
    doc.HasVBProject = True
    target = tmp_path / "new.docx"

    with pytest.raises(ValueError, match="VBA project"):
        service.save_document("doc_test", str(target))

    doc.SaveAs2.assert_not_called()
    doc.Save.assert_not_called()
    assert app.DisplayAlerts == 1
    assert not target.exists()


def test_existing_destination_requires_explicit_overwrite(word, tmp_path):
    service, doc, _ = word
    target = tmp_path / "existing.docx"
    target.write_bytes(b"existing content must survive")

    with pytest.raises(ValueError, match="Destination exists"):
        service.save_document("doc_test", str(target))

    doc.SaveAs2.assert_not_called()
    assert target.read_bytes() == b"existing content must survive"


def test_in_place_save_preserves_legacy_macro_format(word):
    service, doc, _ = word
    doc.Name = "legacy.doc"
    doc.HasVBProject = True

    service.save_document("doc_test")

    doc.Save.assert_called_once_with()
    doc.SaveAs2.assert_not_called()


@pytest.mark.parametrize(
    ("attribute", "value", "error"),
    [
        ("ReadOnly", True, "read-only"),
        ("ProtectionType", 2, "protected"),
        ("TrackRevisions", True, "Track Changes"),
    ],
)
def test_document_guards_reject_before_starting_conversion(word, attribute, value, error):
    service, doc, _ = word
    setattr(doc, attribute, value)
    service._conversion = Mock()

    with pytest.raises(ValueError, match=error):
        service.insert_equation("doc_test", "x^2")

    service._conversion.assert_not_called()


def test_selection_mismatch_rejects_before_starting_conversion(word):
    service, _, app = word
    selection = SimpleNamespace(Text="changed since inspection", StoryType=1)
    app.Selection = SimpleNamespace(Range=SimpleNamespace(Duplicate=selection))
    service._conversion = Mock()

    with pytest.raises(ValueError, match="expected_selection"):
        service.insert_equation(
            "doc_test", "x^2", position="selection", expected_selection="original text"
        )

    service._conversion.assert_not_called()


def test_failed_macro_restores_both_smart_paste_options(word, monkeypatch):
    service, _, app = word
    service._mutex = object()
    service._ready = Mock()
    app.Documents = SimpleNamespace(Count=0)
    app.ScreenUpdating = True
    app.Options = SimpleNamespace(
        SmartCutPaste=True, PasteSmartCutPaste=True, ReplaceSelection=True
    )
    release = Mock()
    monkeypatch.setattr(
        backend_module.win32event,
        "WaitForSingleObject",
        lambda *_: backend_module.win32event.WAIT_OBJECT_0,
    )
    monkeypatch.setattr(backend_module.win32event, "ReleaseMutex", release)

    with pytest.raises(RuntimeError, match="failed macro"), service._conversion(app):
        app.ScreenUpdating = False
        app.Options.SmartCutPaste = False
        app.Options.PasteSmartCutPaste = False
        app.Options.ReplaceSelection = False
        raise RuntimeError("failed macro")

    assert app.ScreenUpdating is True
    assert app.Options.SmartCutPaste is True
    assert app.Options.PasteSmartCutPaste is True
    assert app.Options.ReplaceSelection is True
    release.assert_called_once_with(service._mutex)


def test_ole_hidden_field_span_is_accepted_without_repeating_insertion(word):
    """A real Word OLE can display one character while occupying 28 positions."""
    service, doc, app = word
    inserted_shape = SimpleNamespace(Type=1, OLEFormat=SimpleNamespace(ProgID="Equation.AxMath"))
    generated_text = object()

    class Destination:
        Start = 0
        End = 0
        Text = ""
        copies = 0
        InlineShapes = Mock(Count=0)

        @property
        def FormattedText(self):
            return generated_text

        @FormattedText.setter
        def FormattedText(self, value):
            assert value is generated_text
            self.copies += 1
            self.End = self.Start + 28
            self.Text = "\x01"
            self.InlineShapes = Mock(Count=1, return_value=inserted_shape)

    target = Destination()
    generated = SimpleNamespace(Range=SimpleNamespace(FormattedText=generated_text))

    def get_range(start, end):
        # A one-position range excludes the OLE's hidden EMBED field characters.
        contains_object = start <= target.Start and end >= target.End
        return SimpleNamespace(
            InlineShapes=Mock(Count=1 if contains_object else 0, return_value=inserted_shape)
        )

    doc.Range = Mock(side_effect=get_range)
    doc.Activate = Mock()
    app.UndoRecord = SimpleNamespace(StartCustomRecord=Mock(), EndCustomRecord=Mock())

    result = service._replace(doc, target, generated)

    assert result is inserted_shape
    assert target.End - target.Start == 28
    assert target.Text == "\x01"
    assert target.copies == 1
    app.UndoRecord.StartCustomRecord.assert_called_once()
    app.UndoRecord.EndCustomRecord.assert_called_once()


@pytest.mark.parametrize("macro", [backend_module.TO_AXMATH, backend_module.TO_LATEX])
@pytest.mark.parametrize("completion_flag", ["0", None, "1"])
def test_vendor_success_requires_completion_acknowledgement(
    word, monkeypatch, macro, completion_flag
):
    """The vendor returns True even after its wait times out; never trust that alone."""
    service, _, app = word
    app.Run = Mock(return_value=True)
    monkeypatch.setattr(backend_module, "_reg", Mock(return_value=completion_flag))

    if completion_flag == "1":
        service._run_macro(app, macro)
    else:
        with pytest.raises(RuntimeError, match="did not acknowledge completion"):
            service._run_macro(app, macro)

    app.Run.assert_called_once_with(macro, False, False)


@pytest.mark.parametrize(
    ("previous_document", "previous_start", "reuse_id"),
    [("doc_test", 0, True), ("doc_test", 40, False), ("doc_other", 0, False)],
)
def test_equation_ids_follow_live_ranges_not_com_wrapper_identity(
    word, previous_document, previous_start, reuse_id
):
    service, _, _ = word

    def make_shape(document, start):
        rng = SimpleNamespace(
            document=document, Start=start, End=start + 27, Paragraphs=_plain_paragraphs()
        )
        rng.IsEqual = lambda other: (
            rng.document == other.document and rng.Start == other.Start and rng.End == other.End
        )
        return SimpleNamespace(
            _oleobj_=object(),
            Type=1,
            OLEFormat=SimpleNamespace(ProgID="Equation.AxMath"),
            Range=rng,
            Width=35.0,
            Height=26.0,
        )

    previous = make_shape(previous_document, previous_start)
    current = make_shape("doc_test", 0)
    service._equations = {"eq_existing": (previous_document, previous)}
    assert previous._oleobj_ != current._oleobj_

    info = service._equation_info("doc_test", current)

    assert (info["equation_id"] == "eq_existing") is reuse_id
    assert service._equations[info["equation_id"]][1] is current


def test_deleted_equation_proxy_is_rejected_even_after_location_is_reused(word):
    """Live Word keeps deleted InlineShape proxies invalid after reinsertion."""
    service, _, _ = word

    class DeletedShape:
        @property
        def Type(self):
            raise backend_module.pywintypes.com_error(
                -2147352567, "Object has been deleted", None, None
            )

        @property
        def Range(self):
            raise AssertionError("Do not inspect the range of a deleted shape")

    service._equations = {"eq_deleted": ("doc_test", DeletedShape())}

    with pytest.raises(ValueError, match="stale equation_id"):
        service._eq("doc_test", "eq_deleted")

    replacement = SimpleNamespace(
        _oleobj_=object(),
        Type=1,
        OLEFormat=SimpleNamespace(ProgID="Equation.AxMath"),
        Range=SimpleNamespace(Start=0, End=27, Paragraphs=_plain_paragraphs()),
        Width=35.0,
        Height=26.0,
    )
    info = service._equation_info("doc_test", replacement)
    assert info["equation_id"] != "eq_deleted"


@pytest.mark.parametrize("error_code", [32, 33])
def test_file_sharing_locks_fail_before_word_is_created(word, monkeypatch, tmp_path, error_code):
    service, _, _ = word
    path = tmp_path / "locked.docx"
    path.write_bytes(b"test document")
    service._active_app = Mock()
    open_file = Mock(
        side_effect=backend_module.pywintypes.error(error_code, "CreateFile", "Sharing violation")
    )
    monkeypatch.setattr(backend_module.win32file, "CreateFile", open_file)

    with pytest.raises(ValueError, match="another process"):
        service.open_document(str(path), read_only=True)

    service._active_app.assert_not_called()
    open_file.assert_called_once()
    assert path.read_bytes() == b"test document"


def test_file_availability_probe_uses_read_access_and_immediately_closes_handle(
    monkeypatch, tmp_path
):
    path = tmp_path / "available.docx"
    handle = Mock()
    open_file = Mock(return_value=handle)
    monkeypatch.setattr(backend_module.win32file, "CreateFile", open_file)

    backend_module._assert_file_available(path)

    open_file.assert_called_once_with(
        str(path),
        backend_module.win32con.GENERIC_READ,
        backend_module.win32con.FILE_SHARE_READ,
        None,
        backend_module.win32con.OPEN_EXISTING,
        backend_module.win32con.FILE_ATTRIBUTE_NORMAL,
        None,
    )
    handle.Close.assert_called_once_with()


def test_file_probe_preserves_other_access_error_details(monkeypatch, tmp_path):
    path = tmp_path / "denied.docx"
    monkeypatch.setattr(
        backend_module.win32file,
        "CreateFile",
        Mock(side_effect=backend_module.pywintypes.error(5, "CreateFile", "Access is denied")),
    )

    with pytest.raises(OSError, match="Access is denied") as error:
        backend_module._assert_file_available(path)

    assert error.value.filename == str(path)


def test_existing_registered_document_returns_without_a_second_file_open(
    word, monkeypatch, tmp_path
):
    service, _, _ = word
    path = tmp_path / "already-open.docx"
    path.write_bytes(b"test document")
    existing = {"document_id": "doc_existing", "path": str(path), "read_only": True}
    service.list_documents.return_value = {"documents": [existing]}
    service._active_app = Mock()
    probe = Mock(side_effect=AssertionError("An already registered document must not be reopened"))
    monkeypatch.setattr(backend_module, "_assert_file_available", probe)

    result = service.open_document(str(path), read_only=True)

    assert result is existing
    probe.assert_not_called()
    service._active_app.assert_not_called()

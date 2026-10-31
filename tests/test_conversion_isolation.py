"""Keep internal conversion Word state separate from user document operations."""

from contextlib import nullcontext
from copy import deepcopy
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

pytest.importorskip("pythoncom")

from axmath_mcp import backend as backend_module
from axmath_mcp.backend import WordBackend
from axmath_mcp.office import OfficeBackend


def app_proxy():
    return SimpleNamespace(_oleobj_=object(), Documents=SimpleNamespace(Count=1))


@pytest.fixture
def isolated_backend():
    private = app_proxy()
    service = WordBackend.__new__(WordBackend)
    service._apps = []
    service._workspace = SimpleNamespace(
        owns_app=Mock(side_effect=lambda app: app._oleobj_ == private._oleobj_),
        owns_document=Mock(return_value=False),
        scratch=Mock(),
    )
    return service, private


def test_private_converter_is_not_returned_as_the_active_user_word(isolated_backend, monkeypatch):
    service, private = isolated_backend
    monkeypatch.setattr(backend_module.win32com.client, "GetActiveObject", lambda _: private)
    create = Mock()
    monkeypatch.setattr(backend_module.win32com.client, "DispatchEx", create)

    assert service._active_app() is None
    assert service._apps == []
    create.assert_not_called()


def test_user_lookup_falls_back_to_cached_user_app_when_rot_points_at_converter(
    isolated_backend, monkeypatch
):
    service, private = isolated_backend
    user = app_proxy()
    service._apps = [private, user]
    monkeypatch.setattr(backend_module.win32com.client, "GetActiveObject", lambda _: private)

    assert service._active_app() is user


def test_create_user_word_uses_a_separate_instance_not_private_converter(
    isolated_backend, monkeypatch
):
    service, private = isolated_backend
    user = app_proxy()
    monkeypatch.setattr(backend_module.win32com.client, "GetActiveObject", lambda _: private)
    create = Mock(return_value=user)
    monkeypatch.setattr(backend_module.win32com.client, "DispatchEx", create)

    assert service._active_app(create=True) is user
    create.assert_called_once_with("Word.Application")
    assert service._apps == [user]


def test_registering_converter_as_user_host_is_refused(isolated_backend):
    service, private = isolated_backend
    with pytest.raises(RuntimeError, match="not a user document host"):
        service._remember_app(private)
    assert service._apps == []


def test_scratch_refuses_user_app_before_document_allocation(isolated_backend):
    service, _ = isolated_backend
    user = app_proxy()
    user.Documents.Add = Mock()

    with pytest.raises(RuntimeError, match="private conversion"), service._scratch(user):
        pytest.fail("A user application cannot host internal scratch documents")

    service._workspace.scratch.assert_not_called()
    user.Documents.Add.assert_not_called()


def test_owned_scratch_uses_workspace_lifecycle(isolated_backend):
    service, private = isolated_backend
    document = object()
    service._workspace.scratch.return_value = nullcontext(document)

    with service._scratch(private) as actual:
        assert actual is document
    service._workspace.scratch.assert_called_once_with()


def test_destructive_export_rejects_unowned_document_before_selection_or_macro(isolated_backend):
    service, private = isolated_backend
    document = SimpleNamespace(Activate=Mock(), Content=SimpleNamespace(Select=Mock()))
    service._run_macro = Mock()

    with pytest.raises(RuntimeError, match="owned scratch documents"):
        service._export_inplace(private, document)

    document.Activate.assert_not_called()
    document.Content.Select.assert_not_called()
    service._run_macro.assert_not_called()


def test_busy_vendor_preflight_does_not_start_private_word(monkeypatch):
    service = WordBackend.__new__(WordBackend)
    service._mutex = object()
    service._converter_app = Mock()
    service._ready = Mock()
    monkeypatch.setattr(
        backend_module.win32event,
        "WaitForSingleObject",
        lambda *_: backend_module.win32event.WAIT_OBJECT_0,
    )
    release = Mock()
    monkeypatch.setattr(backend_module.win32event, "ReleaseMutex", release)
    monkeypatch.setattr(
        backend_module,
        "_reg",
        lambda _root, _key, name="": "1" if name == "ConvertOccupy" else None,
    )

    with pytest.raises(RuntimeError, match="busy"), service._conversion():
        pytest.fail("A busy converter must be rejected before creating its Word instance")

    service._converter_app.assert_not_called()
    service._ready.assert_not_called()
    release.assert_called_once_with(service._mutex)


@pytest.mark.parametrize("method_name", ["insert_equation", "ppt_insert_equation"])
def test_completed_edit_returns_pending_cleanup_warning_without_mutating_backend_result(
    method_name,
):
    original = {"equation_id": "eq_done", "saved": False, "warnings": ["Existing note"]}
    before = deepcopy(original)
    operation = Mock(return_value=original)
    word = SimpleNamespace(
        conversion_status=Mock(
            return_value={"pending_count": 1, "cleanup_warnings": ["Internal cleanup pending"]}
        ),
    )
    ppt = SimpleNamespace()
    setattr(ppt if method_name.startswith("ppt_") else word, method_name, operation)
    office = OfficeBackend(word_backend=word, ppt_backend=ppt)

    result = getattr(office, method_name)("target", latex="x")

    operation.assert_called_once_with("target", latex="x")
    assert result["equation_id"] == "eq_done"
    assert result["temporary_documents_pending"] == 1
    assert result["warnings"] == ["Existing note", "Internal cleanup pending"]
    assert result is not original
    assert original == before


def test_success_without_pending_cleanup_keeps_original_result():
    original = {"equation_id": "eq_done", "saved": False}
    word = SimpleNamespace(
        insert_equation=Mock(return_value=original),
        conversion_status=Mock(return_value={"pending_count": 0, "cleanup_warnings": []}),
    )
    office = OfficeBackend(word_backend=word, ppt_backend=SimpleNamespace())

    assert office.insert_equation("target", latex="x") is original

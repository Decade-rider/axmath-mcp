"""Own-call recovery guards with a fake registry; never launches Office."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

pytest.importorskip("pythoncom")

from axmath_mcp import backend as backend_module
from axmath_mcp.backend import WordBackend

KEY = r"Software\AxMath\WordCmds"
DISP_E_EXCEPTION = -2147352567
TRANSPORT_ERRORS = [-2147418111, -2147417846, -2147417848, -2147220995]


class FakeRegistry:
    def __init__(self):
        self.values = {
            "ConvertOccupy": "0",
            "AxMathOccupy": "0",
            "DoVerbOccupy": "0",
            "WaitingConvert": "1",
        }
        self.writes = []
        self.call_in_progress = False
        self.write_error = None
        self.open_error = None

    def read(self, root, key, value=""):
        assert root == backend_module.winreg.HKEY_CURRENT_USER
        assert key == KEY
        return self.values.get(value)

    def open(self, root, key, reserved, access):
        assert not self.call_in_progress, "Recovery must not run while the native call is pending."
        assert root == backend_module.winreg.HKEY_CURRENT_USER and key == KEY
        assert reserved == 0 and access == backend_module.winreg.KEY_SET_VALUE
        if self.open_error:
            raise self.open_error
        return self

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False

    def write(self, handle, name, reserved, value_type, value):
        assert not self.call_in_progress, "Never clear a still-running macro's flag."
        assert handle is self
        assert (name, reserved, value_type, value) == (
            "ConvertOccupy",
            0,
            backend_module.winreg.REG_SZ,
            "0",
        )
        if self.write_error:
            raise self.write_error
        self.writes.append((name, value))
        self.values[name] = value

    def returned_failure(self, *_, **__):
        self.call_in_progress = True
        self.values["ConvertOccupy"] = "1"
        try:
            assert self.writes == []
            return False
        finally:
            self.call_in_progress = False

    def raising_failure(self, error):
        def run(*_):
            self.call_in_progress = True
            self.values["ConvertOccupy"] = "1"
            try:
                assert self.writes == []
                raise error
            finally:
                self.call_in_progress = False

        return run


@pytest.fixture
def owned_macro(monkeypatch):
    registry = FakeRegistry()
    app = SimpleNamespace(Run=Mock(), Documents=SimpleNamespace(Count=0))
    service = WordBackend.__new__(WordBackend)
    service._workspace = SimpleNamespace(owns_app=lambda candidate: candidate is app)
    service._conversion_app = app
    service._busy_recoveries = 0
    service._recovery_warnings = []
    monkeypatch.setattr(backend_module, "_reg", registry.read)
    monkeypatch.setattr(backend_module.winreg, "OpenKey", registry.open)
    monkeypatch.setattr(backend_module.winreg, "SetValueEx", registry.write)
    return service, app, registry


def vendor_error(hresult=DISP_E_EXCEPTION, *, nested_code=0):
    return backend_module.pywintypes.com_error(
        hresult, "vendor test failure", (0, "Word", "invalid formula", None, 0, nested_code), None
    )


def test_completed_owned_com_exception_releases_only_its_flag_and_preserves_error(owned_macro):
    service, app, registry = owned_macro
    error = vendor_error()
    app.Run.side_effect = registry.raising_failure(error)

    with pytest.raises(RuntimeError, match="macro failed") as caught:
        service._run_macro(app, "convert")

    assert caught.value.__cause__ is error
    assert "vendor test failure" in str(caught.value)
    assert registry.values["ConvertOccupy"] == "0"
    assert registry.writes == [("ConvertOccupy", "0")]
    assert service._busy_recoveries == 1
    assert service._recovery_warnings == []
    app.Run.assert_called_once_with("convert", False, False)


def test_completed_false_result_is_still_an_error_after_owned_flag_is_released(owned_macro):
    service, app, registry = owned_macro
    app.Run.side_effect = registry.returned_failure

    with pytest.raises(RuntimeError, match="did not convert"):
        service._run_macro(app, "convert")

    assert registry.values["ConvertOccupy"] == "0"
    assert service._busy_recoveries == 1
    app.Run.assert_called_once()


def test_success_with_vendor_cleanup_never_rewrites_registry(owned_macro):
    service, app, registry = owned_macro

    def successful(*_):
        registry.call_in_progress = True
        registry.values["ConvertOccupy"] = "1"
        assert registry.writes == []
        registry.values["ConvertOccupy"] = "0"
        registry.call_in_progress = False
        return True

    app.Run.side_effect = successful
    assert service._run_macro(app, "convert") is None
    assert registry.writes == []
    assert service._busy_recoveries == 0


def test_preexisting_busy_state_refuses_macro_and_is_never_cleared(owned_macro):
    service, app, registry = owned_macro
    registry.values["ConvertOccupy"] = "1"

    with pytest.raises(RuntimeError, match="became busy"):
        service._run_macro(app, "convert")

    app.Run.assert_not_called()
    assert registry.values["ConvertOccupy"] == "1" and registry.writes == []


@pytest.mark.parametrize(
    "missing_boundary", ["workspace", "private_app", "context", "other_context"]
)
def test_unowned_call_never_resets_global_occupancy(owned_macro, missing_boundary):
    service, app, registry = owned_macro
    if missing_boundary == "workspace":
        del service._workspace
    elif missing_boundary == "private_app":
        service._workspace.owns_app = lambda _: False
    elif missing_boundary == "context":
        service._conversion_app = None
    else:
        service._conversion_app = object()
    app.Run.side_effect = registry.returned_failure

    with pytest.raises(RuntimeError, match="did not convert"):
        service._run_macro(app, "convert")

    assert registry.values["ConvertOccupy"] == "1" and registry.writes == []
    assert service._busy_recoveries == 0


@pytest.mark.parametrize(
    "flag,value",
    [
        ("AxMathOccupy", "1"),
        ("DoVerbOccupy", "1"),
        ("WaitingConvert", "0"),
        ("AxMathOccupy", None),
        ("DoVerbOccupy", None),
        ("WaitingConvert", None),
    ],
)
def test_native_busy_or_unknown_completion_is_not_recovered(owned_macro, flag, value):
    service, app, registry = owned_macro
    registry.values[flag] = value
    app.Run.side_effect = registry.returned_failure

    with pytest.raises(RuntimeError, match="did not convert"):
        service._run_macro(app, "convert")

    assert registry.values["ConvertOccupy"] == "1" and registry.writes == []
    assert service._busy_recoveries == 0


@pytest.mark.parametrize("hresult", TRANSPORT_ERRORS)
def test_transport_errors_do_not_prove_that_native_work_finished(owned_macro, hresult):
    service, app, registry = owned_macro
    error = vendor_error(hresult)
    app.Run.side_effect = registry.raising_failure(error)

    with pytest.raises(RuntimeError, match="macro failed") as caught:
        service._run_macro(app, "convert")

    assert caught.value.__cause__ is error
    assert registry.values["ConvertOccupy"] == "1" and registry.writes == []
    assert service._busy_recoveries == 0


@pytest.mark.parametrize("transport_code", TRANSPORT_ERRORS)
@pytest.mark.parametrize(
    "location,unsigned", [("hresult", True), ("nested", False), ("nested", True)]
)
def test_wrapped_or_unsigned_transport_codes_are_also_not_recovered(
    owned_macro, transport_code, location, unsigned
):
    service, app, registry = owned_macro
    code = transport_code + 2**32 if unsigned else transport_code
    error = vendor_error(code) if location == "hresult" else vendor_error(nested_code=code)
    app.Run.side_effect = registry.raising_failure(error)

    with pytest.raises(RuntimeError, match="macro failed") as caught:
        service._run_macro(app, "convert")

    assert caught.value.__cause__ is error
    assert registry.values["ConvertOccupy"] == "1" and registry.writes == []
    assert service._busy_recoveries == 0


def test_unknown_initial_flag_is_not_claimed_as_owned(owned_macro):
    service, app, registry = owned_macro
    registry.values["ConvertOccupy"] = None
    app.Run.side_effect = registry.returned_failure
    with pytest.raises(RuntimeError, match="did not convert"):
        service._run_macro(app, "convert")
    assert registry.writes == [] and service._busy_recoveries == 0


@pytest.mark.parametrize("failure", ["open", "write"])
def test_registry_recovery_failure_is_diagnostic_and_does_not_mask_original_error(
    owned_macro, failure
):
    service, app, registry = owned_macro
    error = vendor_error()
    app.Run.side_effect = registry.raising_failure(error)
    if failure == "open":
        registry.open_error = OSError("cannot open recovery key")
    else:
        registry.write_error = OSError("cannot write recovery flag")

    with pytest.raises(RuntimeError, match="macro failed") as caught:
        service._run_macro(app, "convert")

    assert caught.value.__cause__ is error
    assert registry.values["ConvertOccupy"] == "1"
    assert service._busy_recoveries == 0
    assert len(service._recovery_warnings) == 1
    assert "Could not release completed private conversion state" in service._recovery_warnings[0]


def test_unresponsive_private_word_is_not_recovered(owned_macro):
    service, app, registry = owned_macro

    class DisconnectedDocuments:
        @property
        def Count(self):
            raise vendor_error(-2147417848)

    app.Documents = DisconnectedDocuments()
    app.Run.side_effect = registry.returned_failure
    with pytest.raises(RuntimeError, match="did not convert"):
        service._run_macro(app, "convert")
    assert registry.writes == [] and service._busy_recoveries == 0
    assert service._recovery_warnings


def test_public_gcd_insert_is_rejected_before_documents_or_converter_are_accessed():
    service = WordBackend.__new__(WordBackend)
    service._doc = Mock(side_effect=AssertionError("Do not inspect user documents yet"))
    service._converter_app = Mock()
    service._conversion = Mock()

    with pytest.raises(ValueError, match=r"\\gcd"):
        service.insert_equation("unopened-document", r"\gcd(a,b)")

    service._doc.assert_not_called()
    service._converter_app.assert_not_called()
    service._conversion.assert_not_called()


def test_internal_generation_rechecks_gcd_before_allocating_scratch():
    service = WordBackend.__new__(WordBackend)
    service._scratch = Mock()
    with pytest.raises(ValueError, match=r"\\gcd"), service._generated(object(), r"\(\gcd(a,b)\)"):
        pytest.fail("Blocked input must never reach the converter")
    service._scratch.assert_not_called()

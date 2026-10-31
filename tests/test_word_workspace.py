"""Private conversion ownership and failure cleanup without launching Word."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

pytest.importorskip("pythoncom")

from axmath_mcp import word_workspace as workspace_module
from axmath_mcp.word_workspace import ConversionWorkspace, configure_private_printer


def com_error(code=-2147418111):
    return workspace_module.pywintypes.com_error(code, "Simulated Office error", None, None)


class Documents:
    def __init__(self):
        self.values = []
        self.add_calls = 0
        self.plans = []
        self.count_error = None

    @property
    def Count(self):
        if self.count_error:
            raise self.count_error
        return len(self.values)

    def Add(self, *, Visible):
        assert Visible is False
        self.add_calls += 1
        doc = Document(self)
        if self.plans:
            for name, value in self.plans.pop(0).items():
                setattr(doc, name, value)
        self.values.append(doc)
        return doc


class Document:
    def __init__(self, collection):
        self.collection = collection
        self._oleobj_ = object()
        self.track_error = None
        self.activate_error = None
        self.close_effects = []
        self.close_calls = 0
        self.activate_calls = 0

    @property
    def TrackRevisions(self):
        return False

    @TrackRevisions.setter
    def TrackRevisions(self, value):
        assert value is False
        if self.track_error:
            raise self.track_error

    def Activate(self):
        self.activate_calls += 1
        if self.activate_error:
            raise self.activate_error

    def Close(self, save):
        assert save == 0
        self.close_calls += 1
        if self.close_effects:
            effect = self.close_effects.pop(0)
            if isinstance(effect, BaseException):
                raise effect
        if self in self.collection.values:
            self.collection.values.remove(self)


@pytest.fixture
def private_word():
    app = SimpleNamespace(Documents=Documents(), Visible=True, Quit=Mock(), _oleobj_=object())
    factory = Mock(return_value=app)
    configure = Mock(return_value={"configured": True, "warnings": []})
    pump = Mock()
    sleep = Mock()
    workspace = ConversionWorkspace(
        factory=factory, configurator=configure, pump=pump, sleeper=sleep
    )
    return workspace, app, factory, configure, pump, sleep


def test_no_application_or_printer_is_touched_until_conversion(private_word):
    workspace, app, factory, configure, _, _ = private_word
    assert workspace.status()["word_started"] is False
    assert not workspace.owns_app(app)
    factory.assert_not_called()
    configure.assert_not_called()
    workspace.shutdown()
    app.Quit.assert_not_called()


def test_private_app_is_created_and_configured_once_before_document_add(private_word):
    workspace, app, factory, configure, _, _ = private_word
    configure.side_effect = lambda instance: {
        "configured": True,
        "count_before": instance.Documents.Count,
    }
    with workspace.scratch() as first:
        assert workspace.owns_document(first)
        assert workspace.status()["printer"]["count_before"] == 0
    with workspace.scratch():
        pass
    factory.assert_called_once_with()
    configure.assert_called_once_with(app)
    assert app.Visible is False
    assert app.Documents.Count == 0
    assert workspace.status()["temporary_document_count"] == 0


def test_ownership_uses_identity_not_names_or_visibility(private_word):
    workspace, app, _, _, _, _ = private_word
    workspace.get_app()
    assert workspace.owns_app(SimpleNamespace(_oleobj_=app._oleobj_))
    assert not workspace.owns_app(SimpleNamespace(_oleobj_=object()))
    with workspace.scratch() as document:
        assert workspace.owns_document(SimpleNamespace(_oleobj_=document._oleobj_))
        assert not workspace.owns_document(SimpleNamespace(_oleobj_=object(), Name="Document1"))
    assert not workspace.owns_document(document)


@pytest.mark.parametrize("property_name", ["track_error", "activate_error"])
def test_document_is_tracked_before_any_setter_or_activation(private_word, property_name):
    workspace, app, _, _, _, _ = private_word
    app.Documents.plans = [{property_name: RuntimeError("setup failed")}]
    with pytest.raises(RuntimeError, match="setup failed"), workspace.scratch():
        pytest.fail("Setup should fail before yielding")
    assert app.Documents.Count == 0
    assert workspace.status()["temporary_document_count"] == 0


def test_nested_scratch_never_closes_an_active_source_document(private_word):
    workspace, app, _, _, _, _ = private_word
    with workspace.scratch() as source:
        with workspace.scratch() as copy:
            assert app.Documents.Count == 2
            assert workspace.owns_document(source) and workspace.owns_document(copy)
            workspace.cleanup()
            assert source.close_calls == copy.close_calls == 0
        assert app.Documents.Count == 1
        assert source.close_calls == 0
    assert app.Documents.Count == 0


@pytest.mark.parametrize("code", [-2147418111, -2147417846])
def test_transient_close_is_retried_without_allocating_another_document(private_word, code):
    workspace, app, _, _, pump, sleep = private_word
    with workspace.scratch() as doc:
        doc.close_effects = [com_error(code), com_error(code), None]
    assert doc.close_calls == 3
    assert app.Documents.add_calls == 1
    assert pump.call_count == sleep.call_count == 2
    assert all(call.args == (0.1,) for call in sleep.call_args_list)
    assert workspace.status()["pending_count"] == 0
    assert workspace.status()["cleanup_warnings"] == []


def test_nested_rpc_code_is_also_recognized_as_transient(private_word):
    workspace, _, _, _, pump, _ = private_word
    wrapped = workspace_module.pywintypes.com_error(
        -2147352567, "Word error", (0, "Word", "busy", None, 0, -2147417846), None
    )
    with workspace.scratch() as doc:
        doc.close_effects = [wrapped, None]
    assert doc.close_calls == 2
    pump.assert_called_once_with()


def test_exhausted_close_remains_tracked_and_cannot_mask_completed_edit(private_word):
    workspace, app, _, _, _, _ = private_word

    def edit():
        with workspace.scratch() as doc:
            doc.close_effects = [com_error()] * 3
            return {"equation_inserted": True}

    assert edit() == {"equation_inserted": True}
    assert app.Documents.Count == 1
    assert workspace.status()["pending_count"] == 1
    assert "cleanup is pending" in workspace.status()["cleanup_warnings"][0]


def test_persistent_close_error_is_not_retried_and_preserves_primary_error(private_word):
    workspace, _, _, _, pump, _ = private_word
    with pytest.raises(ValueError, match="conversion failed"), workspace.scratch() as doc:
        doc.close_effects = [com_error(-2147024891)]
        raise ValueError("conversion failed")
    assert doc.close_calls == 1
    pump.assert_not_called()
    assert workspace.status()["pending_count"] == 1


def test_next_scratch_refuses_allocation_until_old_document_can_close(private_word):
    workspace, app, _, _, _, _ = private_word
    with workspace.scratch() as old:
        old.close_effects = [com_error(-2147024891), com_error(-2147024891)]
    with (
        pytest.raises(RuntimeError, match="No additional conversion document"),
        workspace.scratch(),
    ):
        pytest.fail("Pending cleanup must block allocation")
    assert app.Documents.add_calls == 1
    with workspace.scratch() as new:
        assert old not in app.Documents.values
        assert workspace.owns_document(new)
    assert workspace.status()["pending_count"] == 0
    assert workspace.status()["cleanup_warnings"] == []


def test_known_disconnected_document_is_forgotten(private_word):
    workspace, app, _, _, pump, _ = private_word
    with workspace.scratch() as doc:
        app.Documents.values.remove(doc)  # Word has already removed the object.
        doc.close_effects = [com_error(-2147417848)]
    assert workspace.status()["pending_count"] == 0
    pump.assert_not_called()


@pytest.mark.parametrize("code", [-2147417848, -2147220995])
def test_disconnected_idle_app_can_be_recreated_without_quitting_old_app(private_word, code):
    workspace, old, factory, configure, _, _ = private_word
    workspace.get_app()
    replacement = SimpleNamespace(
        Documents=Documents(), Visible=True, Quit=Mock(), _oleobj_=object()
    )
    old.Documents.count_error = com_error(code)
    factory.return_value = replacement

    assert workspace.get_app() is replacement

    assert factory.call_count == configure.call_count == 2
    assert replacement.Visible is False
    assert not workspace.owns_app(old)
    assert workspace.owns_app(replacement)
    old.Quit.assert_not_called()
    assert "interface disconnected" in workspace.status()["cleanup_warnings"][0]


@pytest.mark.parametrize("code", [-2147418111, -2147417846, -2147024891])
def test_busy_or_access_error_does_not_recreate_or_abandon_app(private_word, code):
    workspace, app, factory, _, _, _ = private_word
    workspace.get_app()
    app.Documents.count_error = com_error(code)

    with pytest.raises(workspace_module.pywintypes.com_error):
        workspace.get_app()

    factory.assert_called_once_with()
    assert workspace.owns_app(app)
    app.Quit.assert_not_called()


def test_disconnection_during_active_scratch_does_not_start_another_word(private_word):
    workspace, app, factory, _, _, _ = private_word
    with workspace.scratch() as document:
        app.Documents.count_error = com_error(-2147417848)
        with pytest.raises(RuntimeError, match="during an active conversion"):
            workspace.get_app()
        assert workspace.owns_document(document)
        factory.assert_called_once_with()
    assert document.close_calls == 1
    app.Quit.assert_not_called()


def test_disconnected_app_with_unclosable_owned_doc_does_not_allocate_replacement(private_word):
    workspace, app, factory, _, _, _ = private_word
    with workspace.scratch() as document:
        document.close_effects = [com_error(-2147024891), com_error(-2147024891)]
    app.Documents.count_error = com_error(-2147417848)

    with pytest.raises(RuntimeError, match="pending cleanup"):
        workspace.get_app()

    assert workspace.owns_document(document)
    assert workspace.status()["pending_count"] == 1
    factory.assert_called_once_with()
    app.Quit.assert_not_called()


def test_deleted_document_error_in_excepinfo_is_forgotten_without_retry(private_word):
    workspace, app, _, _, pump, _ = private_word
    deleted = workspace_module.pywintypes.com_error(
        -2147352567,
        "Word error",
        (0, "Word", "Object has been deleted", None, 0, -2146822463),
        None,
    )
    with workspace.scratch() as document:
        app.Documents.values.remove(document)
        document.close_effects = [deleted]
    assert not workspace.owns_document(document)
    assert workspace.status()["pending_count"] == 0
    assert document.close_calls == 1
    pump.assert_not_called()


def test_shutdown_retries_only_owned_document_then_quits_empty_private_app(private_word):
    workspace, app, _, _, _, _ = private_word
    with workspace.scratch() as doc:
        doc.close_effects = [com_error(-2147024891)]
    workspace.shutdown()
    assert doc.close_calls == 2
    app.Quit.assert_called_once_with(0)
    assert not workspace.owns_app(app)
    assert workspace.status()["temporary_document_count"] == 0
    with pytest.raises(RuntimeError, match="shut down"):
        workspace.get_app()


def test_shutdown_with_unresolved_close_never_forces_quit(private_word):
    workspace, app, _, _, _, _ = private_word
    with workspace.scratch() as doc:
        doc.close_effects = [com_error(-2147024891), com_error(-2147024891)]
    workspace.shutdown()
    assert workspace.owns_document(doc)
    assert workspace.status()["pending_count"] == 1
    assert app.Documents.Count == 1
    app.Quit.assert_not_called()


def test_non_com_cleanup_failure_is_retained_instead_of_losing_ownership(private_word):
    workspace, app, _, _, _, _ = private_word
    with workspace.scratch() as doc:
        doc.close_effects = [RuntimeError("unexpected cleanup error")]
    assert workspace.owns_document(doc)
    assert workspace.status()["pending_count"] == 1
    assert "unexpected cleanup error" in workspace.status()["cleanup_warnings"][0]
    workspace.cleanup()
    assert app.Documents.Count == 0


def test_shutdown_does_not_quit_or_hide_private_app_with_untracked_user_document(private_word):
    workspace, app, _, _, _, _ = private_word
    with workspace.scratch():
        foreign = Document(app.Documents)
        app.Documents.values.append(foreign)
        app.Visible = True  # The user opened a real document in this process.
    workspace.shutdown()
    assert app.Visible is True
    assert app.Documents.values == [foreign]
    assert foreign.close_calls == 0
    app.Quit.assert_not_called()


def test_factory_returning_nonempty_word_is_never_hidden_or_adopted(private_word):
    workspace, app, _, configure, _, _ = private_word
    foreign = Document(app.Documents)
    app.Documents.values.append(foreign)
    with pytest.raises(RuntimeError, match="existing documents were left untouched"):
        workspace.get_app()
    assert app.Visible is True
    assert not workspace.owns_app(app)
    configure.assert_not_called()
    workspace.shutdown()
    app.Quit.assert_not_called()
    assert foreign.close_calls == 0


def test_printer_setup_failure_prevents_scratch_creation(private_word):
    workspace, app, _, configure, _, _ = private_word
    configure.side_effect = RuntimeError("printer setup failed")
    with pytest.raises(RuntimeError, match="printer setup failed"), workspace.scratch():
        pytest.fail("No temporary document should be allocated")
    assert app.Documents.add_calls == 0
    app.Quit.assert_called_once_with(0)


@pytest.fixture
def printer_app(monkeypatch):
    monkeypatch.delenv("AXMATH_MCP_PRINTER", raising=False)
    defaults = Mock(side_effect=["HP,winspool,Ne01:", "HP,winspool,Ne01:"])
    monkeypatch.setattr(workspace_module, "_windows_default_printer", defaults)
    monkeypatch.setattr(
        workspace_module, "_installed_printers", lambda: ["HP", "Microsoft Print to PDF"]
    )

    class App:
        def __init__(self):
            self.dialog = SimpleNamespace(Execute=Mock(return_value=0))
            self.Dialogs = Mock(return_value=self.dialog)

        @property
        def ActivePrinter(self):
            raise AssertionError("Do not query printer devices through ActivePrinter")

        @ActivePrinter.setter
        def ActivePrinter(self, value):
            raise AssertionError("Never assign ActivePrinter and alter Windows defaults")

    return App(), defaults


def test_local_pdf_printer_uses_non_global_dialog_not_activeprinter(printer_app):
    app, defaults = printer_app
    result = configure_private_printer(app)
    app.Dialogs.assert_called_once_with(97)
    assert app.dialog.Printer == "Microsoft Print to PDF"
    assert app.dialog.DoNotSetAsSysDefault is True
    app.dialog.Execute.assert_called_once_with()
    assert defaults.call_count == 2
    assert result["configured"] is True
    assert result["windows_default_unchanged"] is True


def test_explicit_installed_printer_is_resolved_case_insensitively(printer_app, monkeypatch):
    app, _ = printer_app
    monkeypatch.setenv("AXMATH_MCP_PRINTER", "hp")
    result = configure_private_printer(app)
    assert app.dialog.Printer == "HP"
    assert result["source"] == "environment"


def test_absent_pdf_keeps_current_printer_and_reports_limitation(printer_app, monkeypatch):
    app, defaults = printer_app
    monkeypatch.setattr(workspace_module, "_installed_printers", lambda: ["HP"])
    result = configure_private_printer(app)
    app.Dialogs.assert_not_called()
    defaults.assert_not_called()
    assert result["configured"] is False and result["source"] == "existing"
    assert "unavailable printer" in result["warnings"][0]


def test_missing_explicit_printer_fails_before_dialog(printer_app, monkeypatch):
    app, _ = printer_app
    monkeypatch.setenv("AXMATH_MCP_PRINTER", "Unknown printer")
    with pytest.raises(ValueError, match="does not name an installed printer"):
        configure_private_printer(app)
    app.Dialogs.assert_not_called()


def test_unexpected_system_default_change_is_reported_without_global_restore(printer_app):
    app, defaults = printer_app
    defaults.side_effect = ["HP,winspool,Ne01:", "Other,winspool,Ne02:"]
    with pytest.raises(RuntimeError, match="Windows default printer changed"):
        configure_private_printer(app)
    app.dialog.Execute.assert_called_once_with()


def test_unreadable_default_does_not_claim_verified_unchanged(printer_app):
    app, defaults = printer_app
    defaults.side_effect = [None, None]
    result = configure_private_printer(app)
    assert result["configured"] is True
    assert result["windows_default_unchanged"] is None
    assert result["warnings"]

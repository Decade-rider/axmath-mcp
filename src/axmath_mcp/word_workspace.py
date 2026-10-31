"""Own a private Word instance and strictly track disposable conversion documents.

This workspace never attaches to an existing application, closes untracked
documents, kills Word, or changes the Windows default printer. Cleanup failures
remain observable without changing the outcome of an already completed edit.
"""

from __future__ import annotations

import logging
import os
import time
import winreg
from collections.abc import Callable
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Any

import pythoncom
import pywintypes
import win32com.client

logger = logging.getLogger(__name__)

TRANSIENT_CLOSE_ERRORS = frozenset({-2147418111, -2147417846})
# RPC_E_DISCONNECTED and CO_E_OBJNOTCONNECTED: these identify disconnected
# proxies, unlike busy/retry-later or access errors which must retain ownership.
DISCONNECTED_ERRORS = frozenset({-2147417848, -2147220995})
# Word error 5825, "Object has been deleted" (including nested EXCEPINFO).
DELETED_DOCUMENT_ERRORS = frozenset({-2146822463})
DEFAULT_LOCAL_PRINTER = "Microsoft Print to PDF"


def _windows_default_printer() -> str | None:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
        ) as key:
            return str(winreg.QueryValueEx(key, "Device")[0])
    except OSError:
        return None


def _installed_printers() -> list[str]:
    names = []
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Devices"
        ) as key:
            for index in range(winreg.QueryInfoKey(key)[1]):
                names.append(str(winreg.EnumValue(key, index)[0]))
    except OSError:
        return []
    return names


def configure_private_printer(app: Any) -> dict[str, Any]:
    """Select a local printer for this Word instance using the non-global dialog.

    Printer discovery and the system default check are registry reads. Importantly,
    Word.Application.ActivePrinter is not assigned: that API may alter Windows'
    default printer. The dialog's DoNotSetAsSysDefault flag is explicit instead.
    """
    requested = os.environ.get("AXMATH_MCP_PRINTER", "").strip()
    names = _installed_printers()
    desired = requested or DEFAULT_LOCAL_PRINTER
    selected = next((name for name in names if name.casefold() == desired.casefold()), None)
    result: dict[str, Any] = {
        "configured": False,
        "source": "environment" if requested else "local_pdf",
        "requested_printer": requested or None,
        "selected_printer": selected,
        "windows_default_unchanged": None,
        "warnings": [],
    }
    if selected is None:
        if requested:
            raise ValueError(
                "AXMATH_MCP_PRINTER does not name an installed printer. "
                "Choose an installed printer or remove this environment variable."
            )
        result["source"] = "existing"
        result["warnings"].append(
            "Microsoft Print to PDF is not registered; the private Word instance keeps "
            "its current printer. An unavailable printer may delay layout."
        )
        return result

    before = _windows_default_printer()
    dialog = app.Dialogs(97)  # wdDialogFilePrintSetup
    dialog.Printer = selected
    dialog.DoNotSetAsSysDefault = True
    dialog.Execute()
    after = _windows_default_printer()
    if before is not None and after is not None:
        result["windows_default_unchanged"] = before == after
        if before != after:
            raise RuntimeError(
                "Windows default printer changed during private Word printer setup. "
                "No conversion document was created; inspect printer settings before retrying."
            )
    else:
        result["warnings"].append("Could not verify the Windows default printer from the registry.")
    result["configured"] = True
    return result


def _new_word() -> Any:
    return win32com.client.DispatchEx("Word.Application")


def _same_com_object(first: Any, second: Any) -> bool:
    if first is second:
        return True
    try:
        return first._oleobj_ == second._oleobj_
    except (AttributeError, pywintypes.com_error):
        return False


def _error_codes(error: pywintypes.com_error) -> set[int]:
    values = [error.hresult]
    if error.excepinfo and len(error.excepinfo) > 5:
        values.append(error.excepinfo[5])
    return {value if value < 2**31 else value - 2**32 for value in values if isinstance(value, int)}


@dataclass(eq=False)
class _OwnedDocument:
    document: Any
    token: str
    active: bool = True


class ConversionWorkspace:
    """Use only from the owning Office STA; no COM references cross threads."""

    def __init__(
        self,
        *,
        factory: Callable[[], Any] = _new_word,
        configurator: Callable[[Any], dict[str, Any]] = configure_private_printer,
        pump: Callable[[], Any] = pythoncom.PumpWaitingMessages,
        sleeper: Callable[[float], Any] = time.sleep,
    ) -> None:
        self._factory = factory
        self._configurator = configurator
        self._pump = pump
        self._sleep = sleeper
        self._app: Any | None = None
        self._documents: list[_OwnedDocument] = []
        self._warnings: dict[str, str] = {}
        self._printer: dict[str, Any] = {
            "configured": False,
            "source": "not_started",
            "warnings": [],
        }
        self._counter = 0
        self._closed = False

    def get_app(self) -> Any:
        if self._closed:
            raise RuntimeError("The private Word conversion workspace has shut down.")
        if self._app is not None:
            try:
                _ = self._app.Documents.Count
            except pywintypes.com_error as exc:
                if not (_error_codes(exc) & DISCONNECTED_ERRORS):
                    raise
                if any(record.active for record in self._documents):
                    raise RuntimeError(
                        "Private Word disconnected during an active conversion. "
                        "No replacement instance was started; inspect the operation before retrying."
                    ) from exc
                self.cleanup()
                if self._documents:
                    raise RuntimeError(
                        "Private Word disconnected with temporary documents pending cleanup. "
                        "No replacement instance was started."
                    ) from exc
                # Do not call Quit on a disconnected interface: it may no longer
                # describe a process we can inspect safely. A fresh DispatchEx
                # will still be required to return an empty application.
                self._app = None
                self._warnings["disconnected_application"] = (
                    "The previous private Word interface disconnected. Its tracked proxies "
                    "were released; no untracked document or Word process was closed."
                )
            else:
                return self._app
        app = self._factory()
        # A replacement/injected factory must not allow us to hide somebody's
        # existing documents. DispatchEx normally starts empty desktop Word.
        if int(app.Documents.Count) != 0:
            raise RuntimeError(
                "Expected an empty private Word instance; existing documents were left untouched."
            )
        self._app = app
        try:
            app.Visible = False
            self._printer = self._configurator(app)
        except BaseException:
            self._quit_if_empty()
            if self._app is not None:
                self._closed = True
            raise
        return app

    def owns_app(self, app: Any) -> bool:
        return self._app is not None and _same_com_object(app, self._app)

    def owns_document(self, doc: Any) -> bool:
        return any(_same_com_object(doc, record.document) for record in self._documents)

    def status(self) -> dict[str, Any]:
        return {
            "word_started": self._app is not None,
            "temporary_document_count": len(self._documents),
            "pending_count": sum(not record.active for record in self._documents),
            "cleanup_warnings": list(self._warnings.values()),
            "printer": {**self._printer, "warnings": list(self._printer.get("warnings", []))},
        }

    def _forget(self, record: _OwnedDocument) -> None:
        if record in self._documents:
            self._documents.remove(record)
        self._warnings.pop(record.token, None)

    def _close(self, record: _OwnedDocument) -> bool:
        for attempt in range(3):
            try:
                record.document.Close(0)
            except pywintypes.com_error as exc:
                codes = _error_codes(exc)
                if codes & (DISCONNECTED_ERRORS | DELETED_DOCUMENT_ERRORS):
                    self._forget(record)
                    return True
                if codes & TRANSIENT_CLOSE_ERRORS and attempt < 2:
                    try:
                        self._pump()
                        self._sleep(0.1)
                    except Exception as wait_error:  # noqa: BLE001 - preserve completed edits if cleanup cannot wait.
                        message = f"Cleanup retry for {record.token} could not wait: {wait_error!r}"
                        self._warnings[record.token] = message
                        logger.warning(message)
                        return False
                    continue
                message = (
                    f"Could not close owned temporary document {record.token}; "
                    f"cleanup is pending (COM codes {sorted(codes)}). "
                    "Do not repeat a completed formula edit."
                )
                self._warnings[record.token] = message
                logger.warning(message)
                return False
            except Exception as exc:  # noqa: BLE001 - cleanup cannot erase a completed edit's outcome.
                # Bookkeeping must survive unexpected cleanup exceptions too;
                # completed target edits must not be reported as failed edits.
                message = f"Cleanup of owned temporary document {record.token} is pending: {exc!r}"
                self._warnings[record.token] = message
                logger.warning(message)
                return False
            else:
                self._forget(record)
                return True
        return False

    def cleanup(self) -> list[str]:
        """Retry inactive owned documents, never one currently used by a caller."""
        for record in list(self._documents):
            if not record.active:
                self._close(record)
        return list(self._warnings.values())

    @contextmanager
    def scratch(self):
        app = self.get_app()
        self.cleanup()
        if any(not record.active for record in self._documents):
            raise RuntimeError(
                "A previous temporary Word document is still waiting for cleanup. "
                "No additional conversion document was created; resolve Word dialogs and retry."
            )
        doc = app.Documents.Add(Visible=False)
        self._counter += 1
        record = _OwnedDocument(doc, f"scratch_{self._counter}")
        self._documents.append(record)  # Before any property setter or activation.
        try:
            doc.TrackRevisions = False
            doc.Activate()
            yield doc
        finally:
            record.active = False
            self._close(record)

    def _quit_if_empty(self) -> None:
        app = self._app
        if app is None:
            return
        try:
            if int(app.Documents.Count) != 0:
                self._warnings["application"] = (
                    "Private Word still contains documents; it was left open. "
                    "Untracked documents are never closed by the conversion workspace."
                )
                return
            app.Quit(0)
        except pywintypes.com_error as exc:
            self._warnings["application"] = f"Could not close empty private Word instance: {exc!r}"
            logger.warning(self._warnings["application"])
            return
        except Exception as exc:  # noqa: BLE001 - never force application termination on cleanup failure.
            self._warnings["application"] = f"Private Word shutdown check failed: {exc!r}"
            logger.warning(self._warnings["application"])
            return
        self._app = None
        self._documents.clear()
        self._warnings.clear()

    def shutdown(self) -> None:
        """Close owned temporaries; quit only a confirmed-empty private instance."""
        self._closed = True
        for record in list(self._documents):
            record.active = False
            self._close(record)
        self._quit_if_empty()

"""Windows Word COM backend. All methods must run on the owning STA thread.

Conversions use the user's installed AxMath add-in, not redistributed vendor code.
No macro-security settings are relaxed and no existing document is saved implicitly.
"""

from __future__ import annotations

import logging
import os
import uuid
import winreg
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pywintypes
import win32api
import win32com.client
import win32con
import win32event
import win32file

from axmath_mcp import __version__
from axmath_mcp.axmath_compat import check_axmath_input
from axmath_mcp.font_metrics import equation_font_size
from axmath_mcp.layout import (
    check_block_fits,
    describe_layout,
    inline_paragraph,
    managed_target,
    place_equation,
    preflight_numbers,
    refresh_numbers,
)
from axmath_mcp.validation import (
    latex_for_layout,
    normalize_latex,
    validate_layout_options,
    validate_local_path,
)
from axmath_mcp.word_transfer import copy_native_range
from axmath_mcp.word_workspace import ConversionWorkspace

logger = logging.getLogger(__name__)

PROGID = "Equation.AxMath"
TO_AXMATH = "AMCCTex2AMEqn.ConvertScript2AMEqn"
TO_LATEX = "AMCCAMEqn2TeX.ConvertAMEqn2Script"


def _reg(root: int, key: str, value: str = "") -> str | None:
    try:
        with winreg.OpenKey(root, key) as handle:
            return str(winreg.QueryValueEx(handle, value)[0])
    except OSError:
        return None


def _same(a: Any, b: Any) -> bool:
    try:
        return a._oleobj_ == b._oleobj_
    except (AttributeError, pywintypes.com_error):
        return False


def _assert_file_available(path: Path) -> None:
    """Avoid Word's blocking file-in-use dialog before opening an unknown file."""
    try:
        handle = win32file.CreateFile(
            str(path),
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )
    except pywintypes.error as exc:
        if exc.winerror in (32, 33):
            raise ValueError(
                "The file is open or locked in another process. Close it or open a copy; "
                "use word_list_documents to select an already reachable document."
            ) from exc
        raise OSError(exc.winerror, "Cannot access document: " + exc.strerror, str(path)) from exc
    else:
        handle.Close()


class WordBackend:
    def __init__(self) -> None:
        self._apps: list[Any] = []
        self._documents: dict[str, Any] = {}
        self._equations: dict[str, tuple[str, Any]] = {}
        self._workspace = ConversionWorkspace()
        self._conversion_app = None
        self._restore_state_warnings: list[str] = []
        self._recovery_warnings: list[str] = []
        self._busy_recoveries = 0
        self._mutex = win32event.CreateMutex(None, False, "Local\\AxMathMCP.Automation")

    def shutdown(self) -> None:
        # Close only tracked internal scratch documents. User documents, including
        # files explicitly created by tools, remain owned by the user.
        try:
            self._workspace.shutdown()
        finally:
            self._equations.clear()
            self._documents.clear()
            self._apps.clear()
            win32api.CloseHandle(self._mutex)

    def conversion_status(self) -> dict[str, Any]:
        status = self._workspace.status()
        status["cleanup_warnings"] += self._restore_state_warnings + self._recovery_warnings
        status["recovered_owned_busy_flags"] = self._busy_recoveries
        return status

    def _converter_app(self) -> Any:
        # Remember the user's current Word instance before the private converter
        # has a chance to become Word's active COM registration.
        self._active_app()
        return self._workspace.get_app()

    def health(self) -> dict[str, Any]:
        server = _reg(winreg.HKEY_CLASSES_ROOT, PROGID + r"\protocol\StdFileEditing\server")
        license_flag = _reg(winreg.HKEY_CURRENT_USER, r"Software\AxMath\AuthUser", "Licensed")
        startup = Path(os.environ.get("APPDATA", "")) / "Microsoft/Word/STARTUP/AxMath.dotm"
        return {
            "platform": "win32",
            "version": __version__,
            "conversion_workspace": self.conversion_status(),
            "word_registered": bool(_reg(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID")),
            "axmath_registered": bool(server),
            "axmath_executable": server,
            "axmath_executable_exists": bool(server and Path(server.strip('"')).is_file()),
            "startup_template": str(startup),
            "startup_template_exists": startup.is_file(),
            "license_flag": license_flag,
            "conversion_busy": _reg(
                winreg.HKEY_CURRENT_USER, r"Software\AxMath\WordCmds", "ConvertOccupy"
            )
            == "1",
            "supported": "Windows Word; native AxMath with inline, display and SEQ-numbered layouts",
            "macro_entry": TO_AXMATH,
            "notes": [
                "Does not change macro security or licensing.",
                "Conversion/export temporarily uses Word selection and the clipboard.",
                "Other separately launched Word instances may require opening the file explicitly.",
            ],
        }

    def _remember_app(self, app: Any) -> Any:
        if self._workspace.owns_app(app):
            raise RuntimeError("The internal conversion Word instance is not a user document host.")
        if not any(_same(app, previous) for previous in self._apps):
            self._apps.append(app)
        return app

    def _active_app(self, create: bool = False) -> Any | None:
        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except pywintypes.com_error:
            app = None
        if app is not None and not self._workspace.owns_app(app):
            return self._remember_app(app)
        for previous in self._apps:
            try:
                if not self._workspace.owns_app(previous):
                    _ = previous.Documents.Count
                    return previous
            except pywintypes.com_error:
                continue
        if create:
            return self._remember_app(win32com.client.DispatchEx("Word.Application"))
        return None

    def _remember_doc(self, doc: Any) -> str:
        for document_id, previous in self._documents.items():
            if _same(doc, previous):
                return document_id
        document_id = "doc_" + uuid.uuid4().hex
        self._documents[document_id] = doc
        return document_id

    def _doc(self, document_id: str) -> Any:
        doc = self._documents.get(document_id)
        if doc is None:
            raise ValueError("Unknown document_id; call word_list_documents or open the document.")
        try:
            _ = doc.Name
        except pywintypes.com_error as exc:
            raise ValueError("This document was closed. List/open documents again.") from exc
        return doc

    def _info(self, doc: Any) -> dict[str, Any]:
        return {
            "document_id": self._remember_doc(doc),
            "name": str(doc.Name),
            "path": str(doc.FullName) if doc.Path else None,
            "saved": bool(doc.Saved),
            "read_only": bool(doc.ReadOnly),
            "track_changes": bool(doc.TrackRevisions),
            "protection_type": int(doc.ProtectionType),
        }

    def list_documents(self) -> dict[str, Any]:
        self._active_app()
        docs = []
        seen: set[str] = set()
        for app in self._apps:
            try:
                for doc in app.Documents:
                    info = self._info(doc)
                    if info["document_id"] not in seen:
                        seen.add(info["document_id"])
                        docs.append(info)
            except pywintypes.com_error:
                continue
        return {"documents": docs, "scope": "Word instances reachable by this server"}

    def create_document(self, visible: bool = True) -> dict[str, Any]:
        app = self._active_app(create=True)
        if visible:
            app.Visible = True
        doc = app.Documents.Add(Visible=visible)
        return self._info(doc)

    def open_document(
        self, path: str, read_only: bool = False, visible: bool = True
    ) -> dict[str, Any]:
        file = validate_local_path(path, {".docx", ".docm", ".doc"})
        if not file.is_file():
            raise ValueError("Document does not exist: " + str(file))
        for info in self.list_documents()["documents"]:
            if info["path"] and os.path.normcase(info["path"]) == os.path.normcase(str(file)):
                if read_only and not info["read_only"]:
                    raise ValueError(
                        "Document is already open for editing; cannot claim read-only mode."
                    )
                return info
        _assert_file_available(file)
        app = self._active_app(create=True)
        if visible:
            app.Visible = True
        security = app.AutomationSecurity
        try:
            # Block macros in the opened document, preserving trusted installed add-ins.
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            doc = app.Documents.Open(
                FileName=str(file),
                ReadOnly=read_only,
                AddToRecentFiles=False,
                Visible=visible,
                ConfirmConversions=False,
                NoEncodingDialog=True,
            )
        finally:
            app.AutomationSecurity = security
        return self._info(doc)

    def get_document_info(self, document_id: str) -> dict[str, Any]:
        doc = self._doc(document_id)
        info = self._info(doc)
        info["bookmarks"] = [str(b.Name) for b in doc.Bookmarks]
        info["inline_shape_count"] = int(doc.InlineShapes.Count)
        app = doc.Application
        info["selection"] = None
        if _same(app.ActiveDocument, doc):
            selection = app.Selection.Range
            info["selection"] = {
                "start": int(selection.Start),
                "end": int(selection.End),
                "text": str(selection.Text)[:2000],
                "text_truncated": len(str(selection.Text)) > 2000,
                "story_type": int(selection.StoryType),
            }
        return info

    def _is_equation(self, shape: Any) -> bool:
        try:
            return shape.Type == 1 and str(shape.OLEFormat.ProgID).lower() == PROGID.lower()
        except pywintypes.com_error:
            return False

    def _equation_info(
        self, document_id: str, shape: Any, equation_id: str | None = None
    ) -> dict[str, Any]:
        if equation_id is None:
            for old_id, (owner, old_shape) in self._equations.items():
                # Word may return a fresh COM wrapper each time InlineShapes is
                # enumerated. Compare live ranges, not just IDispatch pointers.
                same_range = False
                if owner == document_id and self._is_equation(old_shape):
                    try:
                        same_range = bool(shape.Range.IsEqual(old_shape.Range))
                    except pywintypes.com_error:
                        pass
                if owner == document_id and (_same(shape, old_shape) or same_range):
                    equation_id = old_id
                    break
        equation_id = equation_id or "eq_" + uuid.uuid4().hex
        self._equations[equation_id] = (document_id, shape)
        return {
            "equation_id": equation_id,
            "document_id": document_id,
            "progid": str(shape.OLEFormat.ProgID),
            "start": int(shape.Range.Start),
            "end": int(shape.Range.End),
            "width_pt": float(shape.Width),
            "height_pt": float(shape.Height),
            "font_size_pt": equation_font_size(shape.Range),
            **describe_layout(shape),
        }

    def list_equations(self, document_id: str) -> dict[str, Any]:
        doc = self._doc(document_id)
        equations = [
            self._equation_info(document_id, shape)
            for shape in doc.InlineShapes
            if self._is_equation(shape)
        ]
        return {
            "document_id": document_id,
            "equations": equations,
            "scope": "native AxMath objects in main text, including display/numbered paragraphs",
        }

    def _eq(self, document_id: str, equation_id: str) -> Any:
        self._doc(document_id)
        entry = self._equations.get(equation_id)
        if not entry or entry[0] != document_id or not self._is_equation(entry[1]):
            raise ValueError("Unknown/stale equation_id or wrong document. List equations again.")
        return entry[1]

    def _editable(self, doc: Any) -> None:
        if doc.ReadOnly or doc.ProtectionType != -1:
            raise ValueError("Document is read-only or protected.")
        if doc.TrackRevisions:
            raise ValueError("Track Changes is enabled. Turn it off before equation replacement.")

    def _ready(self, app: Any) -> None:
        if not _reg(winreg.HKEY_CLASSES_ROOT, PROGID + r"\CLSID"):
            raise RuntimeError("AxMath OLE is not registered. Install/repair AxMath first.")
        if _reg(winreg.HKEY_CURRENT_USER, r"Software\AxMath\AuthUser", "Licensed") == "0":
            raise RuntimeError("AxMath reports unlicensed. Activate it normally before conversion.")
        if _reg(winreg.HKEY_CURRENT_USER, r"Software\AxMath\WordCmds", "ConvertOccupy") == "1":
            raise RuntimeError("AxMath is busy. Finish the current AxMath operation first.")
        if not any(str(a.Name).lower() == "axmath.dotm" and a.Installed for a in app.AddIns):
            raise RuntimeError(
                "AxMath.dotm is not loaded in this Word instance. Repair the AxMath Word add-in."
            )

    @contextmanager
    def _conversion(self, app: Any = None):
        status = win32event.WaitForSingleObject(self._mutex, 0)
        if status not in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
            raise RuntimeError("Another AxMath MCP conversion is running. Wait before retrying.")
        original_doc = original_selection = None
        try:
            if app is None:
                if (
                    _reg(winreg.HKEY_CURRENT_USER, r"Software\AxMath\WordCmds", "ConvertOccupy")
                    == "1"
                ):
                    raise RuntimeError("AxMath is busy. Finish the current AxMath operation first.")
                app = self._converter_app()
            self._ready(app)
            if app.Documents.Count:
                original_doc = app.ActiveDocument
                original_selection = app.Selection.Range.Duplicate
            screen = app.ScreenUpdating
            options = app.Options
            smart = options.SmartCutPaste
            paste_smart = options.PasteSmartCutPaste
            replace = options.ReplaceSelection
            previous_conversion = getattr(self, "_conversion_app", None)
            self._conversion_app = app
            try:
                yield app
            finally:
                self._restore_state_warnings = []
                for obj, name, value in (
                    (app, "ScreenUpdating", screen),
                    (options, "SmartCutPaste", smart),
                    (options, "PasteSmartCutPaste", paste_smart),
                    (options, "ReplaceSelection", replace),
                ):
                    try:
                        setattr(obj, name, value)
                    except pywintypes.com_error as exc:
                        warning = f"Private Word state restore pending ({name}): {exc}"
                        self._restore_state_warnings.append(warning)
                        logger.warning(warning)
                if original_doc is not None:
                    try:
                        original_doc.Activate()
                        original_selection.Select()
                    except pywintypes.com_error:
                        pass
                self._conversion_app = previous_conversion
        finally:
            win32event.ReleaseMutex(self._mutex)

    @contextmanager
    def _scratch(self, app: Any):
        if not self._workspace.owns_app(app):
            raise RuntimeError("Scratch documents must use the private conversion Word instance.")
        with self._workspace.scratch() as doc:
            yield doc

    def _run_macro(self, app: Any, macro: str) -> None:
        owned = (
            getattr(self, "_conversion_app", None) is app
            and hasattr(self, "_workspace")
            and self._workspace.owns_app(app)
        )
        initial_busy = (
            _reg(winreg.HKEY_CURRENT_USER, r"Software\AxMath\WordCmds", "ConvertOccupy")
            if owned
            else None
        )
        if owned and initial_busy == "1":
            raise RuntimeError("AxMath became busy before this macro started; no state was reset.")
        error = None
        try:
            result = app.Run(macro, False, False)
        except pywintypes.com_error as exc:
            error = exc
            raise RuntimeError(
                "AxMath macro failed. Check Word/AxMath dialogs and the installed add-in; "
                "no automatic retry is performed. " + str(exc)
            ) from exc
        finally:
            if owned:
                self._recover_completed_macro(app, initial_busy, error)
        if not result:
            raise RuntimeError(
                "AxMath did not convert the selected formula; check its LaTeX support."
            )
        if _reg(winreg.HKEY_CURRENT_USER, r"Software\AxMath\WordCmds", "WaitingConvert") != "1":
            raise RuntimeError(
                "AxMath did not acknowledge completion. The vendor macro may have timed out; "
                "target document was not changed. Inspect Word/AxMath before another request."
            )

    def _recover_completed_macro(self, app: Any, initial_busy: str | None, error: Any) -> None:
        """Release only our acknowledged, finished macro, never an existing busy state."""
        if initial_busy != "0":
            return
        # These transport failures do not prove that the native operation ended.
        uncertain = {-2147418111, -2147417846, -2147417848, -2147220995}
        if error is not None:
            codes = [error.hresult]
            if error.excepinfo and len(error.excepinfo) > 5:
                codes.append(error.excepinfo[5])
            # Word can wrap an RPC failure in DISP_E_EXCEPTION, and COM callers
            # may expose either signed or unsigned 32-bit HRESULT values.
            normalized = {
                code if code < 2**31 else code - 2**32 for code in codes if isinstance(code, int)
            }
            if normalized & uncertain:
                return
        key = r"Software\AxMath\WordCmds"
        if _reg(winreg.HKEY_CURRENT_USER, key, "ConvertOccupy") != "1":
            return
        if any(
            _reg(winreg.HKEY_CURRENT_USER, key, n) != "0" for n in ("AxMathOccupy", "DoVerbOccupy")
        ):
            return
        if _reg(winreg.HKEY_CURRENT_USER, key, "WaitingConvert") != "1":
            return
        try:
            _ = app.Documents.Count  # The private Word server must still respond.
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE) as handle:
                winreg.SetValueEx(handle, "ConvertOccupy", 0, winreg.REG_SZ, "0")
            self._busy_recoveries += 1
            self._recovery_warnings = []
        except (OSError, pywintypes.com_error) as exc:
            self._recovery_warnings = [
                f"Could not release completed private conversion state: {exc}"
            ]

    def _export_copy(self, app: Any, shape: Any) -> str:
        with self._scratch(app) as copy:
            copy_native_range(copy.Range(0, 0), shape.Range)
            return self._export_inplace(app, copy)

    def _export_inplace(self, app: Any, doc: Any) -> str:
        if not self._workspace.owns_document(doc):
            raise RuntimeError("Destructive LaTeX export is restricted to owned scratch documents.")
        doc.Activate()
        doc.Content.Select()
        self._run_macro(app, TO_LATEX)
        source = self._validated_export_text(doc.Content.Text)
        if doc.InlineShapes.Count:
            raise RuntimeError("AxMath export left an unconverted object in the scratch copy.")
        return source

    @staticmethod
    def _validated_export_text(text: str) -> str:
        source = str(text).replace("\r\n", "\n").replace("\r", "\n").strip()
        if not source or "\x01" in source:
            raise RuntimeError("AxMath round-trip export failed; target document was not changed.")
        normalize_latex(source)
        return source

    def _verify_generated(self, app: Any, scratch: Any, shape: Any) -> str:
        # Verify a duplicate in a second paragraph of the SAME scratch document.
        # Only the duplicate is selected; the original native object stays intact.
        start = int(scratch.Content.End - 1)
        duplicate = scratch.Range(start, start)
        duplicate.FormattedText = shape.Range.FormattedText
        duplicate.Select()
        self._run_macro(app, TO_LATEX)
        tail = scratch.Range(start, scratch.Content.End - 1)
        source = self._validated_export_text(tail.Text)
        tail.Delete()
        if scratch.InlineShapes.Count != 1 or not self._is_equation(shape):
            raise RuntimeError(
                "AxMath round-trip modified the original scratch formula unexpectedly."
            )
        return source

    @contextmanager
    def _generated(self, app: Any, source: str, font_size: float = 12):
        check_axmath_input(source)
        with self._scratch(app) as scratch:
            scratch.Content.Font.Size = font_size if 1 <= font_size <= 1638 else 12
            scratch.Content.Text = source + "\r"
            scratch.Range(0, scratch.Content.End - 1).Select()
            self._run_macro(app, TO_AXMATH)
            if scratch.InlineShapes.Count != 1 or not self._is_equation(scratch.InlineShapes(1)):
                raise RuntimeError(
                    "Expected exactly one native AxMath OLE object; target was not changed."
                )
            shape = scratch.InlineShapes(1)
            # Record the requested host font size even when the vendor's hidden
            # EMBED field uses mixed sizes. Never rewrite native Font.Position.
            shape.Range.Font.Size = font_size
            if shape.Width <= 0 or shape.Height <= 0:
                raise RuntimeError("AxMath produced an empty display; target was not changed.")
            actual_latex = self._verify_generated(app, scratch, shape)
            yield shape, actual_latex

    def _insertion_range(
        self, doc: Any, position: str, bookmark: str | None, expected_selection: str | None
    ) -> Any:
        if position == "end":
            if bookmark is not None or expected_selection is not None:
                raise ValueError(
                    "bookmark/expected_selection are only valid for their respective positions."
                )
            return doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        if position == "bookmark":
            if expected_selection is not None or not bookmark or not doc.Bookmarks.Exists(bookmark):
                raise ValueError("Choose an existing bookmark and omit expected_selection.")
            rng = doc.Bookmarks(bookmark).Range.Duplicate
            if rng.StoryType != 1:
                raise ValueError("Only bookmarks in the main document text are supported.")
            rng.Collapse(1)  # Insert before bookmark text, do not replace it.
            return rng
        if position == "selection":
            if bookmark is not None or not _same(doc.Application.ActiveDocument, doc):
                raise ValueError(
                    "Target document must be active for selection insertion; omit bookmark."
                )
            rng = doc.Application.Selection.Range.Duplicate
            if rng.StoryType != 1:
                raise ValueError("Only main-text selection is supported.")
            if expected_selection is None or str(rng.Text) != expected_selection:
                raise ValueError(
                    "Pass expected_selection exactly as inspected (empty string for a cursor)."
                )
            if any(c in str(rng.Text) for c in ("\r", "\x07", "\x01")) or rng.Fields.Count:
                raise ValueError(
                    "Selection must be plain text within one paragraph, without fields/objects."
                )
            return rng
        raise ValueError("position must be end, bookmark, or selection.")

    def _replace(self, doc: Any, target: Any, generated: Any) -> Any:
        doc.Activate()
        start = int(target.Start)
        doc.Application.UndoRecord.StartCustomRecord("AxMath MCP equation")
        try:
            copy_native_range(target, generated.Range)
        finally:
            doc.Application.UndoRecord.EndCustomRecord()
        # Word's OLE range can include hidden field characters (28 positions on
        # the tested installation), even though Range.Text shows one U+0001.
        # FormattedText updates the destination Range to cover the whole object.
        inserted = target
        if (
            inserted.Start != start
            or inserted.InlineShapes.Count != 1
            or not self._is_equation(inserted.InlineShapes(1))
        ):
            raise RuntimeError(
                "Insertion outcome needs inspection. Do not retry blindly; use Word Undo if needed."
            )
        return inserted.InlineShapes(1)

    def insert_equation(
        self,
        document_id: str,
        latex: str,
        position: str = "end",
        bookmark: str | None = None,
        expected_selection: str | None = None,
        layout: str = "inline",
        font_size_pt: float | None = None,
        number_label: str | None = None,
    ) -> dict[str, Any]:
        options = validate_layout_options(layout, font_size_pt, number_label)
        source = latex_for_layout(latex, layout)
        doc = self._doc(document_id)
        self._editable(doc)
        target = self._insertion_range(doc, position, bookmark, expected_selection)
        original_text = str(target.Text)
        size = options.font_size_pt or self._font_size(target)
        with (
            self._conversion() as converter,
            self._generated(converter, source, size) as (
                shape,
                actual,
            ),
        ):
            self._editable(doc)
            if str(target.Text) != original_text:
                raise RuntimeError(
                    "Target changed during conversion. Inspect the selection and retry."
                )
            doc.Application.UndoRecord.StartCustomRecord("AxMath MCP layout")
            try:
                inserted = place_equation(doc, target, shape, options, size, self._replace)
            finally:
                doc.Application.UndoRecord.EndCustomRecord()
        result = self._equation_info(document_id, inserted)
        result.update(
            {
                "latex_exported": actual,
                "saved": False,
                "note": "Native AxMath OLE verified; visually review mathematical meaning in Word.",
            }
        )
        return result

    def update_equation(self, document_id: str, equation_id: str, latex: str) -> dict[str, Any]:
        doc = self._doc(document_id)
        self._editable(doc)
        original = self._eq(document_id, equation_id)
        state = describe_layout(original)
        options = validate_layout_options(state["layout"], None, state["number_label"])
        source = latex_for_layout(latex, options.layout)
        size = self._font_size(original.Range)
        with (
            self._conversion() as converter,
            self._generated(converter, source, size) as (
                shape,
                actual,
            ),
        ):
            self._editable(doc)
            original = self._eq(document_id, equation_id)
            if options.layout != "inline":
                check_block_fits(original.Range, shape, options, size)
            inserted = self._replace(doc, original.Range.Duplicate, shape)
            if options.layout == "inline":
                inline_paragraph(inserted)
        result = self._equation_info(document_id, inserted, equation_id)
        result.update({"latex_exported": actual, "saved": False})
        return result

    @staticmethod
    def _font_size(rng: Any) -> float:
        return equation_font_size(rng)

    def format_equation(
        self,
        document_id: str,
        equation_id: str,
        layout: str = "inline",
        font_size_pt: float | None = None,
        number_label: str | None = None,
    ) -> dict[str, Any]:
        options = validate_layout_options(layout, font_size_pt, number_label)
        doc = self._doc(document_id)
        self._editable(doc)
        original = self._eq(document_id, equation_id)
        target = managed_target(doc, original)
        preflight_numbers(doc)
        expected = str(target.Text)
        size = options.font_size_pt or self._font_size(original.Range)
        with self._conversion() as converter:
            source = self._export_copy(converter, original)
            with self._generated(converter, latex_for_layout(source, layout), size) as (
                shape,
                actual,
            ):
                self._editable(doc)
                if str(target.Text) != expected:
                    raise RuntimeError(
                        "Formula paragraph changed during conversion; inspect it before retrying."
                    )
                doc.Application.UndoRecord.StartCustomRecord("AxMath MCP reformat")
                try:
                    inserted = place_equation(doc, target, shape, options, size, self._replace)
                    refresh_numbers(doc)
                finally:
                    doc.Application.UndoRecord.EndCustomRecord()
        result = self._equation_info(document_id, inserted, equation_id)
        result.update({"latex_exported": actual, "saved": False})
        return result

    def refresh_equation_numbers(self, document_id: str) -> dict[str, Any]:
        doc = self._doc(document_id)
        self._editable(doc)
        result = refresh_numbers(doc)
        result.update({"document_id": document_id, "saved": False})
        return result

    def get_equation_latex(self, document_id: str, equation_id: str) -> dict[str, Any]:
        shape = self._eq(document_id, equation_id)
        with self._conversion() as converter:
            source = self._export_copy(converter, shape)
        return {
            "document_id": document_id,
            "equation_id": equation_id,
            "latex": source,
            "source": "exported from a disposable copy of the live AxMath object",
        }

    def open_equation_editor(self, document_id: str, equation_id: str) -> dict[str, Any]:
        doc = self._doc(document_id)
        self._editable(doc)
        shape = self._eq(document_id, equation_id)
        doc.Application.Visible = True
        doc.Activate()
        shape.Range.Select()
        shape.OLEFormat.DoVerb(0)
        return {
            "document_id": document_id,
            "equation_id": equation_id,
            "editor_requested": True,
            "note": "Finish and close AxMath editor before further calls.",
        }

    def save_document(
        self, document_id: str, path: str | None = None, overwrite: bool = False
    ) -> dict[str, Any]:
        doc = self._doc(document_id)
        if doc.ReadOnly:
            raise ValueError("Read-only documents cannot be saved by this tool.")
        if path is None:
            if not doc.Path:
                raise ValueError("An absolute .docx path is required for an unsaved document.")
            doc.Save()
        else:
            target = validate_local_path(path, {".docx"})
            if not target.parent.is_dir():
                raise ValueError("Destination directory does not exist.")
            current = str(doc.FullName) if doc.Path else None
            same_path = current and os.path.normcase(current) == os.path.normcase(str(target))
            if target.exists() and not same_path and not overwrite:
                raise ValueError(
                    "Destination exists; choose a new path or explicitly set overwrite=True."
                )
            if not same_path:
                for info in self.list_documents()["documents"]:
                    if info["path"] and os.path.normcase(info["path"]) == os.path.normcase(
                        str(target)
                    ):
                        raise ValueError(
                            "Destination is open in another document; choose a different path."
                        )
            if doc.HasVBProject or str(doc.Name).lower().endswith(".docm"):
                raise ValueError(
                    "Use path=None to preserve this document's VBA project or macro-enabled "
                    "format; Save As .docx would remove them."
                )
            alerts = doc.Application.DisplayAlerts
            try:
                doc.Application.DisplayAlerts = 0
                doc.SaveAs2(FileName=str(target), FileFormat=16, AddToRecentFiles=False)
            finally:
                doc.Application.DisplayAlerts = alerts
        return self._info(doc)

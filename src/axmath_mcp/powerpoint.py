"""Editable PowerPoint equations using the installed Word/AxMath converter.

Runs only on the shared Office STA. Formula contents always come from native OLE
objects; Tags hold numeric layout metadata only, never cached LaTeX.
"""

from __future__ import annotations

import os
import tempfile
import uuid
import winreg
from pathlib import Path
from typing import Any

import pywintypes
import win32com.client

from axmath_mcp.backend import PROGID, WordBackend, _assert_file_available, _reg, _same
from axmath_mcp.ole_bridge import extract_axmath_storage
from axmath_mcp.ppt_validation import (
    ensure_fits,
    validate_font_size,
    validate_number_label,
    validate_position,
)
from axmath_mcp.validation import latex_for_layout, validate_local_path

FONT_TAG = "AXMATH_MCP_FONT"
LABEL_TAG = "AXMATH_MCP_LABEL"
OWNER_TAG = "AXMATH_MCP_OWNER"
WIDTH_TAG = "AXMATH_MCP_WIDTH"
HEIGHT_TAG = "AXMATH_MCP_HEIGHT"
LABEL_GAP = 12.0


def _items(collection: Any) -> list[Any]:
    return [collection.Item(i) for i in range(1, int(collection.Count) + 1)]


def _tag(shape: Any, name: str) -> str:
    return str(shape.Tags.Item(name))


class PowerPointBackend:
    def __init__(self, word_backend: WordBackend) -> None:
        self.word = word_backend
        self._apps: list[Any] = []
        self._presentations: dict[str, Any] = {}
        self._equations: dict[str, tuple[str, int, int]] = {}

    def shutdown(self) -> None:
        # Never close the user's presentations or PowerPoint application.
        self._equations.clear()
        self._presentations.clear()
        self._apps.clear()

    def health(self) -> dict[str, Any]:
        return {
            "powerpoint_registered": bool(
                _reg(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID")
            ),
            "conversion_backend": "Installed Word + AxMath.dotm; PowerPoint embeds native AxMath OLE",
            "requires_word": True,
            "requires_powerpoint_addin": False,
            "supported": "Top-level slide OLE objects; position in points; optional fixed text number",
            "limitations": "No inline TextFrame equations, groups, masters or automatic SEQ numbering",
        }

    def _remember_app(self, app: Any) -> Any:
        if not any(_same(app, old) for old in self._apps):
            self._apps.append(app)
        return app

    def _app(self, create: bool = False) -> Any | None:
        try:
            return self._remember_app(win32com.client.GetActiveObject("PowerPoint.Application"))
        except pywintypes.com_error:
            for app in self._apps:
                try:
                    _ = app.Presentations.Count
                    return app
                except pywintypes.com_error:
                    continue
            if create:
                return self._remember_app(win32com.client.DispatchEx("PowerPoint.Application"))
            return None

    def _remember(self, presentation: Any) -> str:
        for pid, old in self._presentations.items():
            if _same(presentation, old):
                return pid
        pid = "pres_" + uuid.uuid4().hex
        self._presentations[pid] = presentation
        return pid

    def _pres(self, presentation_id: str) -> Any:
        pres = self._presentations.get(presentation_id)
        if pres is None:
            raise ValueError("Unknown presentation_id. List, create or open a presentation first.")
        try:
            _ = pres.Name
        except pywintypes.com_error as exc:
            raise ValueError("Presentation was closed; list or open it again.") from exc
        return pres

    def _info(self, pres: Any) -> dict[str, Any]:
        return {
            "presentation_id": self._remember(pres),
            "name": str(pres.Name),
            "path": str(pres.FullName) if pres.Path else None,
            "read_only": bool(pres.ReadOnly),
            "saved": bool(pres.Saved),
            "slide_count": int(pres.Slides.Count),
            "slide_width_pt": float(pres.PageSetup.SlideWidth),
            "slide_height_pt": float(pres.PageSetup.SlideHeight),
        }

    def ppt_list_presentations(self) -> dict[str, Any]:
        self._app()
        results = {}
        for app in self._apps:
            try:
                for pres in _items(app.Presentations):
                    info = self._info(pres)
                    results[info["presentation_id"]] = info
            except pywintypes.com_error:
                continue
        return {
            "presentations": list(results.values()),
            "scope": "PowerPoint instances reachable by this server",
        }

    def ppt_create_presentation(self, visible: bool = True) -> dict[str, Any]:
        app = self._app(create=True)
        if visible:
            app.Visible = True
        pres = app.Presentations.Add(WithWindow=visible)
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540
        return self._info(pres)

    def ppt_open_presentation(
        self, path: str, read_only: bool = False, visible: bool = True
    ) -> dict[str, Any]:
        file = validate_local_path(path, {".pptx", ".pptm", ".ppt"})
        if not file.is_file():
            raise ValueError("Presentation does not exist: " + str(file))
        for info in self.ppt_list_presentations()["presentations"]:
            if info["path"] and os.path.normcase(info["path"]) == os.path.normcase(str(file)):
                if read_only and not info["read_only"]:
                    raise ValueError("Presentation is already open for editing, not read-only.")
                return info
        _assert_file_available(file)
        app = self._app(create=True)
        if visible:
            app.Visible = True
        security = app.AutomationSecurity
        try:
            app.AutomationSecurity = 3
            pres = app.Presentations.Open(
                FileName=str(file), ReadOnly=read_only, WithWindow=visible
            )
        finally:
            app.AutomationSecurity = security
        return self._info(pres)

    def _slide(self, presentation_id: str, slide_id: int) -> tuple[Any, Any]:
        pres = self._pres(presentation_id)
        if isinstance(slide_id, bool) or not isinstance(slide_id, int) or slide_id <= 0:
            raise ValueError("slide_id must be the positive SlideID returned by ppt_list_slides.")
        try:
            return pres, pres.Slides.FindBySlideID(slide_id)
        except pywintypes.com_error as exc:
            raise ValueError(
                "SlideID not found in this presentation; do not use a slide index."
            ) from exc

    @staticmethod
    def _editable(pres: Any) -> None:
        if pres.ReadOnly:
            raise ValueError("Presentation is read-only.")

    def ppt_list_slides(self, presentation_id: str) -> dict[str, Any]:
        pres = self._pres(presentation_id)
        slides = []
        for slide in _items(pres.Slides):
            title = (
                str(slide.Shapes.Title.TextFrame.TextRange.Text) if slide.Shapes.HasTitle else ""
            )
            slides.append(
                {
                    "slide_id": int(slide.SlideID),
                    "index": int(slide.SlideIndex),
                    "name": str(slide.Name),
                    "title": title,
                    "shape_count": int(slide.Shapes.Count),
                }
            )
        return {
            "presentation_id": presentation_id,
            "slides": slides,
            "slide_width_pt": float(pres.PageSetup.SlideWidth),
            "slide_height_pt": float(pres.PageSetup.SlideHeight),
        }

    def ppt_add_slide(self, presentation_id: str, title: str = "") -> dict[str, Any]:
        pres = self._pres(presentation_id)
        self._editable(pres)
        if len(title) > 500 or any(ord(c) < 32 and c not in "\n\t" for c in title):
            raise ValueError("Title must be at most 500 characters without control codes.")
        slide = pres.Slides.Add(pres.Slides.Count + 1, 11 if title else 12)  # title-only or blank
        if title:
            slide.Shapes.Title.TextFrame.TextRange.Text = title
        return {
            "presentation_id": presentation_id,
            "slide_id": int(slide.SlideID),
            "index": int(slide.SlideIndex),
            "title": title,
        }

    @staticmethod
    def _is_equation(shape: Any) -> bool:
        try:
            return int(shape.Type) == 7 and str(shape.OLEFormat.ProgID).lower() == PROGID.lower()
        except pywintypes.com_error:
            return False

    @staticmethod
    def _by_shape_id(slide: Any, shape_id: int) -> Any | None:
        return next((s for s in _items(slide.Shapes) if int(s.Id) == shape_id), None)

    def _label(self, slide: Any, shape: Any) -> Any | None:
        label_id = _tag(shape, LABEL_TAG)
        if not label_id.isdecimal():
            return None
        label = self._by_shape_id(slide, int(label_id))
        if label is not None and _tag(label, OWNER_TAG) == str(shape.Id) and label.HasTextFrame:
            return label
        return None

    def _equation_info(
        self, pid: str, slide: Any, shape: Any, equation_id: str | None = None
    ) -> dict[str, Any]:
        identity = (pid, int(slide.SlideID), int(shape.Id))
        if equation_id is None:
            equation_id = next(
                (eid for eid, ref in self._equations.items() if ref == identity), None
            )
        equation_id = equation_id or "ppt_eq_" + uuid.uuid4().hex
        self._equations[equation_id] = identity
        raw_size = _tag(shape, FONT_TAG)
        try:
            font_size = validate_font_size(float(raw_size))
        except (ValueError, TypeError):
            font_size = None
        label = self._label(slide, shape)
        return {
            "presentation_id": pid,
            "slide_id": int(slide.SlideID),
            "equation_id": equation_id,
            "shape_id": int(shape.Id),
            "name": str(shape.Name),
            "progid": str(shape.OLEFormat.ProgID),
            "left_pt": float(shape.Left),
            "top_pt": float(shape.Top),
            "width_pt": float(shape.Width),
            "height_pt": float(shape.Height),
            "font_size_pt": font_size,
            "number_label": str(label.TextFrame.TextRange.Text) if label is not None else None,
        }

    def ppt_list_equations(
        self, presentation_id: str, slide_id: int | None = None
    ) -> dict[str, Any]:
        pres = self._pres(presentation_id)
        slides = (
            [self._slide(presentation_id, slide_id)[1]]
            if slide_id is not None
            else _items(pres.Slides)
        )
        return {
            "presentation_id": presentation_id,
            "equations": [
                self._equation_info(presentation_id, slide, shape)
                for slide in slides
                for shape in _items(slide.Shapes)
                if self._is_equation(shape)
            ],
            "scope": "top-level native AxMath OLE shapes; no grouped/master objects",
        }

    def _eq(self, pid: str, slide_id: int, equation_id: str) -> tuple[Any, Any, Any]:
        pres, slide = self._slide(pid, slide_id)
        ref = self._equations.get(equation_id)
        if ref is None or ref[:2] != (pid, slide_id):
            raise ValueError(
                "Unknown equation_id or wrong presentation/slide. List equations again."
            )
        shape = self._by_shape_id(slide, ref[2])
        if shape is None or not self._is_equation(shape):
            raise ValueError("Equation was removed/replaced. List equations again.")
        return pres, slide, shape

    @staticmethod
    def _import_native(slide: Any, word_shape: Any, left: float, top: float) -> Any:
        data = extract_axmath_storage(str(word_shape.Range.WordOpenXML))
        with tempfile.TemporaryDirectory(prefix="axmath-mcp-") as directory:
            path = Path(directory) / "equation.afx"
            path.write_bytes(data)
            shape = slide.Shapes.AddOLEObject(
                Left=left, Top=top, FileName=str(path), Link=0, DisplayAsIcon=0
            )
        if not PowerPointBackend._is_equation(shape):
            shape.Delete()
            raise RuntimeError("PowerPoint did not create a native Equation.AxMath object.")
        return shape

    @staticmethod
    def _new_label(slide: Any, text: str, size: float) -> Any:
        label = slide.Shapes.AddTextbox(1, 0, 0, 100, size * 1.5)
        try:
            frame = label.TextFrame
            frame.MarginLeft = frame.MarginRight = frame.MarginTop = frame.MarginBottom = 0
            frame.WordWrap = 0
            frame.TextRange.Text = text
            frame.TextRange.Font.Name = "Times New Roman"
            frame.TextRange.Font.Size = size
            frame.TextRange.ParagraphFormat.Alignment = 2
            label.Width = max(24.0, float(frame.TextRange.BoundWidth) + 4)
            label.Height = max(size * 1.4, float(frame.TextRange.BoundHeight) + 2)
            return label
        except Exception:
            label.Delete()
            raise

    @staticmethod
    def _check_bounds(
        pres: Any, shape: Any, label: Any | None, left: float, top: float
    ) -> tuple[float, float] | None:
        width, height = float(shape.Width), float(shape.Height)
        sw, sh = float(pres.PageSetup.SlideWidth), float(pres.PageSetup.SlideHeight)
        ensure_fits(
            left,
            top,
            width,
            height,
            sw,
            sh,
            number_width=float(label.Width) if label is not None else 0,
            gap=LABEL_GAP,
        )
        if label is not None:
            lx, ly = left + width + LABEL_GAP, top + (height - float(label.Height)) / 2
            ensure_fits(lx, ly, float(label.Width), float(label.Height), sw, sh)
            return lx, ly
        return None

    @staticmethod
    def _attach(shape: Any, label: Any | None, size: float) -> None:
        # PowerPoint uppercases Tags values, so only numeric metadata is stored.
        shape.Tags.Add(FONT_TAG, str(size))
        shape.Tags.Add(WIDTH_TAG, str(float(shape.Width)))
        shape.Tags.Add(HEIGHT_TAG, str(float(shape.Height)))
        if label is not None:
            shape.Tags.Add(LABEL_TAG, str(label.Id))
            label.Tags.Add(OWNER_TAG, str(shape.Id))

    def _generate(
        self,
        pres: Any,
        slide: Any,
        latex: str,
        left: float,
        top: float,
        size: float,
        label_text: str | None,
    ) -> tuple[Any, Any | None, str]:
        app = self.word._active_app(create=True)
        shape = label = None
        try:
            with (
                self.word._conversion(app),
                self.word._generated(app, latex_for_layout(latex, "display"), size) as (
                    native,
                    actual,
                ),
            ):
                self._editable(pres)
                shape = self._import_native(slide, native, left, top)
                if label_text is not None:
                    label = self._new_label(slide, label_text, size)
                label_xy = self._check_bounds(pres, shape, label, left, top)
                if label_xy:
                    label.Left, label.Top = label_xy
                self._attach(shape, label, size)
                return shape, label, actual
        except Exception:
            if label is not None:
                label.Delete()
            if shape is not None:
                shape.Delete()
            raise

    def ppt_insert_equation(
        self,
        presentation_id: str,
        slide_id: int,
        latex: str,
        left_pt: float = 72,
        top_pt: float = 144,
        font_size_pt: float = 24,
        number_label: str | None = None,
    ) -> dict[str, Any]:
        source = latex_for_layout(latex, "display")
        left, top = validate_position(left_pt, top_pt)
        size = validate_font_size(font_size_pt)
        label_text = validate_number_label(number_label)
        pres, slide = self._slide(presentation_id, slide_id)
        self._editable(pres)
        shape, _, actual = self._generate(pres, slide, source, left, top, size, label_text)
        result = self._equation_info(presentation_id, slide, shape)
        result.update({"latex_exported": actual, "saved": False})
        return result

    @staticmethod
    def _assert_replaceable(slide: Any, shapes: list[Any]) -> None:
        identities = {int(s.Id) for s in shapes}
        for shape in shapes:
            if (
                abs(float(shape.Rotation)) > 0.001
                or int(shape.HorizontalFlip) != 0
                or int(shape.VerticalFlip) != 0
            ):
                raise ValueError(
                    "Replacing rotated or flipped equations/numbers is not supported. "
                    "Restore their original orientation before replacing them."
                )
        for candidate in _items(slide.Shapes):
            if int(candidate.Connector) == 0:
                continue
            connector = candidate.ConnectorFormat
            if (
                connector.BeginConnected and int(connector.BeginConnectedShape.Id) in identities
            ) or (connector.EndConnected and int(connector.EndConnectedShape.Id) in identities):
                raise ValueError(
                    "Replacing equations/numbers with attached connectors is not supported. "
                    "Detach the connectors before replacing them."
                )
        sequences = [slide.TimeLine.MainSequence, *_items(slide.TimeLine.InteractiveSequences)]
        for sequence in sequences:
            for effect in _items(sequence):
                if int(effect.Shape.Id) in identities:
                    raise ValueError(
                        "This equation/number has animations; remove them before replacing it."
                    )
        for shape in shapes:
            if any(int(shape.ActionSettings.Item(i).Action) != 0 for i in (1, 2)):
                raise ValueError(
                    "This equation/number has click or mouse actions; remove them before replacement."
                )

    def ppt_update_equation(
        self,
        presentation_id: str,
        slide_id: int,
        equation_id: str,
        latex: str,
        font_size_pt: float | None = None,
    ) -> dict[str, Any]:
        source = latex_for_layout(latex, "display")
        pres, slide, old = self._eq(presentation_id, slide_id, equation_id)
        self._editable(pres)
        info = self._equation_info(presentation_id, slide, old, equation_id)
        if font_size_pt is None and info["font_size_pt"] is None:
            raise ValueError(
                "This equation has no recorded font size. Specify font_size_pt when replacing it."
            )
        size = validate_font_size(
            font_size_pt if font_size_pt is not None else info["font_size_pt"]
        )
        if font_size_pt is None:
            for tag, dimension in ((WIDTH_TAG, old.Width), (HEIGHT_TAG, old.Height)):
                stored = _tag(old, tag)
                if stored and abs(float(stored) - float(dimension)) > 0.5:
                    raise ValueError(
                        "Equation was resized manually; specify font_size_pt to regenerate it."
                    )
        old_label = self._label(slide, old)
        self._assert_replaceable(slide, [old] + ([old_label] if old_label is not None else []))
        left, top, z, name = float(old.Left), float(old.Top), int(old.ZOrderPosition), str(old.Name)
        shape, label, actual = self._generate(
            pres, slide, source, left, top, size, info["number_label"]
        )
        try:
            shape.Name = name
            while shape.ZOrderPosition > z:
                shape.ZOrder(3)  # msoSendBackward
        except Exception:
            if label is not None:
                label.Delete()
            shape.Delete()
            raise
        try:
            old.Delete()  # Original remains intact until candidate and number are verified.
        except pywintypes.com_error as exc:
            # A cross-process error may arrive after Delete has executed. Keep the
            # validated candidate as a recovery copy rather than risk losing both.
            recovery = self._equation_info(presentation_id, slide, shape)
            raise RuntimeError(
                "Could not confirm removal of the old equation. A validated replacement "
                f"was retained as {recovery['equation_id']}; inspect the slide before retrying."
            ) from exc
        result = self._equation_info(presentation_id, slide, shape, equation_id)
        result.update({"latex_exported": actual, "saved": False})
        if old_label is not None:
            try:
                old_label.Delete()
            except pywintypes.com_error:
                result["warnings"] = [
                    (
                        "Formula replaced, but the old number label could not be removed. "
                        "Inspect the slide and remove the redundant label; do not repeat the update."
                    )
                ]
        return result

    def ppt_position_equation(
        self, presentation_id: str, slide_id: int, equation_id: str, left_pt: float, top_pt: float
    ) -> dict[str, Any]:
        left, top = validate_position(left_pt, top_pt)
        pres, slide, shape = self._eq(presentation_id, slide_id, equation_id)
        self._editable(pres)
        label = self._label(slide, shape)
        xy = self._check_bounds(pres, shape, label, left, top)
        old = (
            shape.Left,
            shape.Top,
            label.Left if label is not None else None,
            label.Top if label is not None else None,
        )
        try:
            shape.Left, shape.Top = left, top
            if xy:
                label.Left, label.Top = xy
        except Exception:
            shape.Left, shape.Top = old[:2]
            if label is not None:
                label.Left, label.Top = old[2:]
            raise
        result = self._equation_info(presentation_id, slide, shape, equation_id)
        result["saved"] = False
        return result

    def ppt_get_equation_latex(
        self, presentation_id: str, slide_id: int, equation_id: str
    ) -> dict[str, Any]:
        _, _, shape = self._eq(presentation_id, slide_id, equation_id)
        app = self.word._active_app(create=True)
        with self.word._conversion(app), self.word._scratch(app) as copy:
            shape.Copy()
            copy.Range(0, 0).PasteSpecial(DataType=0, Placement=0, DisplayAsIcon=False, Link=False)
            # PowerPoint's clipboard yields a floating Word OLE despite Placement=0.
            if copy.InlineShapes.Count == 0 and copy.Shapes.Count == 1:
                floating = copy.Shapes(1)
                if self._is_equation(floating):
                    floating.ConvertToInlineShape()
            if copy.InlineShapes.Count != 1 or not self.word._is_equation(copy.InlineShapes(1)):
                raise RuntimeError(
                    "Could not read the native PowerPoint equation into the temporary Word copy."
                )
            source = self.word._export_copy(app, copy.InlineShapes(1))
        return {
            "presentation_id": presentation_id,
            "slide_id": slide_id,
            "equation_id": equation_id,
            "latex": source,
            "source": "exported from a native OLE copy; original presentation unchanged",
        }

    def ppt_open_equation_editor(
        self, presentation_id: str, slide_id: int, equation_id: str
    ) -> dict[str, Any]:
        pres, slide, shape = self._eq(presentation_id, slide_id, equation_id)
        self._editable(pres)
        pres.Application.Visible = True
        if not pres.Windows.Count:
            pres.NewWindow()
        window = pres.Windows.Item(1)
        window.Activate()
        window.View.GotoSlide(slide.SlideIndex)
        shape.OLEFormat.Activate()
        return {
            "presentation_id": presentation_id,
            "slide_id": slide_id,
            "equation_id": equation_id,
            "editor_requested": True,
            "note": "Finish and close AxMath editor before further Office calls.",
        }

    def ppt_save_presentation(
        self, presentation_id: str, path: str | None = None, overwrite: bool = False
    ) -> dict[str, Any]:
        pres = self._pres(presentation_id)
        self._editable(pres)
        if path is None:
            if not pres.Path:
                raise ValueError("An absolute .pptx path is required for an unsaved presentation.")
            pres.Save()
            return self._info(pres)
        target = validate_local_path(path, {".pptx"})
        if not target.parent.is_dir():
            raise ValueError("Destination directory does not exist.")
        current = str(pres.FullName) if pres.Path else None
        same = current and os.path.normcase(current) == os.path.normcase(str(target))
        if target.exists() and not same and not overwrite:
            raise ValueError(
                "Destination exists; choose another path or explicitly authorize overwrite=True."
            )
        if not same:
            for info in self.ppt_list_presentations()["presentations"]:
                if info["path"] and os.path.normcase(info["path"]) == os.path.normcase(str(target)):
                    raise ValueError("Destination is already open in another presentation.")
        if pres.HasVBProject or str(pres.Name).lower().endswith(".pptm"):
            raise ValueError(
                "Use path=None to preserve VBA; Save As .pptx would remove the macros."
            )
        alerts = pres.Application.DisplayAlerts
        try:
            pres.Application.DisplayAlerts = 1  # ppAlertsNone
            pres.SaveAs(FileName=str(target), FileFormat=24)
        finally:
            pres.Application.DisplayAlerts = alerts
        return self._info(pres)

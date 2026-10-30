"""Route Word and PowerPoint through one shared, serialized Office backend."""

from __future__ import annotations

from typing import Any

WORD_METHODS = frozenset(
    {
        "list_documents",
        "create_document",
        "open_document",
        "get_document_info",
        "list_equations",
        "insert_equation",
        "update_equation",
        "format_equation",
        "refresh_equation_numbers",
        "get_equation_latex",
        "open_equation_editor",
        "save_document",
    }
)
PPT_METHODS = frozenset(
    {
        "ppt_list_presentations",
        "ppt_create_presentation",
        "ppt_open_presentation",
        "ppt_list_slides",
        "ppt_add_slide",
        "ppt_list_equations",
        "ppt_insert_equation",
        "ppt_update_equation",
        "ppt_position_equation",
        "ppt_get_equation_latex",
        "ppt_open_equation_editor",
        "ppt_save_presentation",
    }
)


class OfficeBackend:
    """Own both backends on the STA, sharing Word's AxMath conversion bridge.

    Optional objects make routing and cleanup testable without opening Office.
    Constructing the production backends does not start either application.
    """

    def __init__(self, *, word_backend: Any = None, ppt_backend: Any = None) -> None:
        if word_backend is None:
            from axmath_mcp.backend import WordBackend

            word_backend = WordBackend()
        self.word = word_backend
        try:
            if ppt_backend is None:
                from axmath_mcp.powerpoint import PowerPointBackend

                ppt_backend = PowerPointBackend(self.word)
            self.ppt = ppt_backend
        except BaseException:
            self.word.shutdown()
            raise

    def __getattr__(self, name: str) -> Any:
        if name in WORD_METHODS:
            return getattr(self.word, name)
        if name in PPT_METHODS:
            return getattr(self.ppt, name)
        raise AttributeError(f"Unknown Office backend method: {name}")

    def health(self) -> dict[str, Any]:
        result = dict(self.word.health())
        powerpoint = self.ppt.health()
        result.update(
            {
                "powerpoint": powerpoint,
                "powerpoint_registered": bool(powerpoint.get("powerpoint_registered")),
                "supported": (
                    "Windows Word inline/display/numbered formulas and PowerPoint native "
                    "AxMath equation shapes on slides"
                ),
                "notes": [
                    *result.get("notes", []),
                    "PowerPoint LaTeX conversion/export also requires desktop Word and AxMath.dotm.",
                ],
            }
        )
        return result

    def shutdown(self) -> None:
        # Release proxies only: neither backend may close the user's documents.
        try:
            self.ppt.shutdown()
        finally:
            self.word.shutdown()

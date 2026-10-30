"""Validate slide geometry before mutating PowerPoint objects.

All coordinates use PowerPoint points. These checks deliberately reject shapes
that do not fit rather than resizing them or moving other slide content.
"""

from __future__ import annotations

import math
import unicodedata


def _number(value: float, name: str, *, positive: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        # Keep one input-boundary exception type, matching Word layout validation.
        raise ValueError(f"{name} must be a finite number of points.")  # noqa: TRY004
    value = float(value)
    if not math.isfinite(value) or (value <= 0 if positive else value < 0):
        bound = "positive" if positive else "nonnegative"
        raise ValueError(f"{name} must be a finite {bound} number of points.")
    return value


def validate_position(left_pt: float, top_pt: float) -> tuple[float, float]:
    """Require an on-slide origin; size/slide bounds are checked by ensure_fits."""
    return _number(left_pt, "left_pt"), _number(top_pt, "top_pt")


def validate_font_size(value: float = 24.0) -> float:
    """Return an explicit base font size in the supported 6–72 point range."""
    value = _number(value, "font_size_pt", positive=True)
    if not 6 <= value <= 72:
        raise ValueError("font_size_pt must be between 6 and 72 points.")
    return value


def validate_number_label(value: str | None) -> str | None:
    """Preserve the exact manual label, or None for an unnumbered equation."""
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip() or len(value) > 64:
        raise ValueError("number_label must contain 1 to 64 visible characters, or be None.")
    if any(unicodedata.category(c) in {"Cc", "Cf", "Cs", "Zl", "Zp"} for c in value):
        raise ValueError("number_label must be single-line text without control characters.")
    return value


def ensure_fits(
    left: float,
    top: float,
    width: float,
    height: float,
    slide_width: float,
    slide_height: float,
    number_width: float = 0,
    gap: float = 12,
) -> None:
    """Require the equation and optional adjacent number to fit on the slide.

    The number's horizontal allocation starts at left + width + gap. A zero
    number_width means no number and therefore no extra gap. Callers placing a
    number taller than the equation should validate that label's box separately.
    """
    left, top = validate_position(left, top)
    width = _number(width, "width", positive=True)
    height = _number(height, "height", positive=True)
    slide_width = _number(slide_width, "slide_width", positive=True)
    slide_height = _number(slide_height, "slide_height", positive=True)
    number_width = _number(number_width, "number_width")
    gap = _number(gap, "gap")
    total_width = width + (gap + number_width if number_width > 0 else 0)
    if left + total_width > slide_width or top + height > slide_height:
        raise ValueError(
            f"Equation/number box ({total_width:g} × {height:g} pt) at "
            f"({left:g}, {top:g}) does not fit within the "
            f"{slide_width:g} × {slide_height:g} pt slide. "
            "Choose a different position or smaller font; no automatic scaling is applied."
        )

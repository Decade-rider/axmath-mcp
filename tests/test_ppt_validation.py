"""Input-boundary and edge-of-slide checks without PowerPoint installed."""

import pytest

from axmath_mcp.ppt_validation import (
    ensure_fits,
    validate_font_size,
    validate_number_label,
    validate_position,
)


def test_position_accepts_zero_origin_and_returns_points_as_floats():
    assert validate_position(0, 12) == (0.0, 12.0)
    assert validate_position(8.25, 18.5) == (8.25, 18.5)
    assert all(isinstance(v, float) for v in validate_position(0, 12))


@pytest.mark.parametrize(
    "value", [-0.01, float("nan"), float("inf"), -float("inf"), True, "12", None]
)
@pytest.mark.parametrize("axis", [0, 1])
def test_position_rejects_negative_nonfinite_and_nonnumeric_coordinates(value, axis):
    coordinates = [0.0, 0.0]
    coordinates[axis] = value
    with pytest.raises(ValueError):
        validate_position(*coordinates)


def test_font_default_and_supported_boundaries():
    assert validate_font_size() == 24.0
    assert validate_font_size(6) == 6.0
    assert validate_font_size(72) == 72.0
    assert validate_font_size(12.5) == 12.5


@pytest.mark.parametrize("value", [0, 5.99, 72.01, float("nan"), float("inf"), False, "24", None])
def test_font_invalid_values_are_not_silently_defaulted_or_clamped(value):
    with pytest.raises(ValueError):
        validate_font_size(value)


@pytest.mark.parametrize("value", [None, "(1)", "(A.1)", "（附录 A）", " (Exact label) ", "a" * 64])
def test_manual_labels_preserve_exact_content(value):
    assert validate_number_label(value) == value


@pytest.mark.parametrize(
    "value",
    [
        "",
        " \t ",
        "a" * 65,
        1,
        True,
        "a\nb",
        "a\rb",
        "a\tb",
        "a\x00b",
        "a\u200bb",
        "a\u2028b",
        "a\u2029b",
        "a\ud800b",
    ],
)
def test_labels_reject_empty_multiline_controls_and_invalid_types(value):
    with pytest.raises(ValueError):
        validate_number_label(value)


def test_exact_fit_is_allowed_without_scaling():
    assert ensure_fits(10, 20, 90, 30, 100, 50) is None


def test_gap_is_ignored_when_there_is_no_number():
    ensure_fits(0, 0, 100, 50, 100, 50, number_width=0, gap=12)


def test_adjacent_number_reserves_its_width_and_gap():
    ensure_fits(10, 20, 60, 30, 100, 50, number_width=18, gap=12)
    with pytest.raises(ValueError, match="does not fit"):
        ensure_fits(10, 20, 60, 30, 100, 50, number_width=18.1, gap=12)


@pytest.mark.parametrize("kwargs", [{"width": 101}, {"height": 51}, {"left": 0.01}, {"top": 0.01}])
def test_each_slide_edge_rejects_overflow(kwargs):
    box = {"left": 0, "top": 0, "width": 100, "height": 50, "slide_width": 100, "slide_height": 50}
    box.update(kwargs)
    with pytest.raises(ValueError, match="does not fit"):
        ensure_fits(**box)


@pytest.mark.parametrize("key", ["width", "height", "slide_width", "slide_height"])
@pytest.mark.parametrize("value", [0, -1, float("inf"), float("nan"), True])
def test_geometry_requires_positive_finite_dimensions(key, value):
    box = {"left": 0, "top": 0, "width": 100, "height": 50, "slide_width": 200, "slide_height": 100}
    box[key] = value
    with pytest.raises(ValueError):
        ensure_fits(**box)


@pytest.mark.parametrize("key", ["number_width", "gap"])
@pytest.mark.parametrize("value", [-1, float("inf"), float("nan"), True])
def test_number_allocation_requires_finite_nonnegative_dimensions(key, value):
    box = {"left": 0, "top": 0, "width": 100, "height": 50, "slide_width": 200, "slide_height": 100}
    box[key] = value
    with pytest.raises(ValueError):
        ensure_fits(**box)


def test_finite_but_overflowing_sum_is_rejected():
    with pytest.raises(ValueError, match="does not fit"):
        ensure_fits(1e308, 0, 1e308, 10, 1e308, 20)

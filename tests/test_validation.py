import pytest

from axmath_mcp.validation import (
    LayoutOptions,
    latex_for_layout,
    normalize_latex,
    validate_layout_options,
    validate_local_path,
)


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        (" x^2 + y^2 ", r"\(x^2 + y^2\)"),
        (r"\frac{a}{b}", r"\(\frac{a}{b}\)"),
        (r"\(x+1\)", r"\(x+1\)"),
        (r"\[x+1\]", r"\[x+1\]"),
        ("$x+1$", "$x+1$"),
        ("$$x+1$$", "$$x+1$$"),
        (r"\text{cost \$5}", r"\(\text{cost \$5}\)"),
        (r"$\text{cost \$5}$", r"$\text{cost \$5}$"),
        (r"\begin{bmatrix}1&2\\3&4\end{bmatrix}", r"\(\begin{bmatrix}1&2\\3&4\end{bmatrix}\)"),
        ("a\t+\nb", "\\(a\t+\nb\\)"),
    ],
)
def test_normalize_single_math_expression(source, expected):
    assert normalize_latex(source) == expected


@pytest.mark.parametrize(
    "source",
    [
        "",
        "  ",
        "$$",
        "$ $",
        r"\(\)",
        r"\[ \]",
        "$$$$",
        "$x",
        "x$",
        "$$x$",
        "$x$$",
        r"\(x\]",
        r"\]x\[",
        "$x$ + $y$",
        r"\(x\)\(y\)",
        r"\[x\(y\)\]",
        "before $x$",
        "$x$ after",
        r"\(x\) after",
        "x $y$ z",
        "$x$y$",
        "a\x00b",
        "a\rb",
        "a\x1fb",
        "a\x7fb",
        "a\x85b",
    ],
)
def test_reject_empty_multiple_mismatched_or_unsafe_content(source):
    with pytest.raises(ValueError):
        normalize_latex(source)


@pytest.mark.parametrize(
    "source",
    [
        r"\input{secret}",
        r"\include{x}",
        r"\write18{command}",
        r"\openout1=file",
        r"\read1to\a",
        r"\usepackage{package}",
        r"\documentclass{article}",
        r"\begin{document}x\end{document}",
        r"\begin { document } x",
        r"\end{document}",
        r"\catcode`x=0",
        r"\csname input\endcsname",
        r"\def\foo{x}",
        r"\gdef\foo{x}",
        r"\newcommand{\foo}{x}",
        r"\renewcommand{\foo}{x}",
        r"\let\foo\input",
        r"\directlua{os.execute('x')}",
        r"\\\input{secret}",
    ],
)
def test_reject_document_and_macro_commands(source):
    with pytest.raises(ValueError):
        normalize_latex(source)


def test_length_limit_checks_entire_input_before_trimming():
    assert normalize_latex("x" * 20_000) == r"\(" + "x" * 20_000 + r"\)"
    with pytest.raises(ValueError, match="limit"):
        normalize_latex(" " * 20_000 + "x")


def test_control_symbols_are_not_misread_as_control_words():
    assert normalize_latex(r"\\input") == r"\(\\input\)"
    assert normalize_latex(r"\text{50\%}") == r"\(\text{50\%}\)"


@pytest.mark.parametrize("path", [r"C:\Documents\file.docx", "E:/论文/公式.DOCX"])
def test_accept_absolute_local_windows_paths(path):
    assert str(validate_local_path(path)).replace("\\", "/") == path.replace("\\", "/")


@pytest.mark.parametrize(
    "path",
    [
        "",
        "file.docx",
        r".\file.docx",
        r"~\file.docx",
        r"C:file.docx",
        r"\file.docx",
        "/file.docx",
        r"\\server\share\file.docx",
        r"\\?\C:\file.docx",
        r"\\.\C:\file.docx",
        "https://example.com/file.docx",
        r"C:\file.doc",
        r"C:\file.docm",
        r"C:\file.docx:stream",
        r"C:\file:stream.docx",
        r"C:\*.docx",
        r"C:\file?.docx",
        r"C:\NUL.docx",
        r"C:\CON\file.docx",
        r"C:\COM1.docx",
        r"C:\name.\file.docx",
        r"C:\name \file.docx",
        "C:\\file.docx\x00",
    ],
)
def test_reject_nonlocal_ambiguous_or_disallowed_paths(path):
    with pytest.raises(ValueError):
        validate_local_path(path)


def test_custom_suffixes_are_case_insensitive():
    assert validate_local_path(r"C:\formula.TEX", {".tex"}).name == "formula.TEX"
    with pytest.raises(ValueError):
        validate_local_path(r"C:\formula.docx", {".tex"})


def test_default_and_numbered_layout_options():
    assert validate_layout_options() == LayoutOptions("inline", None, None)
    assert validate_layout_options("numbered", 12, None) == LayoutOptions("numbered", 12.0, None)
    assert validate_layout_options("numbered", 10.5, "(附录 A.1)").number_label == "(附录 A.1)"
    assert validate_layout_options("display", 6).font_size_pt == 6.0
    assert validate_layout_options("display", 72).font_size_pt == 72.0


@pytest.mark.parametrize(
    "size", [float("nan"), float("inf"), -float("inf"), 0, 5.9, 72.1, True, "12"]
)
def test_reject_invalid_equation_sizes_before_word_mutation(size):
    with pytest.raises(ValueError, match="finite number"):
        validate_layout_options("display", size)


@pytest.mark.parametrize("layout", ["", "center", "INLINE", "number"])
def test_reject_unknown_layout(layout):
    with pytest.raises(ValueError, match="layout"):
        validate_layout_options(layout)


@pytest.mark.parametrize("layout", ["inline", "display"])
def test_cannot_attach_number_to_unnumbered_layout(layout):
    with pytest.raises(ValueError, match="only valid"):
        validate_layout_options(layout, number_label="(1)")


@pytest.mark.parametrize(
    "label", ["", " ", "x" * 65, "(1)\n", "(1)\r", "1\t2", "\x00", "A\u2028B", "A\u202eB"]
)
def test_reject_number_labels_that_can_escape_the_number_position(label):
    with pytest.raises(ValueError):
        validate_layout_options("numbered", number_label=label)


@pytest.mark.parametrize("source", ["x_i^2", "$x_i^2$", "$$x_i^2$$", r"\(x_i^2\)", r"\[x_i^2\]"])
def test_layout_style_overrides_outer_delimiters_without_nesting(source):
    assert latex_for_layout(source, "inline") == r"\(x_i^2\)"
    assert latex_for_layout(source, "display") == r"\[x_i^2\]"
    assert latex_for_layout(source, "numbered") == r"\[x_i^2\]"


def test_layout_style_preserves_matrix_body_and_still_validates_source():
    matrix = r"\begin{bmatrix}1&2\\3&4\end{bmatrix}"
    assert latex_for_layout(matrix, "numbered") == r"\[" + matrix + r"\]"
    with pytest.raises(ValueError):
        latex_for_layout("$x$ and $y$", "display")
    with pytest.raises(ValueError):
        latex_for_layout(r"\input{file}", "display")

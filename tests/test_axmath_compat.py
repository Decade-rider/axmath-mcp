"""Check parser-hazard detection without invoking Word or AxMath."""

import pytest

from axmath_mcp.axmath_compat import check_axmath_input


@pytest.mark.parametrize(
    "source",
    [
        r"\gcd(a,b)",
        r"d=\gcd\{n\ge1:p_{ii}^{(n)}>0\}",
        r"\gcd_{n\ge1} n",
        r"\gcd1",
        r"\text{inside \gcd(a,b)}",
        r"\\\gcd(a,b)",
        "x % comment\n\\gcd(a,b)",
        r"\%\gcd(a,b)",
    ],
)
def test_known_direct_gcd_is_rejected_with_explicit_nonrewriting_guidance(source):
    with pytest.raises(ValueError, match=r"AxMath 2\.7061") as caught:
        check_axmath_input(source)
    message = str(caught.value)
    assert r"\operatorname{gcd}" in message
    assert "subscripts or limits" in message
    assert "no automatic rewrite" in message


@pytest.mark.parametrize(
    "source",
    [
        r"\forall x\in\mathbf{R},x\ge0",
        r"\forall x\geq0",
        r"\mathbf{p}_{t+1}=P\mathbf{p}_t",
        r"\operatorname{gcd}(a,b)",
        r"\mathrm{gcd}(a,b)",
        r"\gcdextra(a,b)",
        r"\GCD(a,b)",
        r"gcd(a,b)",
        r"\\gcd(a,b)",
        r"\\\\gcd(a,b)",
        "% \\gcd(a,b) {\nx+1",
        "x % } \\gcd(a,b)",
        "{x % } \\gcd(a,b)\n}",
        "{x % } \\gcd(a,b)\r\n}",
        "\\\\%\\gcd(a,b)\nx",
        r"\left\{x\right\}",
        r"\frac{a+\{b\}}{c}",
        r"\begin{pmatrix}a&b\\c&d\end{pmatrix}",
        r"\text{100\% complete}",
        r"x\}",
        "x\\",
    ],
)
def test_guard_does_not_create_a_global_allowlist_or_misread_escaped_symbols(source):
    assert check_axmath_input(source) is None


@pytest.mark.parametrize(
    "source",
    [
        r"\frac{a}{b",
        r"x}",
        r"}{",
        r"{{x}",
        r"\\{x",
        r"\{x}",
        "{x% closing brace is commented out }",
        "x% opening brace is commented out {\n}",
    ],
)
def test_unmatched_real_braces_are_rejected_before_the_vendor_parser(source):
    with pytest.raises(ValueError, match="Unbalanced LaTeX braces"):
        check_axmath_input(source)


def test_error_positions_refer_to_the_original_input():
    with pytest.raises(ValueError, match="character 4"):
        check_axmath_input("x +}")
    with pytest.raises(ValueError, match="character 3"):
        check_axmath_input("x {y")


def test_nontext_input_is_not_reinterpreted_as_math():
    with pytest.raises(TypeError, match="string"):
        check_axmath_input(None)

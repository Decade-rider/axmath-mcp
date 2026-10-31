"""Narrow compatibility checks before invoking the installed AxMath parser.

This is not a complete TeX parser or an exhaustive supported-command list. The
checks prevent a reported AxMath 2.7061 integration hazard and obvious unmatched
grouping braces without discarding or silently rewriting mathematical content.
"""

from __future__ import annotations


def _ascii_letter(char: str) -> bool:
    return "a" <= char <= "z" or "A" <= char <= "Z"


def check_axmath_input(source: str) -> None:
    r"""Reject direct ``\gcd`` and unmatched unescaped grouping braces.

    TeX control words are scanned as complete, case-sensitive ASCII names.
    Escaped symbols (including ``\\``, ``\%``, ``\{`` and ``\}``) are consumed
    together. Unescaped percent comments end at a newline and are ignored for
    these checks. Other syntax and command support remain AxMath's responsibility.
    """
    if not isinstance(source, str):
        raise TypeError("source must be a string.")
    opening_braces: list[int] = []
    index = 0
    while index < len(source):
        char = source[index]
        if char == "%":
            while index < len(source) and source[index] not in "\r\n":
                index += 1
            continue
        if char == "\\":
            following = index + 1
            if following < len(source) and _ascii_letter(source[following]):
                end = following + 1
                while end < len(source) and _ascii_letter(source[end]):
                    end += 1
                if source[following:end] == "gcd":
                    raise ValueError(
                        "\\gcd is blocked by the AxMath 2.7061 compatibility guard before "
                        "conversion. For ordinary function notation, try "
                        "\\operatorname{gcd}(a,b). Expressions with subscripts or limits "
                        "need an explicit operator/limits form; no automatic rewrite was made."
                    )
                index = end
            else:
                # A control symbol consumes one following character; a final
                # backslash is left for AxMath's full syntax validation.
                index += 2
            continue
        if char == "{":
            opening_braces.append(index)
        elif char == "}":
            if not opening_braces:
                raise ValueError(
                    f"Unbalanced LaTeX braces: unexpected closing brace at character {index + 1}."
                )
            opening_braces.pop()
        index += 1
    if opening_braces:
        raise ValueError(
            "Unbalanced LaTeX braces: opening brace at character "
            f"{opening_braces[-1] + 1} has no matching closing brace."
        )

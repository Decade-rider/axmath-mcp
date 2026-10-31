"""Validate MCP inputs before handing them to Word or the AxMath converter.

This is an input boundary, not a TeX parser. AxMath determines which mathematical
commands it supports. In particular, accepting a formula here does not certify its
mathematical meaning or its rendering.
"""

from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Collection
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Literal

from axmath_mcp.axmath_compat import check_axmath_input

MAX_LATEX_LENGTH = 20_000
EquationLayout = Literal["inline", "display", "numbered"]


@dataclass(frozen=True)
class LayoutOptions:
    layout: EquationLayout
    font_size_pt: float | None
    number_label: str | None


def validate_layout_options(
    layout: str = "inline", font_size_pt: float | None = None, number_label: str | None = None
) -> LayoutOptions:
    """Validate paragraph-layout inputs before any document or converter mutation."""
    if layout not in {"inline", "display", "numbered"}:
        raise ValueError("layout must be inline, display, or numbered.")
    if font_size_pt is not None:
        if (
            isinstance(font_size_pt, bool)
            or not isinstance(font_size_pt, (float, int))
            or not math.isfinite(font_size_pt)
            or not 6 <= font_size_pt <= 72
        ):
            raise ValueError("font_size_pt must be a finite number between 6 and 72 points.")
        font_size_pt = float(font_size_pt)
    if number_label is not None:
        if layout != "numbered":
            raise ValueError("number_label is only valid with layout='numbered'.")
        if not isinstance(number_label, str) or not number_label.strip() or len(number_label) > 64:
            raise ValueError("number_label must contain 1 to 64 visible characters.")
        if any(
            unicodedata.category(char) in {"Cc", "Cf", "Cs", "Zl", "Zp"} for char in number_label
        ):
            raise ValueError("number_label must be single-line text without control characters.")
    return LayoutOptions(layout, font_size_pt, number_label)


# Accept math expressions, not a TeX program, document, or macro definition.
_BLOCKED_COMMANDS = frozenset(
    {
        "input",
        "include",
        "includeonly",
        "includegraphics",
        "includepdf",
        "write",
        "openout",
        "closeout",
        "openin",
        "closein",
        "read",
        "readline",
        "usepackage",
        "RequirePackage",
        "documentclass",
        "LoadClass",
        "catcode",
        "csname",
        "endcsname",
        "def",
        "edef",
        "gdef",
        "xdef",
        "let",
        "futurelet",
        "newcommand",
        "renewcommand",
        "providecommand",
        "newenvironment",
        "renewenvironment",
        "DeclareRobustCommand",
        "DeclareMathOperator",
        "newtheorem",
        "newif",
        "newcount",
        "newtoks",
        "chardef",
        "mathchardef",
        "NewDocumentCommand",
        "RenewDocumentCommand",
        "ProvideDocumentCommand",
        "DeclareDocumentCommand",
        "newread",
        "newwrite",
        "immediate",
        "special",
        "directlua",
        "latelua",
        "luadirect",
        "scantokens",
        "everyjob",
        "everyeof",
        "everypar",
        "everymath",
        "everydisplay",
        "loop",
        "repeat",
        "expandafter",
        "noexpand",
        "global",
        "long",
        "outer",
        "protected",
        "endinput",
        "shipout",
        "output",
    }
)
_CLOSING_DELIMITER = {"$": "$", "$$": "$$", r"\(": r"\)", r"\[": r"\]"}
_COMMAND = re.compile(r"[A-Za-z]+")
_ENVIRONMENT_NAME = re.compile(r"\s*\{\s*([^{}]+?)\s*\}")
_RESERVED_WINDOWS_NAME = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])", re.IGNORECASE)


def normalize_latex(latex: str) -> str:
    r"""Return exactly one delimited math expression, adding ``\(...\)`` if bare.

    Existing $, $$, \( and \[ delimiters must form one complete matching pair.
    Escaped dollar signs are ordinary mathematical/text content. Document and
    macro commands are rejected; matrix/align-style math environments are allowed.
    """
    if not isinstance(latex, str):
        raise TypeError("latex must be a string.")
    if len(latex) > MAX_LATEX_LENGTH:
        raise ValueError(f"latex exceeds the {MAX_LATEX_LENGTH} character limit.")
    if any(unicodedata.category(char) == "Cc" and char not in "\t\n" for char in latex):
        raise ValueError("latex contains a control character; only tabs and newlines are allowed.")
    source = latex.strip()
    if not source:
        raise ValueError("latex must contain a nonempty mathematical expression.")

    delimiters: list[tuple[str, int, int]] = []
    index = 0
    while index < len(source):
        char = source[index]
        if char == "\\":
            if index + 1 >= len(source):
                # Syntax validity beyond the input boundary belongs to AxMath.
                index += 1
                continue
            next_char = source[index + 1]
            if next_char in "()[]":
                delimiters.append((source[index : index + 2], index, index + 2))
                index += 2
                continue
            command = _COMMAND.match(source, index + 1)
            if command:
                name = command.group()
                if name in _BLOCKED_COMMANDS:
                    raise ValueError(f"TeX command \\{name} is not allowed in a math expression.")
                if name in {"begin", "end"}:
                    environment = _ENVIRONMENT_NAME.match(source, command.end())
                    if environment and environment.group(1).strip() == "document":
                        raise ValueError("Full TeX document environments are not allowed.")
                index = command.end()
            else:
                # Skip one escaped symbol (including \\, \$, \%, \{ and \}).
                index += 2
            continue
        if char == "$":
            end = index + (2 if source.startswith("$$", index) else 1)
            delimiters.append((source[index:end], index, end))
            index = end
            continue
        index += 1

    if not delimiters:
        return r"\(" + source + r"\)"
    if len(delimiters) != 2:
        raise ValueError("Supply exactly one math expression with one matching delimiter pair.")
    opening, closing = delimiters
    if _CLOSING_DELIMITER.get(opening[0]) != closing[0]:
        raise ValueError("The math delimiters do not match.")
    if opening[1] != 0 or closing[2] != len(source):
        raise ValueError("Text outside the math delimiters is not allowed.")
    if not source[opening[2] : closing[1]].strip():
        raise ValueError("The mathematical expression inside the delimiters is empty.")
    return source


def latex_for_layout(source: str, layout: str) -> str:
    r"""Select AxMath's inline/display math style without duplicating delimiters.

    Paragraph placement is handled separately by Word. A numbered formula uses
    display math; its number is a Word field or label, never part of the LaTeX.
    """
    options = validate_layout_options(layout)
    normalized = normalize_latex(source)
    check_axmath_input(normalized)
    width = 1 if normalized.startswith("$") and not normalized.startswith("$$") else 2
    body = normalized[width:-width]
    if options.layout == "inline":
        return r"\(" + body + r"\)"
    return r"\[" + body + r"\]"


def validate_local_path(path: str, suffixes: Collection[str] = frozenset({".docx"})) -> Path:
    """Require an absolute drive-letter Windows path with an allowed file suffix.

    This checks path syntax, not existence, access rights, or symlink destinations.
    UNC shares, device namespaces, ADS, wildcards and reserved device names are
    rejected. Filesystem access and overwrite decisions belong to the caller.
    """
    if not isinstance(path, str) or not path:
        raise ValueError("path must be a nonempty absolute Windows filesystem path.")
    if any(unicodedata.category(char) == "Cc" for char in path):
        raise ValueError("path contains a control character.")
    parsed = PureWindowsPath(path)
    if not re.fullmatch(r"[A-Za-z]:", parsed.drive) or not parsed.is_absolute():
        raise ValueError("Use an absolute local drive path, for example C:\\Documents\\file.docx.")
    if any(char in path[2:] for char in ':<>"|?*'):
        raise ValueError("Device, stream, wildcard and invalid Windows path syntax is not allowed.")
    for part in parsed.parts[1:]:
        if part == "..":
            continue
        if part.endswith((" ", ".")):
            raise ValueError("Windows path components must not end in a dot or space.")
        if _RESERVED_WINDOWS_NAME.fullmatch(part.split(".", 1)[0]):
            raise ValueError("Windows device names cannot be used as filesystem paths.")
    allowed = {suffix.lower() for suffix in suffixes}
    if parsed.suffix.lower() not in allowed:
        raise ValueError("The file extension must be one of: " + ", ".join(sorted(allowed)))
    return Path(path)

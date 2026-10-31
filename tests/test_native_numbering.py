"""Native numbering boundaries and temporary document setting restoration."""

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from axmath_mcp.layout import refresh_numbers
from axmath_mcp.native_numbering import preflight_native_numbering, temporary_property


class Properties(list):
    def __call__(self, name):
        return next(p for p in self if p.Name == name)

    def Add(self, name, _linked, _kind, value):
        prop = SimpleNamespace(Name=name, Value=value)
        prop.Delete = lambda: self.remove(prop)
        self.append(prop)


def document():
    return SimpleNamespace(
        CustomDocumentProperties=Properties(),
        Application=SimpleNamespace(Templates=[SimpleNamespace(Name="AxMath.dotm")]),
    )


@pytest.mark.parametrize("existing", [None, "0", "1"])
def test_temporary_vendor_settings_restored_after_macro_failure(existing):
    doc = document()
    if existing is not None:
        doc.CustomDocumentProperties.Add("AMDeferFieldUpdate", False, 4, existing)
    with pytest.raises(RuntimeError), temporary_property(doc, "AMDeferFieldUpdate", "1"):
        assert doc.CustomDocumentProperties("AMDeferFieldUpdate").Value == "1"
        raise RuntimeError("macro failed")
    assert [p.Value for p in doc.CustomDocumentProperties] == (
        [] if existing is None else [existing]
    )


def test_custom_chapter_settings_are_never_overwritten():
    doc = document()
    doc.CustomDocumentProperties.Add("AMEquationNumber2", False, 4, "(#S1.#E1)")
    with pytest.raises(ValueError, match="custom/chapter"):
        preflight_native_numbering(doc)
    assert doc.CustomDocumentProperties("AMEquationNumber2").Value == "(#S1.#E1)"


def test_target_requires_its_own_loaded_axmath_template():
    doc = document()
    doc.Application.Templates = []
    with pytest.raises(RuntimeError, match="target Word"):
        preflight_native_numbering(doc)


def test_native_refresh_skips_legacy_and_external_fields():
    def field(code, value):
        return SimpleNamespace(
            Type=12,
            Code=SimpleNamespace(Text=code),
            Locked=False,
            Update=Mock(return_value=True),
            Result=SimpleNamespace(Text=value),
        )

    fields = [
        field(r"SEQ AMEqn \h", ""),
        field(r"SEQ AMEqn \c", "1"),
        field("SEQ AxMathEquation", "9"),
        field("INCLUDETEXT external.docx", "external"),
    ]
    assert refresh_numbers(SimpleNamespace(Fields=fields)) == {"updated_count": 1, "numbers": ["1"]}
    fields[0].Update.assert_called_once()
    fields[1].Update.assert_called_once()
    fields[2].Update.assert_not_called()
    fields[3].Update.assert_not_called()


def test_word_tools_offer_only_native_numbering():
    import asyncio

    from axmath_mcp.server import mcp

    schemas = {t.name: t.inputSchema for t in asyncio.run(mcp.list_tools())}
    assert "word_migrate_equation_numbers" not in schemas
    assert "number_label" not in schemas["word_insert_axmath"]["properties"]
    assert "number_label" not in schemas["word_format_axmath"]["properties"]

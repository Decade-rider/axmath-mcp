"""Check shared backend routing and MCP contracts without launching Office."""

from __future__ import annotations

import asyncio
import sys
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock

import pytest

from axmath_mcp import server
from axmath_mcp.office import OfficeBackend


def test_word_and_ppt_routes_share_owned_backends_and_reject_unknown_methods():
    word = SimpleNamespace(insert_equation=Mock(return_value={"kind": "word"}))
    ppt = SimpleNamespace(ppt_insert_equation=Mock(return_value={"kind": "ppt"}))
    office = OfficeBackend(word_backend=word, ppt_backend=ppt)

    assert office.insert_equation(document_id="d", latex="x") == {"kind": "word"}
    assert office.ppt_insert_equation(presentation_id="p", slide_id=257, latex="x") == {
        "kind": "ppt"
    }
    word.insert_equation.assert_called_once_with(document_id="d", latex="x")
    ppt.ppt_insert_equation.assert_called_once_with(presentation_id="p", slide_id=257, latex="x")
    with pytest.raises(AttributeError, match="Unknown Office"):
        _ = office.arbitrary_macro


def test_shared_health_preserves_word_fields_and_does_not_mutate_backend_result():
    source = {"word_registered": True, "version": "0.3.0", "notes": ["word note"]}
    details = {"powerpoint_registered": True, "conversion_backend": "Word AxMath.dotm"}
    word = SimpleNamespace(health=Mock(return_value=source))
    ppt = SimpleNamespace(health=Mock(return_value=details))

    result = OfficeBackend(word_backend=word, ppt_backend=ppt).health()

    assert result["word_registered"] is True and result["version"] == "0.3.0"
    assert result["powerpoint_registered"] is True and result["powerpoint"] == details
    assert "PowerPoint" in result["supported"]
    assert source == {"word_registered": True, "version": "0.3.0", "notes": ["word note"]}


def test_release_ppt_before_word_even_when_ppt_shutdown_fails():
    events = []

    def ppt_close():
        events.append("ppt")
        raise RuntimeError("PPT release failed")

    office = OfficeBackend(
        word_backend=SimpleNamespace(shutdown=lambda: events.append("word")),
        ppt_backend=SimpleNamespace(shutdown=ppt_close),
    )
    with pytest.raises(RuntimeError, match="PPT release failed"):
        office.shutdown()
    assert events == ["ppt", "word"]


def test_ppt_construction_failure_releases_the_shared_word_backend(monkeypatch):
    word = SimpleNamespace(shutdown=Mock())
    factory = Mock(side_effect=RuntimeError("PPT setup failed"))
    monkeypatch.setitem(
        sys.modules, "axmath_mcp.powerpoint", SimpleNamespace(PowerPointBackend=factory)
    )
    with pytest.raises(RuntimeError, match="PPT setup failed"):
        OfficeBackend(word_backend=word)
    factory.assert_called_once_with(word)
    word.shutdown.assert_called_once_with()


def test_all_tools_expose_ids_without_context_and_ppt_slide_id_is_an_integer():
    tools = asyncio.run(server.mcp.list_tools())
    by_name = {tool.name: tool for tool in tools}
    assert len(tools) == 25
    assert len([name for name in by_name if name.startswith("ppt_")]) == 12
    assert len([name for name in by_name if name.startswith("word_")]) == 12
    for tool in tools:
        assert "ctx" not in tool.inputSchema["properties"]
    schema = by_name["ppt_insert_axmath"].inputSchema
    assert {"presentation_id", "slide_id", "latex"}.issubset(schema["required"])
    assert schema["properties"]["slide_id"]["type"] == "integer"
    assert schema["properties"]["font_size_pt"]["default"] == 24
    assert by_name["ppt_get_equation_latex"].annotations.readOnlyHint is True
    assert by_name["ppt_update_axmath"].annotations.destructiveHint is True


def test_ppt_insert_frontend_forwards_stable_id_position_and_label_unchanged():
    worker = SimpleNamespace(call=AsyncMock(return_value={"equation_id": "eq"}))
    ctx = SimpleNamespace(request_context=SimpleNamespace(lifespan_context={"worker": worker}))
    result = asyncio.run(
        server.ppt_insert_axmath(
            ctx,
            presentation_id="pres",
            slide_id=257,
            latex=r"\frac{a}{b}",
            left_pt=90.5,
            top_pt=120.25,
            font_size_pt=28,
            number_label="(A.1)",
        )
    )
    assert result == {"equation_id": "eq"}
    worker.call.assert_awaited_once_with(
        "ppt_insert_equation",
        presentation_id="pres",
        slide_id=257,
        latex=r"\frac{a}{b}",
        left_pt=90.5,
        top_pt=120.25,
        font_size_pt=28,
        number_label="(A.1)",
    )


def test_ppt_update_omits_size_override_and_uses_ppt_backend_method():
    worker = SimpleNamespace(call=AsyncMock(return_value={"equation_id": "eq"}))
    ctx = SimpleNamespace(request_context=SimpleNamespace(lifespan_context={"worker": worker}))
    asyncio.run(server.ppt_update_axmath(ctx, "pres", 257, "eq", "x^2"))
    worker.call.assert_awaited_once_with(
        "ppt_update_equation",
        presentation_id="pres",
        slide_id=257,
        equation_id="eq",
        latex="x^2",
        font_size_pt=None,
    )

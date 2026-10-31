"""Stdio MCP entry point. All Office work is delegated to the dedicated STA."""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any

from mcp.server.fastmcp import Context, FastMCP
from mcp.types import ToolAnnotations

from axmath_mcp.sta import StaExecutor
from axmath_mcp.validation import EquationLayout


@asynccontextmanager
async def lifespan(_server: FastMCP) -> AsyncIterator[dict[str, StaExecutor]]:
    worker = StaExecutor()
    try:
        await worker.start()
        yield {"worker": worker}
    finally:
        await worker.close()


mcp = FastMCP(
    "AxMath Office",
    instructions=(
        "Use local Windows Word, PowerPoint and AxMath to insert and edit native equations. "
        "First inspect axmath_health, then list documents or presentations. Use word_* tools "
        "for Word and ppt_* tools for PowerPoint. Always use the returned "
        "document_id, never guess a document. Obtain equation_id from word_list_equations. "
        "Choose inline for text mixed with formulas, display for a centered equation, "
        "or numbered for a centered equation with a right-aligned number. Display and "
        "numbered formulas use dedicated paragraphs, preserving surrounding text. "
        "PowerPoint uses standalone OLE shapes, not inline text-frame equations. Use the "
        "stable slide_id from ppt_list_slides, not its index, and coordinates in points "
        "(72 pt per inch). PowerPoint numbers are optional fixed labels, not Word SEQ fields. "
        "PowerPoint conversion/export also requires Word and AxMath.dotm. "
        "Changes stay unsaved until the corresponding save tool is explicitly requested. "
        "Do not operate either Office app concurrently with the user during conversion. "
        "Never retry a mutation blindly after cancellation or a lost response: inspect "
        "the target document or slide first, because a queued Office call may still finish."
    ),
    lifespan=lifespan,
)

READ_ONLY = ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=False)
WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=False)
EDIT = ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=False)


async def _call(ctx: Context, method: str, **kwargs: Any) -> dict[str, Any]:
    worker: StaExecutor = ctx.request_context.lifespan_context["worker"]
    return await worker.call(method, **kwargs)


@mcp.tool(annotations=READ_ONLY)
async def axmath_health(ctx: Context) -> dict[str, Any]:
    """Inspect Windows, Word, PowerPoint and AxMath integration without starting Office."""
    return await _call(ctx, "health")


@mcp.tool(annotations=READ_ONLY)
async def word_list_documents(ctx: Context) -> dict[str, Any]:
    """List reachable open Word documents and stable IDs to use in subsequent calls."""
    return await _call(ctx, "list_documents")


@mcp.tool(annotations=WRITE)
async def word_create_document(ctx: Context, visible: bool = True) -> dict[str, Any]:
    """Create a new unsaved Word document; return its document_id. Visible by default."""
    return await _call(ctx, "create_document", visible=visible)


@mcp.tool(annotations=WRITE)
async def word_open_document(
    ctx: Context, path: str, read_only: bool = False, visible: bool = True
) -> dict[str, Any]:
    """Open an existing local document by absolute path and return its document_id.

    Opening is explicit; this does not save changes. Use read_only for inspection.
    """
    return await _call(ctx, "open_document", path=path, read_only=read_only, visible=visible)


@mcp.tool(annotations=READ_ONLY)
async def word_document_info(ctx: Context, document_id: str) -> dict[str, Any]:
    """Inspect a document, its current selection and available insertion locations."""
    return await _call(ctx, "get_document_info", document_id=document_id)


@mcp.tool(annotations=READ_ONLY)
async def word_list_equations(ctx: Context, document_id: str) -> dict[str, Any]:
    """List native AxMath equations and their equation_id in the selected document."""
    return await _call(ctx, "list_equations", document_id=document_id)


@mcp.tool(annotations=EDIT)
async def word_insert_axmath(
    ctx: Context,
    document_id: str,
    latex: str,
    position: str = "end",
    bookmark: str | None = None,
    expected_selection: str | None = None,
    layout: EquationLayout = "inline",
    font_size_pt: float | None = None,
) -> dict[str, Any]:
    r"""Convert LaTeX to an editable AxMath OLE equation in a specific Word document.

    Supply bare math or one matching outer delimiter pair. layout selects inline,
    centered display, or centered numbered layout. Display/numbered formulas get a
    dedicated paragraph; insertion splits surrounding text without deleting it.
    font_size_pt optionally sets equation size (6-72 pt). A numbered equation uses
    an AxMath-native automatic number (1), (2), ... .
    position is 'end', 'selection', or 'bookmark'. For selection, inspect word_document_info
    first and pass the exact expected_selection text (empty for a caret), preventing
    accidental replacement of an uninspected selection. For bookmark, supply its
    name; insertion is before its text. The document remains unsaved. Conversion uses
    the clipboard and may temporarily use Word's UI;
    avoid typing or switching its selection during this operation.
    """
    return await _call(
        ctx,
        "insert_equation",
        document_id=document_id,
        latex=latex,
        position=position,
        bookmark=bookmark,
        expected_selection=expected_selection,
        layout=layout,
        font_size_pt=font_size_pt,
    )


@mcp.tool(annotations=EDIT)
async def word_update_axmath(
    ctx: Context, document_id: str, equation_id: str, latex: str
) -> dict[str, Any]:
    """Replace an identified AxMath equation with new LaTeX, preserving its layout and number.

    Obtain equation_id from word_list_equations. Avoid interacting with Word during
    conversion. Do not use an ordinary text index or an equation from another file.
    """
    return await _call(
        ctx, "update_equation", document_id=document_id, equation_id=equation_id, latex=latex
    )


@mcp.tool(annotations=EDIT)
async def word_format_axmath(
    ctx: Context,
    document_id: str,
    equation_id: str,
    layout: EquationLayout = "inline",
    font_size_pt: float | None = None,
) -> dict[str, Any]:
    """Format an existing AxMath equation as inline, centered display, or numbered.

    Display and numbered layouts use dedicated paragraphs, preserving surrounding
    text. font_size_pt optionally sets equation size (6-72 pt). With numbered layout,
    the number is generated by the installed AxMath add-in.
    The document remains unsaved. Avoid editing Word during this operation.
    """
    return await _call(
        ctx,
        "format_equation",
        document_id=document_id,
        equation_id=equation_id,
        layout=layout,
        font_size_pt=font_size_pt,
    )


@mcp.tool(annotations=EDIT)
async def word_refresh_equation_numbers(ctx: Context, document_id: str) -> dict[str, Any]:
    """Refresh native AxMath sequence numbers in document order, without saving.

    Use after inserting, moving, or deleting numbered equations. This tool updates equation numbering only.
    """
    return await _call(ctx, "refresh_equation_numbers", document_id=document_id)


@mcp.tool(annotations=READ_ONLY)
async def word_get_equation_latex(
    ctx: Context, document_id: str, equation_id: str
) -> dict[str, Any]:
    """Export LaTeX from a disposable copy, preserving the original document.

    The AxMath add-in temporarily uses selection and writes the system clipboard.
    """
    return await _call(ctx, "get_equation_latex", document_id=document_id, equation_id=equation_id)


@mcp.tool(annotations=WRITE)
async def word_edit_axmath(ctx: Context, document_id: str, equation_id: str) -> dict[str, Any]:
    """Open the selected equation in AxMath for interactive manual editing.

    Finish or cancel the AxMath editor before requesting another Word operation.
    The user's edits are not automatically saved to disk by this tool.
    """
    return await _call(
        ctx, "open_equation_editor", document_id=document_id, equation_id=equation_id
    )


@mcp.tool(annotations=EDIT)
async def word_save_document(
    ctx: Context, document_id: str, path: str | None = None, overwrite: bool = False
) -> dict[str, Any]:
    """Explicitly save a Word document. Supply an absolute .docx path for a new file.

    With no path, save to the document's existing path. A different existing target
    must only be replaced when overwrite=True is explicitly authorized by the user.
    """
    return await _call(
        ctx, "save_document", document_id=document_id, path=path, overwrite=overwrite
    )


@mcp.tool(annotations=READ_ONLY)
async def ppt_list_presentations(ctx: Context) -> dict[str, Any]:
    """List reachable open PowerPoint presentations with IDs for subsequent calls."""
    return await _call(ctx, "ppt_list_presentations")


@mcp.tool(annotations=WRITE)
async def ppt_create_presentation(ctx: Context, visible: bool = True) -> dict[str, Any]:
    """Create a new unsaved PowerPoint presentation and return its presentation_id.

    Use ppt_add_slide to add a blank slide before inserting an equation.
    """
    return await _call(ctx, "ppt_create_presentation", visible=visible)


@mcp.tool(annotations=WRITE)
async def ppt_open_presentation(
    ctx: Context, path: str, read_only: bool = False, visible: bool = True
) -> dict[str, Any]:
    """Open a local PowerPoint file by absolute path and return its presentation_id.

    Opening does not save changes. Use read_only for inspection.
    """
    return await _call(
        ctx, "ppt_open_presentation", path=path, read_only=read_only, visible=visible
    )


@mcp.tool(annotations=READ_ONLY)
async def ppt_list_slides(ctx: Context, presentation_id: str) -> dict[str, Any]:
    """List slides, stable slide_id values and slide dimensions in points.

    Use the returned slide_id in equation calls; a slide's index is a different value.
    """
    return await _call(ctx, "ppt_list_slides", presentation_id=presentation_id)


@mcp.tool(annotations=WRITE)
async def ppt_add_slide(ctx: Context, presentation_id: str, title: str = "") -> dict[str, Any]:
    """Append a blank slide, optionally with a title, and return its stable slide_id."""
    return await _call(ctx, "ppt_add_slide", presentation_id=presentation_id, title=title)


@mcp.tool(annotations=READ_ONLY)
async def ppt_list_equations(
    ctx: Context, presentation_id: str, slide_id: int | None = None
) -> dict[str, Any]:
    """List native AxMath equation shapes on one slide or throughout a presentation.

    Return equation_id values to use for reading, updating or positioning shapes.
    """
    return await _call(
        ctx, "ppt_list_equations", presentation_id=presentation_id, slide_id=slide_id
    )


@mcp.tool(annotations=WRITE)
async def ppt_insert_axmath(
    ctx: Context,
    presentation_id: str,
    slide_id: int,
    latex: str,
    left_pt: float = 72,
    top_pt: float = 144,
    font_size_pt: float = 24,
    number_label: str | None = None,
) -> dict[str, Any]:
    """Insert an editable AxMath OLE equation shape on a specified PowerPoint slide.

    slide_id is the stable ID from ppt_list_slides, not the slide index. left_pt and
    top_pt locate the equation's upper-left corner in points (72 pt = one inch).
    font_size_pt is the equation's base size. Supply one LaTeX expression. The shape
    is independent of slide text boxes; this does not insert math inside a TextFrame.
    number_label optionally adds a fixed associated label such as '(1)' in a separate
    text box; PowerPoint labels do not use Word SEQ automatic numbering. Bounds are
    checked before committing. Conversion uses desktop Word, AxMath.dotm and the
    clipboard. Avoid interacting with either app during the operation. No auto-save.
    """
    return await _call(
        ctx,
        "ppt_insert_equation",
        presentation_id=presentation_id,
        slide_id=slide_id,
        latex=latex,
        left_pt=left_pt,
        top_pt=top_pt,
        font_size_pt=font_size_pt,
        number_label=number_label,
    )


@mcp.tool(annotations=EDIT)
async def ppt_update_axmath(
    ctx: Context,
    presentation_id: str,
    slide_id: int,
    equation_id: str,
    latex: str,
    font_size_pt: float | None = None,
) -> dict[str, Any]:
    """Replace a PowerPoint AxMath formula while preserving its location and label.

    Obtain IDs from the list tools. The replacement is prepared and checked before
    the previous shape is removed. Omit font_size_pt to retain the recorded size.
    This changes a standalone OLE shape, not text-frame contents, and does not save.
    """
    return await _call(
        ctx,
        "ppt_update_equation",
        presentation_id=presentation_id,
        slide_id=slide_id,
        equation_id=equation_id,
        latex=latex,
        font_size_pt=font_size_pt,
    )


@mcp.tool(annotations=EDIT)
async def ppt_position_axmath(
    ctx: Context,
    presentation_id: str,
    slide_id: int,
    equation_id: str,
    left_pt: float,
    top_pt: float,
) -> dict[str, Any]:
    """Move an AxMath equation and its associated label to a position on the slide.

    The position is the equation's upper-left corner, in points (72 pt per inch).
    This preserves the native equation content and size. It does not save the file.
    """
    return await _call(
        ctx,
        "ppt_position_equation",
        presentation_id=presentation_id,
        slide_id=slide_id,
        equation_id=equation_id,
        left_pt=left_pt,
        top_pt=top_pt,
    )


@mcp.tool(annotations=READ_ONLY)
async def ppt_get_equation_latex(
    ctx: Context, presentation_id: str, slide_id: int, equation_id: str
) -> dict[str, Any]:
    """Export LaTeX from a copy of a native PowerPoint AxMath equation.

    The source slide is preserved. The Word/AxMath bridge temporarily uses the
    system clipboard and Word selection; avoid editing either app during the call.
    """
    return await _call(
        ctx,
        "ppt_get_equation_latex",
        presentation_id=presentation_id,
        slide_id=slide_id,
        equation_id=equation_id,
    )


@mcp.tool(annotations=WRITE)
async def ppt_edit_axmath(
    ctx: Context, presentation_id: str, slide_id: int, equation_id: str
) -> dict[str, Any]:
    """Open a slide's AxMath equation in the native editor for manual changes.

    Finish or cancel the editor before another Office call. This does not save the
    PowerPoint file; use ppt_save_presentation when the edits should be persisted.
    """
    return await _call(
        ctx,
        "ppt_open_equation_editor",
        presentation_id=presentation_id,
        slide_id=slide_id,
        equation_id=equation_id,
    )


@mcp.tool(annotations=EDIT)
async def ppt_save_presentation(
    ctx: Context, presentation_id: str, path: str | None = None, overwrite: bool = False
) -> dict[str, Any]:
    """Explicitly save a PowerPoint presentation or save a new file as .pptx.

    A new presentation needs an absolute local .pptx path. With no path, save to
    its existing path. Overwriting a different file requires overwrite=True.
    """
    return await _call(
        ctx,
        "ppt_save_presentation",
        presentation_id=presentation_id,
        path=path,
        overwrite=overwrite,
    )


def main() -> None:
    """Run local stdio transport; stdout is reserved for MCP protocol messages."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

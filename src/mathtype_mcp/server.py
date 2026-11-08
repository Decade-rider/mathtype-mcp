"""Stdio MCP service for Windows Word native MathType equations."""

from contextlib import asynccontextmanager
from typing import Literal

from mcp.server.fastmcp import Context, FastMCP
from mcp.types import ToolAnnotations

from mathtype_mcp.sta import StaExecutor


@asynccontextmanager
async def lifespan(server):
    worker = StaExecutor()
    try:
        await worker.start()
        yield {"worker": worker}
    finally:
        await worker.close()


mcp = FastMCP(
    "MathType Word",
    lifespan=lifespan,
    instructions=(
        "Use desktop Windows Word and MathType. First mathtype_health then word_list_documents. "
        "Use returned document/equation IDs, never guess them. Supports Word main text only, "
        "inline/display/numbered native MathType OLE. No PowerPoint or custom internal size. "
        "Inspect selection before replacement. Never operate Office or another equation MCP "
        "concurrently with this server. Conversions use temporary Word documents and clipboard. "
        "Changes are unsaved until explicitly requested. After interruption inspect the document "
        "before retrying a mutation; a cancelled request can still complete. "
        "Export is read from the live OLE copy, not cached source. Review rendered mathematics."
    ),
)
READ = ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=False)
WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=False)
EDIT = ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=False)


async def call(ctx, method, **kwargs):
    return await ctx.request_context.lifespan_context["worker"].call(method, **kwargs)


@mcp.tool(annotations=READ)
async def mathtype_health(ctx: Context) -> dict:
    """Inspect installed Word/MathType integration without starting Office."""
    return await call(ctx, "health")


@mcp.tool(annotations=READ)
async def word_list_documents(ctx: Context) -> dict:
    """List reachable Word documents. Use returned document_id in subsequent calls."""
    return await call(ctx, "list_documents")


@mcp.tool(annotations=WRITE)
async def word_create_document(ctx: Context, visible: bool = True) -> dict:
    """Create an unsaved Word document."""
    return await call(ctx, "create_document", visible=visible)


@mcp.tool(annotations=WRITE)
async def word_open_document(
    ctx: Context, path: str, read_only: bool = False, visible: bool = True
) -> dict:
    """Open an absolute local Word path; no automatic save."""
    return await call(ctx, "open_document", path=path, read_only=read_only, visible=visible)


@mcp.tool(annotations=READ)
async def word_document_info(ctx: Context, document_id: str) -> dict:
    """Inspect bookmarks and the exact current selection before editing."""
    return await call(ctx, "get_document_info", document_id=document_id)


@mcp.tool(annotations=READ)
async def word_list_equations(ctx: Context, document_id: str) -> dict:
    """List native MathType equations in main text, excluding AxMath and Word OMML."""
    return await call(ctx, "list_equations", document_id=document_id)


@mcp.tool(annotations=EDIT)
async def word_insert_mathtype(
    ctx: Context,
    document_id: str,
    latex: str,
    position: Literal["end", "bookmark", "selection"] = "end",
    bookmark: str | None = None,
    expected_selection: str | None = None,
    layout: Literal["inline", "display", "numbered"] = "inline",
) -> dict:
    """Insert one native MathType OLE. Bare math or one delimiter pair is accepted.

    Selection replacement requires exact expected_selection from word_document_info.
    Bookmark insertion is before its text. Display uses a dedicated centered paragraph.
    Numbered adds a right-aligned native MathType automatic number (1), (2), ... .
    Internal font settings come from MathType defaults. Never saves automatically.
    """
    return await call(
        ctx,
        "insert_equation",
        document_id=document_id,
        latex=latex,
        position=position,
        bookmark=bookmark,
        expected_selection=expected_selection,
        layout=layout,
    )


@mcp.tool(annotations=READ)
async def word_get_equation_latex(ctx: Context, document_id: str, equation_id: str) -> dict:
    """Export TeX from a disposable copy of the live MathType object; preserve the source."""
    return await call(ctx, "get_equation_latex", document_id=document_id, equation_id=equation_id)


@mcp.tool(annotations=EDIT)
async def word_update_mathtype(
    ctx: Context, document_id: str, equation_id: str, latex: str
) -> dict:
    """Replace one MathType formula with new TeX, preserving surrounding text and layout.

    Equation ID must come from this document's list. Uses MathType default internal font
    preferences; custom per-equation styling is not guaranteed. Native numbered paragraphs
    retain their numbers. Changes remain unsaved.
    """
    return await call(
        ctx, "update_equation", document_id=document_id, equation_id=equation_id, latex=latex
    )


@mcp.tool(annotations=EDIT)
async def word_save_document(
    ctx: Context, document_id: str, path: str | None = None, overwrite: bool = False
) -> dict:
    """Explicitly save. New documents require an absolute .docx path.

    Overwriting a different existing file requires explicit user authorization.
    """
    return await call(ctx, "save_document", document_id=document_id, path=path, overwrite=overwrite)


@mcp.tool(annotations=EDIT)
async def word_number_mathtype(ctx: Context, document_id: str, equation_id: str) -> dict:
    """Add a native automatic (1), (2), ... number to an equation in a dedicated paragraph.

    Uses installed MathType macros. Existing numbered equations are unchanged.
    Mixed text paragraphs and chapter/custom number formats are rejected. No auto-save.
    """
    return await call(ctx, "number_equation", document_id=document_id, equation_id=equation_id)


@mcp.tool(annotations=EDIT)
async def word_refresh_equation_numbers(ctx: Context, document_id: str) -> dict:
    """Refresh main-text MathType automatic numbers after insertion, deletion or movement.

    Uses the vendor macro on number fields only. Does not refresh cross-references,
    unrelated fields or save the document. Locked number fields are rejected.
    """
    return await call(ctx, "refresh_equation_numbers", document_id=document_id)


def main():
    mcp.run()


if __name__ == "__main__":
    main()

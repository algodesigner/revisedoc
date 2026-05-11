"""
MCP server providing revisedoc operations as structured tools.

Start with:  revisedoc-mcp
"""

import asyncio
import json

from docx import Document
from mcp.server import Server
from mcp.server.stdio import stdio_server

from revisedoc.editor import (
    add_comment,
    finalize_comments,
    get_full_text,
    list_revisions,
    replace_text,
    restore_revision,
)

server = Server("revisedoc")


@server.tool(description="Replace all occurrences of old_text with new_text using tracked changes")
async def docx_replace(
    input_path: str,
    output_path: str,
    old_text: str,
    new_text: str,
    author: str = "Editor",
) -> str:
    doc = Document(input_path)
    replace_text(doc, old_text, new_text, author=author)
    doc.save(output_path)
    return f"Replaced {old_text!r} with {new_text!r} -> {output_path}"


@server.tool(description="Add a comment anchored to a specific text range in the document")
async def docx_comment(
    input_path: str,
    output_path: str,
    target_text: str,
    comment: str,
    author: str = "Editor",
) -> str:
    doc = Document(input_path)
    add_comment(doc, target_text, comment, author=author)
    pending = getattr(doc.part, "_pending_comments", None)
    doc.save(output_path)
    finalize_comments(output_path, pending_comments=pending)
    return f"Added comment on {target_text!r} -> {output_path}"


@server.tool(description="List all tracked changes (insertions and deletions) in a document")
async def docx_list_revisions(
    input_path: str,
    format: str = "text",
) -> str:
    doc = Document(input_path)
    revs = list_revisions(doc)
    if format == "json":
        return json.dumps(revs, indent=2)
    if not revs:
        return "No revisions found."
    lines = []
    for r in revs:
        lines.append(f"ID: {r['id']} | Type: {r['type']} | Author: {r['author']}")
        lines.append(f"  Text: {r['text'][:80]}")
        lines.append(f"  In paragraph: {r['paragraph_text'][:80]}")
        lines.append("")
    return "\n".join(lines)


@server.tool(description="Undo a specific revision by its ID")
async def docx_restore(
    input_path: str,
    output_path: str,
    revision_id: str,
    restore_type: str = "deletion",
) -> str:
    doc = Document(input_path)
    restore_revision(doc, revision_id, restore_type=restore_type)
    doc.save(output_path)
    return f"Restored revision {revision_id} -> {output_path}"


@server.tool(description="Print the document plain text (insertions applied, deletions excluded)")
async def docx_get_text(
    input_path: str,
) -> str:
    doc = Document(input_path)
    return get_full_text(doc)


def main():
    asyncio.run(_run())


async def _run():
    async with stdio_server() as (read, write):
        await server.run(read, write)


if __name__ == "__main__":
    main()

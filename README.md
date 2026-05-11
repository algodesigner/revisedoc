# revisedoc

Edit `.docx` files the way Word intended — with full change tracking, inline
comments, and revision history. No more sending around unmarked copies or
losing track of what changed. Use it from the command line, from your Python
code, or hook it up as an MCP server so your AI coding agent can edit documents
with the same care a human reviewer would.

## Features

- **Tracked-changes replace** — Replace text using Word-compatible `<w:del>` / `<w:ins>` markup
- **Inline comments** — Anchor comments to any text range
- **Revision inspection** — List all tracked changes with id, author, date, type, text
- **Revision restoration** — Undo a specific tracked change
- **MCP server** — Expose all operations as structured tools for AI agents (opencode, Claude Code, etc.)

## Use Cases

- **Contract redlining** — Replace outdated terms in a `.docx` contract with
  tracked changes so the other side sees exactly what was modified.

- **Document review workflows** — Annotate drafts with inline comments,
  then list and restore revisions to review changes systematically.

- **AI-assisted editing** — Connect the MCP server to opencode or Claude Code
  and have your agent propose edits with full change tracking — no more
  "just trust me, I made some changes."

- **Batch processing** — Script document updates across hundreds of files
  using the Python API, with every change recorded in Word's revision log.

## Quick Start

Create a document, then use the MCP tools from your AI agent:

```
1. docx_inspect     — see the paragraph structure and formatting
2. docx_search      — locate exact text before editing
3. docx_replace     — make tracked changes (strikethrough + underline)
4. docx_comment     — annotate clauses with inline comments
5. docx_list_revisions — review what changed
```

Every operation produces standard Word revision markup — open the output in
Word to see tracked changes, or accept/reject them as usual.

## Installation

```bash
pip install revisedoc          # with pip
pipx install revisedoc          # with pipx (isolated)
```

For AI agent integration (opencode, Claude Code):

```bash
pip install "revisedoc[mcp]"     # pip — quotes required in zsh
pipx install "revisedoc[mcp]"    # pipx — quotes required in zsh
```

> `[mcp]` is an optional extra — it installs the same `revisedoc` package plus the
> `mcp` SDK dependency. Without it, the CLI and Python API work fine; only the
> `revisedoc-mcp` server is unavailable.
>
> **Note for zsh users:** the square brackets must be quoted. Use
> `pip install "revisedoc[mcp]"` not `pip install revisedoc[mcp]`.

## CLI Usage

```bash
# Replace text with tracked changes
revisedoc replace input.docx output.docx "old text" "new text" --author "Me"

# Add a comment
revisedoc comment input.docx output.docx "target text" "my comment" --author "Reviewer"

# List tracked changes
revisedoc list-revisions input.docx
revisedoc list-revisions input.docx --format json

# Undo a revision (restore deleted text)
revisedoc restore input.docx output.docx --revision-id 3

# Undo an insertion instead
revisedoc restore input.docx output.docx --revision-id 3 --restore-type insertion

# Print document plain text
revisedoc get-text input.docx
```

## Python API

```python
from docx import Document
from revisedoc import replace_text, add_comment, list_revisions, restore_revision, get_full_text, finalize_comments

doc = Document("input.docx")
replace_text(doc, "old phrase", "new phrase", author="MyBot")
add_comment(doc, "some specific text", "Review this part", author="Reviewer")
pending = doc.part._pending_comments
doc.save("output.docx")
finalize_comments("output.docx", pending_comments=pending)
```

## MCP Server — AI Agent Integration

Start the MCP server:

```bash
revisedoc-mcp
```

### opencode

Add to `opencode.json`:

```json
{
  "mcp": {
    "revisedoc": {
      "type": "local",
      "command": ["revisedoc-mcp"],
      "enabled": true
    }
  }
}
```

### Claude Code

Add to `.claude/settings.json`:

```json
{
  "mcpServers": {
    "revisedoc": {
      "command": "revisedoc-mcp"
    }
  }
}
```

### Available Tools

| Tool | Description |
|---|---|
| `docx_replace` | Replace text with tracked changes |
| `docx_comment` | Add a comment anchored to text |
| `docx_list_revisions` | List all tracked changes |
| `docx_restore` | Undo a revision by ID |
| `docx_get_text` | Print document plain text |

## Development

```bash
pip install -e ".[mcp,test]"
pytest tests/ -v
```

## Architecture

The package structure:

```
src/revisedoc/
├── __init__.py     # Public API exports
├── editor.py       # Core editing operations
├── cli.py          # Command-line interface
└── mcp.py          # MCP server for AI agents
```

# revisedoc

Edit `.docx` files with tracked changes, comments, and revision restoration — for humans and AI agents.

## Features

- **Tracked-changes replace** — Replace text using Word-compatible `<w:del>` / `<w:ins>` markup
- **Inline comments** — Anchor comments to any text range
- **Revision inspection** — List all tracked changes with id, author, date, type, text
- **Revision restoration** — Undo a specific tracked change
- **MCP server** — Expose all operations as structured tools for AI agents (opencode, Claude Code, etc.)

## Installation

```bash
pip install revisedoc
```

For MCP support:

```bash
pip install "revisedoc[mcp]"
```

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

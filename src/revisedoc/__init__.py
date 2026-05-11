"""
revisedoc — Edit .docx files with tracked changes, comments, and revision restoration.
"""

from revisedoc.editor import (
    add_comment,
    finalize_comments,
    get_full_text,
    inspect_document,
    list_revisions,
    replace_text,
    restore_revision,
    search_text,
)

__all__ = [
    "add_comment",
    "finalize_comments",
    "get_full_text",
    "inspect_document",
    "list_revisions",
    "replace_text",
    "restore_revision",
    "search_text",
]

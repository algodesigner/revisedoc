"""
revisedoc — Edit .docx files with tracked changes, comments, and revision restoration.
"""

from revisedoc.editor import (
    add_comment,
    finalize_comments,
    get_full_text,
    list_revisions,
    replace_text,
    restore_revision,
)

__all__ = [
    "add_comment",
    "finalize_comments",
    "get_full_text",
    "list_revisions",
    "replace_text",
    "restore_revision",
]

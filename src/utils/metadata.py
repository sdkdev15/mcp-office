"""Shared document metadata utilities."""

from __future__ import annotations

from typing import Any


def apply_metadata(props: Any, metadata: dict) -> None:
    """Apply metadata dict to a document properties object.

    Works with openpyxl, python-docx, and python-pptx core_properties
    since they all expose the same attribute names.

    Args:
        props: Document properties object (e.g., wb.properties, doc.core_properties).
        metadata: Metadata dictionary with keys: author, company, subject, title,
                  keywords, category, comments.
    """
    _FIELD_MAP = {
        "author": ("author", "creator"),  # docx/pptx use 'author', openpyxl uses 'creator'
        "company": ("company",),
        "subject": ("subject",),
        "title": ("title",),
        "keywords": ("keywords",),
        "category": ("category",),
        "comments": ("description",),
    }

    for key, value in metadata.items():
        if not value:
            continue
        attr_names = _FIELD_MAP.get(key)
        if not attr_names:
            continue
        for attr in attr_names:
            if hasattr(props, attr):
                setattr(props, attr, value)
                break

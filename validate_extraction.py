"""Validation module for pdf2excel extraction results.

Checks extracted items for various potential quality issues by delegating to validation.py.
"""

from __future__ import annotations

from typing import Any, Dict, List, Set, Optional
import validation


def validate_items(items: List[Dict[str, Any]], boilerplate_set: Optional[Set[str]] = None) -> List[Dict[str, Any]]:
    """Validate a list of mapped assessment items, adding confidence and issues."""
    return validation.validate_items(items, boilerplate_set=boilerplate_set)

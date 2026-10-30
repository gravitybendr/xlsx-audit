"""xlsx-audit: find broken references, external links, hard-coded overrides
and formula changes in Excel workbooks.

    >>> import xlsx_audit
    >>> findings = xlsx_audit.audit("model.xlsx")
    >>> changes = xlsx_audit.compare("model_v1.xlsx", "model_v2.xlsx")
"""
from __future__ import annotations

__version__ = "0.1.0"

from pathlib import Path  # noqa: E402

from .checks import lint  # noqa: E402
from .diff import DiffResult, diff  # noqa: E402
from .model import DIFF_RULES, LINT_RULES, Finding  # noqa: E402
from .workbook import Workbook, WorkbookError, load  # noqa: E402


def audit(path: str | Path, select: set[str] | None = None, ignore: set[str] | None = None) -> list[Finding]:
    """Lint one workbook file and return its findings, most severe first."""
    return lint(load(path), select, ignore)


def compare(old: str | Path, new: str | Path, include_values: bool = True) -> list[Finding]:
    """Changes between two versions of a workbook, plus new and resolved lint findings."""
    return diff(load(old), load(new), include_values=include_values).findings


__all__ = [
    "__version__",
    "audit",
    "compare",
    "lint",
    "diff",
    "load",
    "Finding",
    "DiffResult",
    "Workbook",
    "WorkbookError",
    "LINT_RULES",
    "DIFF_RULES",
]

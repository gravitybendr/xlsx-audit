"""Findings, severities and the rule catalogue."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field

SEVERITIES = ("info", "warning", "error")


def severity_rank(severity: str) -> int:
    return SEVERITIES.index(severity)


@dataclass(frozen=True)
class Rule:
    code: str
    slug: str
    severity: str
    summary: str


# Lint rules (one workbook).
LINT_RULES = {
    r.code: r
    for r in [
        Rule("XA001", "error-value", "error", "A cell shows an Excel error such as #REF!, #NAME? or #DIV/0!."),
        Rule("XA002", "broken-reference", "error", "A formula contains #REF!, a reference to cells that were deleted."),
        Rule("XA003", "external-link", "warning", "The workbook or a formula depends on another workbook file."),
        Rule("XA004", "hardcoded-in-formula-range", "warning", "A typed number sits inside a run of cells that all use the same formula."),
        Rule("XA005", "inconsistent-formula", "warning", "A formula breaks the pattern of the identical formulas around it."),
        Rule("XA006", "volatile-function", "info", "A formula uses INDIRECT, OFFSET, NOW, TODAY, RAND or another volatile function."),
        Rule("XA007", "hidden-sheet", "info", "A sheet is hidden, or very hidden (only visible through VBA)."),
        Rule("XA008", "broken-name", "error", "A defined name points to #REF!."),
        Rule("XA009", "hardcoded-number", "info", "A formula contains a typed number, such as =B2*1.07."),
        Rule("XA010", "self-reference", "error", "A formula refers to its own cell, which makes a circular reference."),
    ]
}

# Changes between two versions of a workbook.
DIFF_RULES = {
    r.code: r
    for r in [
        Rule("XD001", "sheet-added", "info", "A sheet was added."),
        Rule("XD002", "sheet-removed", "warning", "A sheet was removed."),
        Rule("XD003", "sheet-renamed", "info", "A sheet was renamed."),
        Rule("XD010", "rows-inserted", "info", "Rows were inserted."),
        Rule("XD011", "rows-deleted", "info", "Rows were deleted."),
        Rule("XD012", "columns-inserted", "info", "Columns were inserted."),
        Rule("XD013", "columns-deleted", "info", "Columns were deleted."),
        Rule("XD020", "formula-changed", "info", "A formula was changed."),
        Rule("XD021", "formula-overwritten", "warning", "A formula was replaced by a typed value."),
        Rule("XD022", "formula-added", "info", "A formula was added where there was a value or nothing."),
        Rule("XD023", "formula-removed", "warning", "A formula was deleted, leaving the cell empty."),
        Rule("XD024", "value-changed", "info", "A typed value was changed."),
        Rule("XD025", "value-added", "info", "A typed value was added to an empty cell."),
        Rule("XD026", "value-removed", "info", "A typed value was deleted."),
        Rule("XD030", "name-changed", "info", "A defined name was added, removed or repointed."),
        Rule("XD031", "external-link-added", "warning", "The new version links to another workbook file."),
        Rule("XD032", "external-link-removed", "info", "A link to another workbook file was removed."),
        Rule("XD040", "new-issue", "warning", "A lint finding that the old version didn't have."),
        Rule("XD041", "resolved-issue", "info", "A lint finding from the old version that is gone."),
    ]
}

ALL_RULES = {**LINT_RULES, **DIFF_RULES}


@dataclass
class Finding:
    """One lint finding or one change.

    cell is an A1 reference (e.g. "C7"), or None for sheet- and
    workbook-level findings. For changes, old_cell is the cell's address in
    the old version when rows or columns moved.
    """

    code: str
    severity: str
    message: str
    sheet: str | None = None
    cell: str | None = None
    old_cell: str | None = None
    old: str | None = None
    new: str | None = None
    detail: dict = field(default_factory=dict)

    @property
    def slug(self) -> str:
        return ALL_RULES[self.code].slug

    @property
    def location(self) -> str:
        cell = self.cell or self.old_cell
        if self.sheet and cell:
            return f"{quote_sheet(self.sheet)}!{cell}"
        return quote_sheet(self.sheet) if self.sheet else "(workbook)"

    def to_dict(self) -> dict:
        d = asdict(self)
        d["rule"] = self.slug
        return {k: v for k, v in d.items() if v not in (None, {}, "")}


def quote_sheet(name: str) -> str:
    """Quote a sheet name the way Excel does in references."""
    if name.replace("_", "").replace(".", "").isalnum() and not name[0].isdigit():
        return name
    return "'" + name.replace("'", "''") + "'"


def sort_findings(findings: list[Finding], sheet_order: list[str]) -> list[Finding]:
    from .refs import parse_cell  # local import to avoid a cycle

    order = {name: i for i, name in enumerate(sheet_order)}

    def key(f: Finding):
        row, col = parse_cell(f.cell) if f.cell else (0, 0)
        return (-severity_rank(f.severity), order.get(f.sheet, -1), row, col, f.code)

    return sorted(findings, key=key)

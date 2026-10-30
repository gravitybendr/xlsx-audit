"""Lint rules for one workbook (XA001-XA010)."""
from __future__ import annotations

from bisect import bisect_left
from collections import Counter

from openpyxl.formula.translate import Translator

from .model import LINT_RULES, Finding, sort_findings
from .refs import (
    VOLATILE,
    cell_name,
    functions,
    has_ref_error,
    is_benign_number,
    normalize,
    numbers,
    references,
    split_sheet,
)
from .workbook import Cell, Sheet, Workbook

ERROR_SEVERITY = {
    "#REF!": "error",
    "#NAME?": "error",
    "#NULL!": "error",
    "#N/A": "info",
}

Pos = tuple[int, int]


def _fmt(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _translate(formula: str, origin: str, target: str) -> str:
    """The formula at origin, copied to target (A1 text with "=")."""
    try:
        return Translator("=" + formula, origin=origin).translate_formula(target)
    except Exception:
        return "=" + formula


def _link_label(book: str, links: list[str]) -> str:
    if book.isdigit() and 0 < int(book) <= len(links):
        return links[int(book) - 1]
    return book


def _safe_refs(text: str):
    try:
        return references(text)
    except Exception:
        return []


def _runs(indexes: list[int]) -> list[list[int]]:
    """[1,2,3,7,8] -> [[1,2,3],[7,8]]."""
    runs: list[list[int]] = []
    for i in sorted(indexes):
        if runs and i == runs[-1][-1] + 1:
            runs[-1].append(i)
        else:
            runs.append([i])
    return runs


class _SheetContext:
    def __init__(self, sheet: Sheet):
        self.sheet = sheet
        self.cells = sheet.cells
        self.norm: dict[Pos, str] = {
            pos: normalize(c.formula, pos[0], pos[1]) for pos, c in sheet.cells.items() if c.formula is not None
        }
        self.pattern: dict[Pos, Finding] = {}
        self._find_pattern_breaks()

    def same_pattern(self, a: Pos, b: Pos) -> bool:
        return a in self.norm and b in self.norm and self.norm[a] == self.norm[b]

    def _expected(self, source: Pos, here: Pos) -> str:
        return _translate(self.cells[source].formula, cell_name(*source), cell_name(*here))

    def _find_pattern_breaks(self) -> None:
        """XA004 and XA005 from runs of cells down columns and across rows.

        In each unbroken run, the formula most cells share (at least three
        copies) is the pattern. Inside the stretch it covers, a typed number
        is flagged (XA004), and so is a formula that appears at most twice
        in the run (XA005): a rare variant is likely a slip, while a common
        one is likely a deliberate second block.
        """
        name = self.sheet.name
        for axis in (0, 1):
            lines: dict[int, list[int]] = {}
            for r, c in self.cells:
                lines.setdefault(c if axis == 0 else r, []).append(r if axis == 0 else c)
            for line, idxs in lines.items():
                for run in _runs(idxs):
                    if len(run) < 3:
                        continue
                    positions = [(i, line) if axis == 0 else (line, i) for i in run]
                    pats = Counter(self.norm[p] for p in positions if p in self.norm)
                    if not pats:
                        continue
                    dom, count = pats.most_common(1)[0]
                    if count < 3:
                        continue
                    dom_at = [k for k, p in enumerate(positions) if self.norm.get(p) == dom]

                    def nearest(k: int, dom_at=dom_at, positions=positions) -> Pos:
                        i = bisect_left(dom_at, k)
                        near = [d for d in (dom_at[i - 1] if i else None, dom_at[i] if i < len(dom_at) else None) if d is not None]
                        return positions[min(near, key=lambda d: (abs(d - k), d))]

                    for k in range(dom_at[0] + 1, dom_at[-1]):
                        p = positions[k]
                        if p in self.pattern or self.norm.get(p) == dom:
                            continue
                        cell = self.cells[p]
                        if p in self.norm and pats[self.norm[p]] <= 2:
                            src = nearest(k)
                            expected = self._expected(src, p)
                            self.pattern[p] = Finding(
                                "XA005", "warning",
                                f"={cell.formula} breaks the pattern of the {count} matching formulas around it (expected {expected})",
                                sheet=name, cell=cell_name(*p), detail={"expected": expected, "like": cell_name(*src)},
                            )
                        elif p not in self.norm and cell.is_number:
                            src = nearest(k)
                            expected = self._expected(src, p)
                            self.pattern[p] = Finding(
                                "XA004", "warning",
                                f"Typed value {_fmt(cell.value)} inside a run of {count} matching formulas (expected {expected})",
                                sheet=name, cell=cell_name(*p), detail={"expected": expected, "like": cell_name(*src)},
                            )
                    # A typed number right after the run ends it: a common overwrite.
                    last = dom_at[-1]
                    if last + 1 == len(positions) - 1:
                        p = positions[last + 1]
                        cell = self.cells[p]
                        if p not in self.pattern and p not in self.norm and cell.is_number:
                            expected = self._expected(positions[last], p)
                            self.pattern[p] = Finding(
                                "XA004", "info",
                                f"Typed value {_fmt(cell.value)} at the end of a run of {count} matching formulas (expected {expected})",
                                sheet=name, cell=cell_name(*p), detail={"expected": expected, "like": cell_name(*positions[last])},
                            )


def lint(wb: Workbook, select: set[str] | None = None, ignore: set[str] | None = None) -> list[Finding]:
    """Run the lint rules. select/ignore take rule codes (XA001) or names (error-value)."""
    selected = _resolve(select) if select else set(LINT_RULES)
    ignored = _resolve(ignore) if ignore else set()
    findings: list[Finding] = []

    for i, target in enumerate(wb.external_links, 1):
        findings.append(Finding("XA003", "warning", f"Workbook links to another file: {target}", detail={"link": i, "target": target}))

    for dn in wb.names:
        if dn.name.lower() == "_xlnm._filterdatabase":
            continue
        where = f" (on sheet {dn.scope})" if dn.scope else ""
        if has_ref_error(dn.text):
            findings.append(Finding("XA008", "error", f"Defined name {dn.name}{where} points to {dn.text}", sheet=dn.scope, detail={"name": dn.name}))
        for ref in _safe_refs(dn.text):
            if ref.book:
                findings.append(
                    Finding("XA003", "warning", f"Defined name {dn.name}{where} points to another workbook: {_link_label(ref.book, wb.external_links)}", sheet=dn.scope, detail={"name": dn.name})
                )
                break

    error_cells = {s.name.lower(): {pos for pos, c in s.cells.items() if c.error} for s in wb.sheets}
    for sheet in wb.sheets:
        if sheet.state in ("hidden", "veryHidden"):
            very = sheet.state == "veryHidden"
            msg = f"Sheet {sheet.name} is very hidden: it can only be shown with VBA" if very else f"Sheet {sheet.name} is hidden"
            findings.append(Finding("XA007", "warning" if very else "info", msg, sheet=sheet.name, detail={"state": sheet.state}))
        ctx = _SheetContext(sheet)
        for _pos, cell in sorted(sheet.cells.items()):
            findings.extend(_cell_findings(wb, ctx, cell, error_cells))

    findings = [f for f in findings if f.code in selected and f.code not in ignored]
    return sort_findings(findings, [s.name for s in wb.sheets])


def _resolve(items: set[str]) -> set[str]:
    by_slug = {r.slug: code for code, r in LINT_RULES.items()}
    out = set()
    for item in items:
        key = item.strip()
        if key.upper() in LINT_RULES:
            out.add(key.upper())
        elif key.lower() in by_slug:
            out.add(by_slug[key.lower()])
        else:
            raise ValueError(f"unknown rule: {item} (see: xlsx-audit rules)")
    return out


def _inherited_from(cell: Cell, sheet: Sheet, error_cells: dict[str, set[Pos]]) -> str | None:
    """A cell this error comes from, if the formula reads another error cell."""
    for ref in _safe_refs(cell.formula or ""):
        if ref.book or ref.area is None:
            continue
        target = (ref.sheet or sheet.name).lower()
        r1, c1, r2, c2 = ref.area
        for r, c in sorted(error_cells.get(target, ())):
            if (r, c) == (cell.row, cell.col) and target == sheet.name.lower():
                continue
            if r1 <= r <= r2 and c1 <= c <= c2:
                prefix = split_sheet(ref.text)[0]
                return f"{prefix}{cell_name(r, c)}"
    return None


def _cell_findings(wb: Workbook, ctx: _SheetContext, cell: Cell, error_cells) -> list[Finding]:
    out: list[Finding] = []
    sheet = ctx.sheet
    row, col = cell.row, cell.col
    here = cell.name

    def add(code: str, message: str, severity: str | None = None, **detail):
        out.append(Finding(code, severity or LINT_RULES[code].severity, message, sheet=sheet.name, cell=here, detail=detail))

    formula = cell.formula
    ref_error = formula is not None and has_ref_error(formula)

    if cell.error and not (ref_error and cell.error == "#REF!"):
        if formula is None:
            add("XA001", f"Shows {cell.error} as a typed or pasted value", "warning", error=cell.error)
        else:
            source = _inherited_from(cell, sheet, error_cells)
            if source:
                add("XA001", f"Shows {cell.error} because it uses {source}, which also shows an error", "info", error=cell.error, source=source)
            else:
                add("XA001", f"Shows {cell.error}: ={formula}", ERROR_SEVERITY.get(cell.error, "warning"), error=cell.error)

    if (row, col) in ctx.pattern:
        out.append(ctx.pattern[(row, col)])

    if formula is not None:
        if ref_error:
            add("XA002", f"Formula refers to deleted cells: ={formula}")

        refs = _safe_refs(formula)
        books = []
        for ref in refs:
            if ref.book:
                label = _link_label(ref.book, wb.external_links)
                if label not in books:
                    books.append(label)
        if books:
            add("XA003", f"Formula uses another workbook ({', '.join(books)}): ={formula}", targets=books)

        for ref in refs:
            if ref.book or ref.area is None:
                continue
            if ref.sheet is not None and ref.sheet.lower() != sheet.name.lower():
                continue
            r1, c1, r2, c2 = ref.area
            if r1 <= row <= r2 and c1 <= col <= c2:
                add("XA010", f"Formula includes its own cell ({ref.text}), a circular reference: ={formula}", ref=ref.text)
                break

        volatile = [f for f in dict.fromkeys(functions(formula)) if f in VOLATILE]
        if volatile:
            hides = " and hides what it reads" if {"INDIRECT", "OFFSET"} & set(volatile) else ""
            add("XA006", f"Uses {', '.join(volatile)}, which recalculates on every change{hides}", functions=volatile)

        typed = [n for n in dict.fromkeys(numbers(formula)) if not is_benign_number(n)]
        if typed:
            add("XA009", f"Typed number{'s' if len(typed) > 1 else ''} {', '.join(typed)} inside ={formula}", numbers=typed)

        if (row, col) not in ctx.pattern:
            mine = ctx.norm.get((row, col))
            for a, b in (((row - 1, col), (row + 1, col)), ((row, col - 1), (row, col + 1))):
                if ctx.same_pattern(a, b) and ctx.norm[a] != mine:
                    expected = _translate(ctx.cells[a].formula, cell_name(*a), here)
                    add("XA005", f"={formula} differs from {cell_name(*a)} and {cell_name(*b)} on either side (expected {expected})",
                        expected=expected, like=cell_name(*a))
                    break

    elif cell.is_number and (row, col) not in ctx.pattern:
        for a, b in (((row - 1, col), (row + 1, col)), ((row, col - 1), (row, col + 1))):
            if ctx.same_pattern(a, b):
                expected = _translate(ctx.cells[a].formula, cell_name(*a), here)
                add("XA004", f"Typed value {_fmt(cell.value)} where {cell_name(*a)} and {cell_name(*b)} on either side use a formula (expected {expected})",
                    expected=expected, like=cell_name(*a))
                break
    return out

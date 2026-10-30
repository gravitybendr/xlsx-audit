"""Compare two versions of a workbook (XD rules).

Two phases:

1. Line up each sheet's rows (or columns, whichever explains the changes
   better), so inserting a row doesn't make every cell below it look changed.
   Rows that changed are matched by similarity, not just by order.
2. Compare cells. An old formula is first rewritten with every reference
   moved to where its cells are in the new version, on any sheet, so
   =SUM(C2:C8) that became =SUM(C3:C9) after a row was inserted above it
   counts as unchanged. Renamed sheets are matched by content.
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from .checks import lint
from .model import DIFF_RULES, Finding, severity_rank, sort_findings
from .refs import canonical, cell_name, normalize, parse_cell, remap, shape
from .workbook import Cell, Sheet, Workbook

Pos = tuple[int, int]
Maps = dict[str, tuple[str, int, dict[int, int]]]  # old sheet (lower) -> (new name, axis, line map)

MAX_REFINE = 40_000  # cells in a changed block above which rows are paired in order


@dataclass
class DiffResult:
    findings: list[Finding]
    renamed: dict[str, str] = field(default_factory=dict)  # old -> new
    stats: dict = field(default_factory=dict)


@dataclass
class _Plan:
    old: Sheet
    new: Sheet
    axis: int  # 0 rows, 1 columns
    pairs: list[tuple[int | None, int | None]]  # 1-based line numbers

    @property
    def line_map(self) -> dict[int, int]:
        return {i: j for i, j in self.pairs if i is not None and j is not None}


def _value_key(v) -> tuple:
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v))
    if isinstance(v, (datetime.date, datetime.time, datetime.datetime)):
        return ("d", v.isoformat())
    if isinstance(v, datetime.timedelta):
        return ("t", v.total_seconds())
    return ("s", str(v))


def _key(cell: Cell, alias: dict[str, str] | None) -> tuple:
    """Coarse content key for lining up rows and columns (see refs.shape)."""
    if cell.formula is not None:
        return ("f", shape(cell.formula, alias))
    if cell.error and cell.value is None:
        return ("e", cell.error)
    return ("v",) + _value_key(cell.value)


def _grid(sheet: Sheet, alias: dict[str, str] | None = None) -> dict[Pos, tuple]:
    return {pos: _key(c, alias) for pos, c in sheet.cells.items()}


def _lines(grid: dict[Pos, tuple], axis: int, count: int) -> list[tuple]:
    """Contents of each row (axis 0) or column (axis 1), 1..count."""
    buckets: dict[int, list] = {}
    for pos, key in grid.items():
        buckets.setdefault(pos[axis], []).append((pos[1 - axis], key))
    return [tuple(sorted(buckets.get(i, []))) for i in range(1, count + 1)]


def _similarity(a: tuple, b: tuple) -> float:
    if not a and not b:
        return 1.0
    da, db = dict(a), dict(b)
    same = sum(1 for k, v in da.items() if db.get(k) == v)
    return same / max(len(da), len(db))


def _refine(old: list[tuple], new: list[tuple], i1: int, i2: int, j1: int, j2: int) -> list[tuple[int | None, int | None]]:
    """Pair the lines of a changed block.

    Equal-sized blocks were edited in place, so they pair in order. Otherwise
    every line on the shorter side is paired, and the extra lines on the
    longer side (the inserted or deleted ones) are chosen to make the pairs
    as similar as possible. Ties put the extra lines last.
    """
    n, m = i2 - i1, j2 - j1
    if n == m or n * m > MAX_REFINE:
        k = min(n, m)
        return [(i1 + t, j1 + t) for t in range(k)] + [(i, None) for i in range(i1 + k, i2)] + [(None, j) for j in range(j1 + k, j2)]
    # Make the old side the longer one; swap back at the end.
    swap = m > n
    a_lines = [new[j1 + t] for t in range(m)] if swap else [old[i1 + t] for t in range(n)]
    b_lines = [old[i1 + t] for t in range(n)] if swap else [new[j1 + t] for t in range(m)]
    la, lb = len(a_lines), len(b_lines)
    neg = float("-inf")
    # score[x][y]: best total pairing the first y b-lines within the first x a-lines.
    score = [[neg] * (lb + 1) for _ in range(la + 1)]
    for x in range(la + 1):
        score[x][0] = 0.0
    for x in range(1, la + 1):
        for y in range(1, min(x, lb) + 1):
            pair = score[x - 1][y - 1] + _similarity(a_lines[x - 1], b_lines[y - 1])
            skip = score[x - 1][y] if x - 1 >= y else neg
            score[x][y] = max(pair, skip)
    pairs: list[tuple[int | None, int | None]] = []
    x, y = la, lb
    while x > 0:
        skip = score[x - 1][y] if x - 1 >= y else neg
        if y > 0 and (skip < score[x][y] - 1e-12 or skip == neg):
            pairs.append((x - 1, y - 1))
            y -= 1
        else:
            pairs.append((x - 1, None))
        x -= 1
    pairs.reverse()
    out: list[tuple[int | None, int | None]] = []
    for xa, yb in pairs:
        if swap:
            out.append((None if yb is None else i1 + yb, j1 + xa))
        else:
            out.append((i1 + xa, None if yb is None else j1 + yb))
    return out


def _align(old: list[tuple], new: list[tuple]) -> list[tuple[int | None, int | None]]:
    """Pair line numbers (1-based) of two sheets."""
    pairs: list[tuple[int | None, int | None]] = []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, old, new, autojunk=False).get_opcodes():
        if tag == "equal":
            pairs.extend((i1 + k, j1 + k) for k in range(i2 - i1))
        elif tag == "replace":
            pairs.extend(_refine(old, new, i1, i2, j1, j2))
        elif tag == "delete":
            pairs.extend((i, None) for i in range(i1, i2))
        else:
            pairs.extend((None, j) for j in range(j1, j2))
    return [(None if i is None else i + 1, None if j is None else j + 1) for i, j in pairs]


def _extend_tail(pairs: list[tuple[int | None, int | None]]) -> list[tuple[int | None, int | None]]:
    """Lines added or removed after the last matched line are just cells
    added or removed at the end of the sheet, not inserted rows: map them
    onto the empty lines past the other version's end."""
    last = max((k for k, (i, j) in enumerate(pairs) if i is not None and j is not None), default=-1)
    li, lj = pairs[last] if last >= 0 else (0, 0)
    out = pairs[: last + 1]
    for i, j in pairs[last + 1 :]:
        if i is None:
            out.append((li + (j - lj), j))
        else:
            out.append((i, lj + (i - li)))
    # Lines from both sides past the end can now collide; keep the first.
    seen_i, seen_j, final = set(), set(), []
    for i, j in out:
        if i in seen_i or j in seen_j:
            final.append((i, None) if i not in seen_i else (None, j))
        else:
            final.append((i, j))
        seen_i.add(i)
        seen_j.add(j)
    return final


def _plan(old: Sheet, new: Sheet, alias: dict[str, str]) -> _Plan:
    """Choose rows or columns, whichever leaves fewer differences."""
    og, ng = _grid(old, alias), _grid(new)
    best: tuple[int, _Plan] | None = None
    for axis in (0, 1):
        n_old = old.max_row if axis == 0 else old.max_col
        n_new = new.max_row if axis == 0 else new.max_col
        pairs = _extend_tail(_align(_lines(og, axis, n_old), _lines(ng, axis, n_new)))
        line_map = {i: j for i, j in pairs if i is not None and j is not None}
        cost = sum(1 for i, j in pairs if (i is None) != (j is None))
        covered = set()
        for (r, c), k in og.items():
            line = r if axis == 0 else c
            if line in line_map:
                np = (line_map[line], c) if axis == 0 else (r, line_map[line])
                covered.add(np)
                cost += ng.get(np) != k
        matched = set(line_map.values())
        cost += sum(1 for p in ng if p not in covered and p[axis] in matched)
        if best is None or cost < best[0]:
            best = (cost, _Plan(old, new, axis, pairs))
    return best[1]


def _fmt(cell: Cell | None) -> str:
    if cell is None:
        return ""
    if cell.formula is not None:
        return "=" + cell.formula
    v = cell.value if cell.value is not None else cell.error
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _same(old: Cell, new: Cell, old_sheet: str, maps: Maps, alias: dict[str, str]) -> bool:
    if old.formula is not None and new.formula is not None:
        return remap(old.formula, old_sheet, maps, alias) == canonical(new.formula)
    if old.formula is not None or new.formula is not None:
        return False
    ok = ("e", old.error) if old.error and old.value is None else _value_key(old.value)
    nk = ("e", new.error) if new.error and new.value is None else _value_key(new.value)
    return ok == nk


def _change(sheet: str, old: Cell | None, new: Cell | None, old_pos: Pos | None, new_pos: Pos | None) -> Finding:
    where = cell_name(*new_pos) if new_pos else cell_name(*old_pos)
    moved = cell_name(*old_pos) if old_pos and new_pos and old_pos != new_pos else None
    o, n = _fmt(old), _fmt(new)

    def f(code: str, message: str) -> Finding:
        if moved:
            message += f" (was {moved} in the old version)"
        return Finding(code, DIFF_RULES[code].severity, message, sheet=sheet, cell=where, old_cell=moved, old=o or None, new=n or None)

    old_f = old is not None and old.formula is not None
    new_f = new is not None and new.formula is not None
    if old_f and new_f:
        return f("XD020", f"Formula changed: {o} -> {n}")
    if old_f and new is not None:
        return f("XD021", f"Formula {o} was replaced by the value {n}")
    if old_f:
        return f("XD023", f"Formula {o} was deleted")
    if new_f:
        return f("XD022", f"Formula {n} added" + (f" (was the value {o})" if old is not None else ""))
    if old is not None and new is not None:
        return f("XD024", f"Value changed: {o} -> {n}")
    if new is not None:
        return f("XD025", f"Value {n} added")
    return f("XD026", f"Value {o} deleted")


def _blocks(indexes: list[int]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for i in sorted(indexes):
        if out and i == out[-1][1] + 1:
            out[-1] = (out[-1][0], i)
        else:
            out.append((i, i))
    return out


def _apply(plan: _Plan, maps: Maps, alias: dict[str, str]) -> tuple[list[Finding], dict[Pos, Pos]]:
    old, new, axis = plan.old, plan.new, plan.axis
    line_map = plan.line_map
    name = new.name
    findings: list[Finding] = []
    pos_map: dict[Pos, Pos] = {}

    # Inserted and deleted rows/columns, as blocks.
    noun = ("Row", "Rows") if axis == 0 else ("Column", "Columns")
    label = str if axis == 0 else (lambda i: cell_name(1, i)[:-1])
    inserted = [j for i, j in plan.pairs if i is None]
    deleted = [i for i, j in plan.pairs if j is None]
    for code, idx, verb, sheet in (("XD010" if axis == 0 else "XD012", inserted, "inserted", new), ("XD011" if axis == 0 else "XD013", deleted, "deleted", old)):
        lines_with_cells = {p[axis] for p in sheet.cells}
        for a, b in _blocks(idx):
            if not any(a <= x <= b for x in lines_with_cells):
                continue  # empty lines at the edges aren't worth a mention
            span = f"{noun[0]} {label(a)}" if a == b else f"{noun[1]} {label(a)}-{label(b)}"
            findings.append(Finding(code, DIFF_RULES[code].severity, f"{span} {verb}", sheet=name, detail={"from": a, "to": b}))

    seen: set[Pos] = set()
    for (r, c), cell in old.cells.items():
        line = r if axis == 0 else c
        if line not in line_map:
            continue  # in a deleted line
        np = (line_map[line], c) if axis == 0 else (r, line_map[line])
        pos_map[(r, c)] = np
        seen.add(np)
        other = new.cells.get(np)
        if other is None or not _same(cell, other, old.name, maps, alias):
            findings.append(_change(name, cell, other, (r, c), np))
    matched = set(line_map.values())
    for (r, c), cell in new.cells.items():
        if (r, c) in seen or (r if axis == 0 else c) not in matched:
            continue
        findings.append(_change(name, None, cell, None, (r, c)))
    return findings, pos_map


def _match_sheets(old: Workbook, new: Workbook) -> tuple[list[tuple[Sheet, Sheet]], dict[str, str], list[Sheet], list[Sheet]]:
    new_by = {s.name.lower(): s for s in new.sheets}
    old_by = {s.name.lower(): s for s in old.sheets}
    removed = [s for s in old.sheets if s.name.lower() not in new_by]
    added = [s for s in new.sheets if s.name.lower() not in old_by]
    renamed: dict[str, str] = {}
    for o in removed:
        og = _grid(o)
        scored = []
        for n in added:
            if n.name in renamed.values():
                continue
            ng = _grid(n)
            total = max(len(og), len(ng)) or 1
            scored.append((sum(1 for p, k in og.items() if ng.get(p) == k) / total, n))
        scored.sort(key=lambda x: -x[0])
        if scored and scored[0][0] >= 0.5:
            renamed[o.name] = scored[0][1].name
    pairs = [(o, new_by[o.name.lower()]) for o in old.sheets if o.name.lower() in new_by]
    pairs += [(o, new_by[renamed[o.name].lower()]) for o in removed if o.name in renamed]
    return pairs, renamed, [s for s in removed if s.name not in renamed], [s for s in added if s.name not in renamed.values()]


def diff(old: Workbook, new: Workbook, *, include_values: bool = True, lint_select: set[str] | None = None,
         lint_ignore: set[str] | None = None) -> DiffResult:
    """Changes between two versions, plus lint findings that are new or resolved."""
    findings: list[Finding] = []
    pairs, renamed, removed, added = _match_sheets(old, new)
    for o_name, n_name in renamed.items():
        findings.append(Finding("XD003", "info", f"Sheet {o_name} renamed to {n_name}", sheet=n_name, old=o_name, new=n_name))
    for s in removed:
        findings.append(Finding("XD002", "warning", f"Sheet {s.name} removed ({len(s.cells)} cells)", sheet=s.name))
    for s in added:
        findings.append(Finding("XD001", "info", f"Sheet {s.name} added ({len(s.cells)} cells)", sheet=s.name))

    plans = [_plan(o, n, renamed) for o, n in pairs]
    maps: Maps = {p.old.name.lower(): (p.new.name, p.axis, p.line_map) for p in plans}
    pos_maps: dict[str, tuple[str, dict[Pos, Pos]]] = {}
    for plan in plans:
        sheet_findings, pos_map = _apply(plan, maps, renamed)
        findings.extend(sheet_findings)
        pos_maps[plan.old.name.lower()] = (plan.new.name, pos_map)

    findings.extend(_diff_names(old, new, renamed))
    findings.extend(_diff_links(old, new))
    findings.extend(_diff_lint(old, new, pos_maps, lint_select, lint_ignore))

    if not include_values:
        findings = [f for f in findings if f.code not in ("XD024", "XD025", "XD026")]
    order = [s.name for s in new.sheets] + [s.name for s in removed]
    return DiffResult(sort_findings(findings, order), renamed, {"old": old.stats, "new": new.stats})


def _diff_names(old: Workbook, new: Workbook, alias: dict[str, str]) -> list[Finding]:
    def key(dn):
        return (dn.name.lower(), (dn.scope or "").lower())

    old_names = {key(dn): dn for dn in old.names}
    new_names = {key(dn): dn for dn in new.names}
    out = []
    for k, dn in new_names.items():
        before = old_names.get(k)
        if before is None:
            out.append(Finding("XD030", "info", f"Defined name {dn.name} added: {dn.text}", sheet=dn.scope, new=dn.text))
        elif normalize(before.text, 1, 1, alias) != normalize(dn.text, 1, 1):
            out.append(Finding("XD030", "info", f"Defined name {dn.name} changed: {before.text} -> {dn.text}", sheet=dn.scope, old=before.text, new=dn.text))
    for k, dn in old_names.items():
        if k not in new_names:
            out.append(Finding("XD030", "info", f"Defined name {dn.name} removed (was {dn.text})", sheet=dn.scope, old=dn.text))
    return out


def _diff_links(old: Workbook, new: Workbook) -> list[Finding]:
    out = []
    for t in new.external_links:
        if t not in old.external_links:
            out.append(Finding("XD031", "warning", f"New link to another workbook: {t}", new=t))
    for t in old.external_links:
        if t not in new.external_links:
            out.append(Finding("XD032", "info", f"Link to another workbook removed: {t}", old=t))
    return out


def _issue_key(f: Finding, sheet: str | None, cell: str | None) -> tuple:
    if cell:
        return (f.code, (sheet or "").lower(), cell)
    if sheet:  # sheet-level: hidden sheets, sheet-scoped names
        return (f.code, sheet.lower(), f.detail.get("name", ""))
    return (f.code, "", f.message)


def _diff_lint(old: Workbook, new: Workbook, pos_maps, select, ignore) -> list[Finding]:
    old_keys: dict[tuple, Finding] = {}
    for f in lint(old, select, ignore):
        sheet, cell = f.sheet, f.cell
        if sheet and sheet.lower() in pos_maps:
            new_sheet, pos_map = pos_maps[sheet.lower()]
            if cell:
                mapped = pos_map.get(parse_cell(cell))
                cell = cell_name(*mapped) if mapped else None
            sheet = new_sheet
        if cell is None and f.cell:
            old_keys[(f.code, "gone", f.sheet, f.cell)] = f  # its cell was deleted
            continue
        old_keys[_issue_key(f, sheet, cell)] = f
    new_keys = set()
    out = []
    for f in lint(new, select, ignore):
        k = _issue_key(f, f.sheet, f.cell)
        new_keys.add(k)
        if k not in old_keys:
            sev = f.severity if severity_rank(f.severity) >= severity_rank("warning") else "info"
            out.append(Finding("XD040", sev, f"New {f.code} {f.slug}: {f.message}", sheet=f.sheet, cell=f.cell, detail={"rule": f.code}))
    for k, f in old_keys.items():
        if k not in new_keys:
            out.append(Finding("XD041", "info", f"Resolved {f.code} {f.slug}: {f.message}", sheet=f.sheet, old_cell=f.cell, detail={"rule": f.code}))
    return out

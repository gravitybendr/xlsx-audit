"""Make the README images and logo PNGs from real xlsx-audit output.

    python scripts/make_images.py

1. Builds docs/demo/model_v1.xlsx and model_v2.xlsx (scripts/demo_workbooks.py).
2. Runs the xlsx-audit CLI on them and saves exactly what it prints to
   docs/demo/output/ (text, JSON and Markdown).
3. Writes one HTML page per image into build/images/, using only that saved
   output and the cells of the demo workbooks.
4. Renders the pages with headless Chromium inside the
   mcr.microsoft.com/playwright image (scripts/render_pages.sh) and copies the
   PNGs to docs/images/ and docs/brand/.

Needs Docker and that image (it has Chromium and the Liberation fonts).
Nothing is downloaded while rendering.
"""
from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEMO = ROOT / "docs" / "demo"
OUTPUT = DEMO / "output"
IMAGES = ROOT / "docs" / "images"
BRAND = ROOT / "docs" / "brand"
BUILD = ROOT / "build" / "images"
IMAGE = "mcr.microsoft.com/playwright:v1.64.1-noble"

sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "scripts"))

import demo_workbooks  # noqa: E402

from xlsx_audit import load  # noqa: E402
from xlsx_audit.refs import cell_name  # noqa: E402

LINT_CMD = ["lint", "model_v1.xlsx"]
DIFF_CMD = ["diff", "model_v1.xlsx", "model_v2.xlsx"]

# Brand and severity colours.
INDIGO, AMBER, PALE = "#25368A", "#FFB020", "#D3DDF5"
SEV = {
    "error": ("#D93A48", "#FDEBEC"),
    "warning": ("#C97A00", "#FFF3D6"),
    "info": ("#2F6FE0", "#E7EFFE"),
}
SEV_RANK = {"info": 0, "warning": 1, "error": 2}
STATE = {"hidden": "hidden", "veryHidden": "very hidden"}

BASE_CSS = f"""
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; background: transparent; }}
body {{ font-family: 'Liberation Sans', sans-serif; color: #1B2140; }}
#frame {{ display: inline-block; }}
.mono {{ font-family: 'Liberation Mono', monospace; }}
.logo {{ display: block; }}
.badge {{ display: inline-block; font: bold 10.5px/1 'Liberation Sans', sans-serif; color: #fff;
  padding: 3px 5px; border-radius: 4px; letter-spacing: .02em; }}
"""

MEASURE = """<script>
addEventListener('load', () => {
  const r = document.getElementById('frame').getBoundingClientRect();
  document.body.dataset.size = Math.ceil(r.width) + 'x' + Math.ceil(r.height);
});
</script>"""


def esc(text: object) -> str:
    return html.escape(str(text), quote=True)


def page(body: str, css: str = "", scale: float = 2) -> str:
    return (
        f'<!doctype html><html data-scale="{scale}"><head><meta charset="utf-8">'
        f"<style>{BASE_CSS}{css}</style></head><body>{body}{MEASURE}</body></html>"
    )


# ---------------------------------------------------------------- CLI output


def run_cli(args: list[str]) -> str:
    env = {**os.environ, "PYTHONPATH": str(ROOT / "src")}
    proc = subprocess.run(
        [sys.executable, "-m", "xlsx_audit", *args], cwd=DEMO, env=env,
        capture_output=True, text=True, encoding="utf-8",
    )
    if proc.returncode not in (0, 1):
        raise SystemExit(f"xlsx-audit {' '.join(args)} failed:\n{proc.stderr}")
    return proc.stdout


def capture() -> dict[str, str]:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    out = {
        "lint.txt": run_cli(LINT_CMD),
        "lint.json": run_cli(LINT_CMD + ["--format", "json"]),
        "lint.md": run_cli(LINT_CMD + ["--format", "markdown"]),
        "diff.txt": run_cli(DIFF_CMD),
        "diff.json": run_cli(DIFF_CMD + ["--format", "json"]),
        "diff.md": run_cli(DIFF_CMD + ["--format", "markdown"]),
    }
    for name, text in out.items():
        (OUTPUT / name).write_text(text, encoding="utf-8", newline="\n")
    return out


# ---------------------------------------------------------------- pieces


def logo_img(size: int) -> str:
    return f'<img class="logo" src="{(BRAND / "logo.svg").as_uri()}" width="{size}" height="{size}" alt="">'


def cell_text(cell) -> str:
    if cell.formula is not None:
        return "=" + cell.formula
    v = cell.value
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)


def worst(findings: list[dict]) -> str:
    return max((f["severity"] for f in findings), key=SEV_RANK.__getitem__)


def badges(findings: list[dict], label=lambda f: f["code"]) -> str:
    return "".join(
        f'<span class="badge" style="background:{SEV[f["severity"]][0]}">{esc(label(f))}</span>' for f in findings
    )


GRID_CSS = """
.sheet { background: #fff; border-radius: 10px; box-shadow: 0 1px 2px rgba(27,33,64,.08), 0 8px 24px rgba(27,33,64,.10);
  overflow: hidden; }
.sheet table { border-collapse: collapse; width: 100%; }
.sheet th { background: #F1F3F9; color: #6B7393; font: 12px 'Liberation Sans', sans-serif; height: 26px;
  border: 1px solid #E1E5EF; font-weight: normal; }
.sheet th.rn { width: 34px; }
.sheet td { position: relative; height: 38px; padding: 0 10px; border: 1px solid #E6E9F2;
  font: 13px 'Liberation Mono', monospace; white-space: nowrap; color: #26304F; }
.sheet td.head { font: bold 13px 'Liberation Sans', sans-serif; color: #1B2140; }
.sheet td.num { text-align: right; }
.sheet td.flag { outline-offset: -2px; }
.sheet td .tags { position: absolute; top: -9px; right: 5px; display: flex; gap: 3px; z-index: 2; }
.sheet tr.band td, .sheet tr.band th { background: #E6F6EC; }
.tabs { display: flex; gap: 2px; padding: 0 0 0 34px; background: #F1F3F9; border-top: 1px solid #E1E5EF; height: 34px; align-items: stretch; }
.tab { display: flex; align-items: center; gap: 6px; padding: 0 14px; font-size: 12.5px; color: #4A5375; }
.tab.on { background: #fff; color: #1B2140; font-weight: bold; box-shadow: inset 0 2px 0 #25368A; }
.tab.ghost { color: #8A92B0; font-style: italic; border: 1px dashed #B8C0D8; border-top: 0; border-bottom: 0; }
"""


def grid_html(sheet, flags: dict[str, list[dict]], cols: int, *, band_row: int | None = None,
              tabs: str = "", label=lambda f: f["code"]) -> str:
    """The cells of one sheet as a table, formulas shown as formulas."""
    rows = [f'<tr><th class="rn"></th>' + "".join(f"<th>{cell_name(1, c)[:-1]}</th>" for c in range(1, cols + 1)) + "</tr>"]
    for r in range(1, sheet.max_row + 1):
        tds = []
        for c in range(1, cols + 1):
            cell = sheet.cells.get((r, c))
            ref = cell_name(r, c)
            text = esc(cell_text(cell)) if cell else ""
            cls = []
            if cell is not None and r == 1:
                cls.append("head")
            elif cell is not None and cell.formula is None and isinstance(cell.value, (int, float)):
                cls.append("num")
            style = ""
            tags = ""
            if ref in flags:
                found = flags[ref]
                line, fill = SEV[worst(found)]
                cls.append("flag")
                style = f' style="background:{fill};outline:2px solid {line}"'
                tags = f'<span class="tags">{badges(found, label)}</span>'
            tds.append(f'<td class="{" ".join(cls)}"{style}>{tags}{text}</td>')
        band = ' class="band"' if band_row == r else ""
        rows.append(f'<tr{band}><th class="rn">{r}</th>{"".join(tds)}</tr>')
    return f'<div class="sheet"><table>{"".join(rows)}</table>{tabs}</div>'


# ---------------------------------------------------------------- hero (lint)


def hero(lint: dict) -> str:
    rep = lint["reports"][0]
    wb = load(DEMO / "model_v1.xlsx")
    sales = wb.sheet("Sales")
    flags: dict[str, list[dict]] = {}
    for f in rep["findings"]:
        if f.get("sheet") == "Sales" and f.get("cell"):
            flags.setdefault(f["cell"], []).append(f)
    by_sheet = {f["sheet"]: f for f in rep["findings"] if f.get("sheet") and not f.get("cell")}
    tabs = []
    for s in wb.sheets:
        if s.state == "visible":
            tabs.append(f'<div class="tab{" on" if s.name == "Sales" else ""}">{esc(s.name)}</div>')
        else:
            f = by_sheet.get(s.name)
            tabs.append(f'<div class="tab ghost">{esc(s.name)} ({STATE[s.state]}){badges([f]) if f else ""}</div>')
    grid = grid_html(sales, flags, sales.max_col, tabs=f'<div class="tabs">{"".join(tabs)}</div>')

    items = []
    for f in rep["findings"]:
        line, _ = SEV[f["severity"]]
        where = f'{f["sheet"]}!{f["cell"]}' if f.get("cell") else (f.get("sheet") or "(workbook)")
        items.append(
            f'<li><span class="dot" style="background:{line}"></span>'
            f'<span class="code">{esc(f["code"])}</span><span class="where mono">{esc(where)}</span>'
            f'<span class="slug">{esc(f["rule"])}</span></li>'
        )
    counts = rep["counts"]
    pills = "".join(
        f'<span class="pill" style="color:{SEV[s][0]};background:{SEV[s][1]}">{counts[s]} {s + ("s" if s != "info" and counts[s] != 1 else "")}</span>'
        for s in ("error", "warning", "info") if counts.get(s)
    )
    stats = rep["stats"]
    detail = [f for f in rep["findings"] if f["code"] in ("XA004", "XA005")]
    notes = "".join(
        f'<div class="note"><span class="badge" style="background:{SEV[f["severity"]][0]}">{esc(f["code"])}</span>'
        f'<span class="mono where">{esc(f["sheet"])}!{esc(f["cell"])}</span> {esc(f["message"])}</div>'
        for f in detail
    )
    css = GRID_CSS + f"""
#frame {{ width: 1240px; padding: 34px 36px 32px; border-radius: 18px;
  background: linear-gradient(135deg, #EEF1FB 0%, #E2E8F8 100%); }}
.top {{ display: flex; align-items: center; gap: 14px; margin-bottom: 22px; }}
.top .cmd {{ font: 17px 'Liberation Mono', monospace; color: #1B2140; }}
.top .cmd b {{ color: {INDIGO}; }}
.top .meta {{ margin-left: auto; display: flex; gap: 8px; align-items: center; color: #5A6387; font-size: 13px; }}
.pill {{ font: bold 13px 'Liberation Sans', sans-serif; padding: 5px 10px; border-radius: 999px; }}
.main {{ display: flex; gap: 22px; align-items: flex-start; }}
.panel {{ flex: 1; background: #fff; border-radius: 10px; padding: 14px 16px 10px;
  box-shadow: 0 1px 2px rgba(27,33,64,.08), 0 8px 24px rgba(27,33,64,.10); }}
.panel h3 {{ margin: 0 0 8px; font-size: 13px; color: #5A6387; font-weight: bold; letter-spacing: .04em; text-transform: uppercase; }}
.panel ul {{ list-style: none; margin: 0; padding: 0; }}
.panel li {{ display: flex; align-items: center; gap: 8px; height: 29px; border-top: 1px solid #EEF0F6; font-size: 13px; }}
.panel li:first-child {{ border-top: 0; }}
.dot {{ width: 9px; height: 9px; border-radius: 50%; flex: none; }}
.code {{ font-weight: bold; width: 46px; flex: none; }}
.where {{ width: 92px; flex: none; color: #26304F; font-size: 12.5px; }}
.slug {{ color: #5A6387; }}
.notes {{ margin-top: 16px; display: flex; flex-direction: column; gap: 8px; }}
.note {{ background: rgba(255,255,255,.7); border-radius: 8px; padding: 9px 12px; font-size: 13px; color: #26304F; }}
.note .badge {{ margin-right: 8px; }}
.note .where {{ margin-right: 6px; font-weight: bold; }}
"""
    body = f"""<div id="frame">
<div class="top">{logo_img(38)}<span class="cmd"><b>xlsx-audit</b> {esc(' '.join(LINT_CMD))}</span>
<span class="meta">{stats['sheets']} sheets, {stats['cells']} cells, {stats['formulas']} formulas &nbsp; {pills}</span></div>
<div class="main"><div>{grid}<div class="notes">{notes}</div></div>
<div class="panel"><h3>{len(rep['findings'])} findings</h3><ul>{''.join(items)}</ul></div></div>
</div>"""
    return page(body, css)


# ---------------------------------------------------------------- diff grid


def diff_grid(diff: dict) -> str:
    rep = diff["reports"][0]
    old_wb, new_wb = load(DEMO / "model_v1.xlsx"), load(DEMO / "model_v2.xlsx")
    old, new = old_wb.sheet("Sales"), new_wb.sheet("Sales")
    cols = 4  # columns A-D tell the story; E and F only move down a row
    new_flags: dict[str, list[dict]] = {}
    old_flags: dict[str, list[dict]] = {}
    inserted = None
    for f in rep["findings"]:
        if f.get("sheet") != "Sales":
            continue
        if f["code"] == "XD010":
            inserted = f["detail"]["from"]
        if f.get("cell"):
            new_flags.setdefault(f["cell"], []).append(f)
        if f.get("old_cell"):
            old_flags.setdefault(f["old_cell"], []).append(f)
    # Cells that moved to a new address but aren't reported.
    reported_old = {f.get("old_cell") for f in rep["findings"] if f.get("sheet") == "Sales"}
    moved = sum(1 for (r, c) in old.cells if inserted and r >= inserted and cell_name(r, c) not in reported_old)

    def tabs(wb, renamed: dict[str, dict]) -> str:
        out = []
        for s in wb.sheets:
            if s.state != "visible":
                f = renamed.get(s.name)
                out.append(f'<div class="tab ghost">{esc(s.name)} ({STATE[s.state]}){badges([f]) if f else ""}</div>')
            else:
                out.append(f'<div class="tab{" on" if s.name == "Sales" else ""}">{esc(s.name)}</div>')
        return f'<div class="tabs">{"".join(out)}</div>'

    renamed = {f["sheet"]: f for f in rep["findings"] if f["code"] == "XD003"}
    left = grid_html(old, old_flags, cols, tabs=tabs(old_wb, {}))
    right = grid_html(new, new_flags, cols, band_row=inserted, tabs=tabs(new_wb, renamed))
    lines = []
    for f in rep["findings"]:
        line, _ = SEV[f["severity"]]
        where = f'{f["sheet"]}!{f.get("cell") or f.get("old_cell")}' if (f.get("cell") or f.get("old_cell")) else f.get("sheet", "")
        lines.append(
            f'<li><span class="dot" style="background:{line}"></span><span class="code">{esc(f["code"])}</span>'
            f'<span class="where mono">{esc(where)}</span><span class="msg">{esc(f["message"])}</span></li>'
        )
    css = GRID_CSS + f"""
#frame {{ width: 1240px; padding: 30px 36px 30px; border-radius: 18px;
  background: linear-gradient(135deg, #EEF1FB 0%, #E2E8F8 100%); }}
.top {{ display: flex; align-items: center; gap: 14px; margin-bottom: 20px; }}
.top .cmd {{ font: 17px 'Liberation Mono', monospace; }}
.top .cmd b {{ color: {INDIGO}; }}
.pair {{ display: flex; gap: 26px; align-items: flex-start; }}
.pair .col {{ flex: 1; }}
.col h4 {{ margin: 0 0 8px 2px; font: bold 13px 'Liberation Sans', sans-serif; color: #5A6387; letter-spacing: .04em; text-transform: uppercase; }}
.col h4 span {{ font-family: 'Liberation Mono', monospace; text-transform: none; letter-spacing: 0; color: #26304F; }}
.sheet td {{ min-width: 64px; }}
.changes {{ margin-top: 18px; background: #fff; border-radius: 10px; padding: 8px 16px;
  box-shadow: 0 1px 2px rgba(27,33,64,.08), 0 8px 24px rgba(27,33,64,.10); }}
.changes ul {{ list-style: none; margin: 0; padding: 0; }}
.changes li {{ display: flex; align-items: baseline; gap: 9px; padding: 7px 0; border-top: 1px solid #EEF0F6; font-size: 13px; }}
.changes li:first-child {{ border-top: 0; }}
.dot {{ width: 9px; height: 9px; border-radius: 50%; flex: none; align-self: center; }}
.code {{ font-weight: bold; width: 46px; flex: none; }}
.where {{ width: 112px; flex: none; font-size: 12.5px; }}
.msg {{ color: #3A4364; }}
.foot {{ margin: 12px 2px 0; font-size: 13.5px; color: #3A4364; }}
.foot b {{ color: #1B2140; }}
"""
    body = f"""<div id="frame">
<div class="top">{logo_img(38)}<span class="cmd"><b>xlsx-audit</b> {esc(' '.join(DIFF_CMD))}</span></div>
<div class="pair">
<div class="col"><h4>Old <span>model_v1.xlsx</span> &middot; Sales, columns A-D</h4>{left}</div>
<div class="col"><h4>New <span>model_v2.xlsx</span> &middot; Sales, columns A-D</h4>{right}</div>
</div>
<div class="changes"><ul>{''.join(lines)}</ul></div>
<p class="foot">Row {inserted} was inserted, so <b>{moved} other cells</b> on Sales moved down one row. None of them is reported: formulas are compared after following their references to where the cells moved.</p>
</div>"""
    return page(body, css)


# ---------------------------------------------------------------- terminal cards

TERM_CSS = """
#frame { padding: 0; }
.term { background: #131833; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 30px rgba(19,24,51,.25); }
.bar { height: 34px; display: flex; align-items: center; gap: 7px; padding: 0 14px; background: #1C2245; }
.bar i { width: 11px; height: 11px; border-radius: 50%; background: #3A4270; display: block; }
.bar span { margin-left: 10px; font: 12px 'Liberation Sans', sans-serif; color: #8C97BF; }
pre { margin: 0; padding: 18px 22px 22px; font: 13.5px/1.55 'Liberation Mono', monospace; color: #D9E0F5; }
.p { color: #8C97BF; } .c { color: #FFFFFF; font-weight: bold; }
.e { color: #FF7A85; font-weight: bold; } .w { color: #FFC24D; font-weight: bold; } .n { color: #7FA8FF; font-weight: bold; }
.dim { color: #8C97BF; } .loc { color: #B9C6F2; }
"""


def terminal(cmd: list[str], text: str, title: str) -> str:
    out = [f'<span class="p">$</span> <span class="c">xlsx-audit {esc(" ".join(cmd))}</span>']
    for raw in text.rstrip("\n").split("\n"):
        line = esc(raw)
        m = re.match(r"^(ERROR|WARNING|INFO)(\s+)(.*)$", raw)
        if m:
            cls = {"ERROR": "e", "WARNING": "w", "INFO": "n"}[m.group(1)]
            line = f'<span class="{cls}">{m.group(1)}</span>{m.group(2)}{esc(m.group(3))}'
        elif raw.startswith("  ") and raw.strip():
            m2 = re.match(r"^(  )(\S+|'[^']*')(\s+)(.*)$", raw)
            if m2:
                line = f'{m2.group(1)}<span class="loc">{esc(m2.group(2))}</span>{m2.group(3)}{esc(m2.group(4))}'
        elif raw.startswith("xlsx-audit ") or raw.startswith(("old:", "new:")) or re.match(r"^\d+ sheets", raw):
            line = f'<span class="dim">{line}</span>'
        out.append(line)
    body = (
        f'<div id="frame"><div class="term"><div class="bar"><i></i><i></i><i></i><span>{esc(title)}</span></div>'
        f'<pre>{chr(10).join(out)}</pre></div></div>'
    )
    return page(body, TERM_CSS)


# ---------------------------------------------------------------- markdown report


def _inline(text: str) -> str:
    text = esc(text)
    text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
    return re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)


def _cells(row: str) -> list[str]:
    parts = re.split(r"(?<!\\)\|", row.strip())[1:-1]
    return [p.strip().replace("\\|", "|") for p in parts]


def markdown_html(md: str) -> str:
    """Just enough Markdown for xlsx-audit's report: ### headings, paragraphs, one table."""
    out, table = [], []

    def flush():
        if table:
            head, rows = _cells(table[0]), [_cells(r) for r in table[2:]]
            out.append("<table><thead><tr>" + "".join(f"<th>{_inline(h)}</th>" for h in head) + "</tr></thead><tbody>"
                       + "".join("<tr>" + "".join(f"<td>{_inline(c)}</td>" for c in r) + "</tr>" for r in rows)
                       + "</tbody></table>")
            table.clear()

    for line in md.splitlines():
        if line.startswith("|"):
            table.append(line)
            continue
        flush()
        if line.startswith("### "):
            out.append(f"<h3>{_inline(line[4:])}</h3>")
        elif line.strip():
            out.append(f"<p>{_inline(line)}</p>")
    flush()
    return "\n".join(out)


def markdown_card(md: str) -> str:
    css = """
#frame { width: 1180px; padding: 26px 30px 18px; background: #fff; border-radius: 12px; border: 1px solid #D8DEEC; }
.label { font: 12px 'Liberation Sans', sans-serif; color: #6B7393; letter-spacing: .04em; text-transform: uppercase; margin-bottom: 10px; }
h3 { font-size: 20px; margin: 0 0 12px; }
p { font-size: 14.5px; margin: 0 0 14px; }
code { font: 13px 'Liberation Mono', monospace; background: #F1F3F9; padding: 2px 5px; border-radius: 4px; }
table { border-collapse: collapse; font-size: 13.5px; width: 100%; }
th, td { border: 1px solid #DDE2EE; padding: 7px 10px; text-align: left; vertical-align: top; }
th { background: #F6F8FC; }
tr:nth-child(even) td { background: #FAFBFD; }
"""
    body = f'<div id="frame"><div class="label">--format markdown, rendered</div>{markdown_html(md)}</div>'
    return page(body, css)


# ---------------------------------------------------------------- logo and social card


def logo_page(size: int) -> str:
    css = f"#frame {{ width: {size}px; height: {size}px; }}"
    return page(f'<div id="frame">{logo_img(size)}</div>', css, scale=1)


def social() -> str:
    css = f"""
#frame {{ width: 1280px; height: 640px; background: {INDIGO}; position: relative; overflow: hidden;
  display: flex; align-items: center; padding: 0 96px; gap: 56px; }}
.grid {{ position: absolute; inset: 0; background-image:
  linear-gradient(rgba(211,221,245,.07) 1px, transparent 1px), linear-gradient(90deg, rgba(211,221,245,.07) 1px, transparent 1px);
  background-size: 64px 40px; }}
.mark {{ position: relative; background: #fff; border-radius: 40px; padding: 28px; flex: none; }}
.text {{ position: relative; color: #fff; }}
h1 {{ font: bold 76px 'Liberation Sans', sans-serif; margin: 0 0 14px; letter-spacing: -1px; }}
p {{ font: 30px/1.35 'Liberation Sans', sans-serif; margin: 0 0 30px; color: {PALE}; max-width: 720px; }}
.chips {{ display: flex; flex-wrap: wrap; gap: 10px; max-width: 760px; }}
.chips span {{ font: 20px 'Liberation Mono', monospace; color: #1B2140; background: {PALE}; padding: 7px 12px; border-radius: 8px; }}
.chips span.a {{ background: {AMBER}; }}
"""
    chips = ["#REF!", "numbers typed over formulas", "external links", "row-aware diff", "SARIF, JSON, Markdown", "pip install xlsx-audit"]
    body = f"""<div id="frame"><div class="grid"></div><div class="mark">{logo_img(220)}</div>
<div class="text"><h1>xlsx-audit</h1><p>Find the mistakes in Excel workbooks and see what changed between two versions.</p>
<div class="chips">{''.join(f'<span class="{"a" if i == 1 else ""}">{esc(c)}</span>' for i, c in enumerate(chips))}</div></div></div>"""
    return page(body, css, scale=1)


# ---------------------------------------------------------------- main

PAGES = {
    # build/images/<name>.html -> destination of the PNG
    "hero": IMAGES / "hero.png",
    "diff-grid": IMAGES / "diff.png",
    "lint-terminal": IMAGES / "lint-output.png",
    "diff-terminal": IMAGES / "diff-output.png",
    "markdown": IMAGES / "markdown-report.png",
    "social": BRAND / "social-preview.png",
    "logo-64": BRAND / "logo-64.png",
    "logo-128": BRAND / "logo-128.png",
    "logo-256": BRAND / "logo-256.png",
    "logo-512": BRAND / "logo-512.png",
}


def render() -> None:
    cmd = [
        "docker", "run", "--rm", "-v", f"{ROOT}:/work", "-w", "/work", IMAGE,
        "bash", "scripts/render_pages.sh", "build/images",
    ]
    env = {**os.environ, "MSYS_NO_PATHCONV": "1"}
    subprocess.run(cmd, check=True, env=env)


def main() -> None:
    demo_workbooks.build(DEMO)
    out = capture()
    # The pages load files by absolute path; inside Docker the repo is /work.
    pages = {
        "hero": hero(json.loads(out["lint.json"])),
        "diff-grid": diff_grid(json.loads(out["diff.json"])),
        "lint-terminal": terminal(LINT_CMD, out["lint.txt"], "xlsx-audit lint"),
        "diff-terminal": terminal(DIFF_CMD, out["diff.txt"], "xlsx-audit diff"),
        "markdown": markdown_card(out["lint.md"]),
        "social": social(),
        **{f"logo-{n}": logo_page(n) for n in (64, 128, 256, 512)},
    }
    if BUILD.exists():
        shutil.rmtree(BUILD)
    BUILD.mkdir(parents=True)
    local = ROOT.as_uri()
    for name, text in pages.items():
        (BUILD / f"{name}.html").write_text(text.replace(local, "file:///work"), encoding="utf-8")
    render()
    IMAGES.mkdir(parents=True, exist_ok=True)
    for name, dest in PAGES.items():
        shutil.copyfile(BUILD / f"{name}.png", dest)
        print(dest.relative_to(ROOT))


if __name__ == "__main__":
    main()

"""Write small .xlsx files from raw SpreadsheetML for tests.

openpyxl can't write cached formula results, shared formulas or external
links, which are exactly the cases the audit has to handle. This builder
writes those parts directly, the way Excel stores them.
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


@dataclass
class F:
    """A formula cell.

    text: formula without the leading "=".
    cached: last calculated value (number, str, bool) or an error like "#REF!".
    shared: (ref, si) for the master cell of a shared formula.
    shared_child: si for a cell that reuses a shared formula (text ignored).
    array_ref: range for an array (CSE) formula.
    """

    text: str = ""
    cached: object = None
    shared: tuple[str, int] | None = None
    shared_child: int | None = None
    array_ref: str | None = None


def _col_row(ref: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return col, int(m.group(2))


def _cell_xml(ref: str, value) -> str:
    if isinstance(value, F):
        attrs = ""
        cached = value.cached
        if isinstance(cached, str) and cached.startswith("#"):
            attrs = ' t="e"'
        elif isinstance(cached, bool):
            attrs = ' t="b"'
        elif isinstance(cached, str):
            attrs = ' t="str"'
        if value.shared_child is not None:
            f = f'<f t="shared" si="{value.shared_child}"/>'
        elif value.shared is not None:
            sref, si = value.shared
            f = f'<f t="shared" ref="{sref}" si="{si}">{escape(value.text)}</f>'
        elif value.array_ref is not None:
            f = f'<f t="array" ref="{value.array_ref}">{escape(value.text)}</f>'
        else:
            f = f"<f>{escape(value.text)}</f>"
        v = ""
        if cached is not None:
            cv = ("1" if cached else "0") if isinstance(cached, bool) else str(cached)
            v = f"<v>{escape(cv)}</v>"
        return f'<c r="{ref}"{attrs}>{f}{v}</c>'
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{"1" if value else "0"}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"><v>{value}</v></c>'
    if isinstance(value, str) and value.startswith("#") and value.endswith(("!", "?", "A")):
        return f'<c r="{ref}" t="e"><v>{escape(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def _sheet_xml(cells: dict[str, object]) -> str:
    rows: dict[int, list[tuple[int, str]]] = {}
    for ref, value in cells.items():
        col, row = _col_row(ref)
        rows.setdefault(row, []).append((col, _cell_xml(ref, value)))
    body = "".join(
        f'<row r="{r}">' + "".join(x for _, x in sorted(cs)) + "</row>"
        for r, cs in sorted(rows.items())
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}"><sheetData>{body}</sheetData></worksheet>'


def build_xlsx(
    path: Path,
    sheets: dict[str, dict[str, object]],
    *,
    states: dict[str, str] | None = None,
    names: dict[str, str] | None = None,
    sheet_names: dict[str, tuple[str, str]] | None = None,
    external_links: list[str] | None = None,
) -> Path:
    """Write an .xlsx file.

    sheets: sheet name -> {cell ref: value or F(...)}, in tab order.
    states: sheet name -> "hidden" or "veryHidden".
    names: workbook-scoped defined names -> formula text (no "=").
    sheet_names: sheet-scoped defined names, name -> (sheet, formula text).
    external_links: target paths of linked workbooks, referenced as [1], [2]...
    """
    states = states or {}
    names = names or {}
    sheet_names = sheet_names or {}
    external_links = external_links or []
    sheet_list = list(sheets)

    overrides = [
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    ]
    wb_rels = []
    sheet_tags = []
    for i, name in enumerate(sheet_list, 1):
        state = f' state="{states[name]}"' if name in states else ""
        sheet_tags.append(f'<sheet name={quoteattr(name)} sheetId="{i}"{state} r:id="rId{i}"/>')
        wb_rels.append(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>')
        overrides.append((f"/xl/worksheets/sheet{i}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))

    ext_tags = []
    for j, _target in enumerate(external_links, 1):
        rid = f"rIdX{j}"
        ext_tags.append(f'<externalReference r:id="{rid}"/>')
        wb_rels.append(f'<Relationship Id="{rid}" Type="{REL}/externalLink" Target="externalLinks/externalLink{j}.xml"/>')
        overrides.append((f"/xl/externalLinks/externalLink{j}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"))

    name_tags = [f"<definedName name={quoteattr(n)}>{escape(t)}</definedName>" for n, t in names.items()]
    for n, (sheet, t) in sheet_names.items():
        idx = sheet_list.index(sheet)
        name_tags.append(f'<definedName name={quoteattr(n)} localSheetId="{idx}">{escape(t)}</definedName>')

    workbook = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN}" xmlns:r="{REL}">'
        f"<sheets>{''.join(sheet_tags)}</sheets>"
        + (f"<externalReferences>{''.join(ext_tags)}</externalReferences>" if ext_tags else "")
        + (f"<definedNames>{''.join(name_tags)}</definedNames>" if name_tags else "")
        + "</workbook>"
    )
    content_types = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="{CT}">'
        f'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides)
        + "</Types>"
    )
    root_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG_REL}">{"".join(wb_rels)}</Relationships>',
        )
        for i, name in enumerate(sheet_list, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(sheets[name]))
        for j, target in enumerate(external_links, 1):
            z.writestr(
                f"xl/externalLinks/externalLink{j}.xml",
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="{MAIN}" xmlns:r="{REL}">'
                f'<externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>',
            )
            z.writestr(
                f"xl/externalLinks/_rels/externalLink{j}.xml.rels",
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG_REL}">'
                f'<Relationship Id="rId1" Type="{REL}/externalLinkPath" Target={quoteattr(target)} TargetMode="External"/></Relationships>',
            )
    return path

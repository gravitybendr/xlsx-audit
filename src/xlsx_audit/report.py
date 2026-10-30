"""Render reports as text, JSON, SARIF 2.1.0 or Markdown. Output is plain ASCII."""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path

from . import __version__
from .model import ALL_RULES, SEVERITIES, Finding

INFO_URI = "https://github.com/gravitybendr/xlsx-audit"


@dataclass
class Report:
    kind: str  # "lint" or "diff"
    files: list[str]
    findings: list[Finding]
    stats: dict = field(default_factory=dict)

    @property
    def counts(self) -> dict[str, int]:
        out = {s: 0 for s in reversed(SEVERITIES)}
        for f in self.findings:
            out[f.severity] += 1
        return out


def _plural(n: int, word: str) -> str:
    if word == "info":
        return f"{n} info"
    return f"{n:,} {word}" + ("" if n == 1 else "s")


def _stats_line(stats: dict) -> str:
    return f"{stats.get('sheets', 0)} sheets, {stats.get('cells', 0):,} cells, {stats.get('formulas', 0):,} formulas"


def _groups(findings: list[Finding]) -> list[tuple[str, str, list[Finding]]]:
    groups: dict[tuple[str, str], list[Finding]] = {}
    for f in findings:
        groups.setdefault((f.severity, f.code), []).append(f)
    order = sorted(groups, key=lambda k: (-SEVERITIES.index(k[0]), k[1]))
    return [(sev, code, groups[(sev, code)]) for sev, code in order]


def render_text(reports: list[Report], max_per_rule: int = 20) -> str:
    out: list[str] = []
    for rep in reports:
        if rep.kind == "diff":
            out.append(f"xlsx-audit {__version__}  {rep.files[0]} -> {rep.files[1]}")
            out.append(f"old: {_stats_line(rep.stats.get('old', {}))}")
            out.append(f"new: {_stats_line(rep.stats.get('new', {}))}")
            noun = "change"
        else:
            out.append(f"xlsx-audit {__version__}  {rep.files[0]}")
            out.append(_stats_line(rep.stats))
            noun = "finding"
        counts = rep.counts
        if not rep.findings:
            out.append(f"No {noun}s.")
            out.append("")
            continue
        parts = [_plural(n, s) for s, n in counts.items() if n]
        out.append(f"{_plural(len(rep.findings), noun)}: " + ", ".join(parts))
        out.append("")
        for sev, code, items in _groups(rep.findings):
            out.append(f"{sev.upper():<8} {code} {ALL_RULES[code].slug} ({len(items)})")
            width = min(max(len(f.location) for f in items[:max_per_rule]), 28)
            for f in items[:max_per_rule]:
                out.append(f"  {f.location:<{width}}  {f.message}")
            if len(items) > max_per_rule:
                out.append(f"  ... and {len(items) - max_per_rule} more (use --max-per-rule 0 or --format json to see all)")
            out.append("")
    text = "\n".join(out).rstrip() + "\n"
    return text.encode("ascii", "replace").decode("ascii")


def render_json(reports: list[Report]) -> str:
    data = {
        "tool": "xlsx-audit",
        "version": __version__,
        "reports": [
            {
                "kind": r.kind,
                "files": r.files,
                "stats": r.stats,
                "counts": r.counts,
                "findings": [f.to_dict() for f in r.findings],
            }
            for r in reports
        ],
    }
    return json.dumps(data, indent=2, default=str) + "\n"


def _uri(path: str) -> str:
    p = Path(path)
    try:
        p = p.resolve().relative_to(Path.cwd().resolve())
    except ValueError:
        pass
    return str(p).replace(os.sep, "/")


SARIF_LEVEL = {"error": "error", "warning": "warning", "info": "note"}


def render_sarif(reports: list[Report]) -> str:
    used = sorted({f.code for r in reports for f in r.findings})
    rules = [
        {
            "id": code,
            "name": ALL_RULES[code].slug,
            "shortDescription": {"text": ALL_RULES[code].summary},
            "helpUri": f"{INFO_URI}#rules",
            "defaultConfiguration": {"level": SARIF_LEVEL[ALL_RULES[code].severity]},
        }
        for code in used
    ]
    results = []
    for r in reports:
        uri = _uri(r.files[-1])
        for f in r.findings:
            results.append(
                {
                    "ruleId": f.code,
                    "level": SARIF_LEVEL[f.severity],
                    "message": {"text": f"{f.location}: {f.message}"},
                    "locations": [
                        {
                            "physicalLocation": {"artifactLocation": {"uri": uri}, "region": {"startLine": 1}},
                            "logicalLocations": [{"fullyQualifiedName": f.location, "kind": "element"}],
                        }
                    ],
                    "partialFingerprints": {"xlsxAudit/v1": f"{f.code}:{f.location}:{f.message}"},
                }
            )
    sarif = {
        "$schema": "https://json.schemastore.org/sarif-2.1.0.json",
        "version": "2.1.0",
        "runs": [
            {
                "tool": {"driver": {"name": "xlsx-audit", "version": __version__, "informationUri": INFO_URI, "rules": rules}},
                "results": results,
            }
        ],
    }
    return json.dumps(sarif, indent=2, default=str) + "\n"


def _md_escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def render_markdown(reports: list[Report], max_rows: int = 200) -> str:
    out: list[str] = []
    for rep in reports:
        title = f"`{rep.files[0]}` -> `{rep.files[1]}`" if rep.kind == "diff" else f"`{rep.files[0]}`"
        out.append(f"### xlsx-audit: {title}")
        out.append("")
        counts = rep.counts
        if not rep.findings:
            out.append("No findings." if rep.kind == "lint" else "No changes.")
            out.append("")
            continue
        out.append(" | ".join(f"**{n}** {s}" for s, n in counts.items()))
        out.append("")
        out.append("| Severity | Rule | Where | What |")
        out.append("|---|---|---|---|")
        for f in rep.findings[:max_rows]:
            out.append(f"| {f.severity} | {f.code} {f.slug} | `{_md_escape(f.location)}` | {_md_escape(f.message)} |")
        if len(rep.findings) > max_rows:
            out.append("")
            out.append(f"... and {len(rep.findings) - max_rows} more.")
        out.append("")
    return ("\n".join(out).rstrip() + "\n").encode("ascii", "replace").decode("ascii")


RENDERERS = {"text": render_text, "json": render_json, "sarif": render_sarif, "markdown": render_markdown}

"""Command line: xlsx-audit lint | diff | rules."""
from __future__ import annotations

import argparse
import glob
import sys
from pathlib import Path

from . import __version__
from .checks import lint
from .diff import diff
from .model import DIFF_RULES, LINT_RULES, SEVERITIES, severity_rank
from .report import RENDERERS, Report, render_markdown, render_text
from .workbook import WorkbookError, load

EXIT_OK, EXIT_FINDINGS, EXIT_ERROR = 0, 1, 2


def _expand(paths: list[str]) -> list[str]:
    """Expand wildcards ourselves: Windows shells pass *.xlsx through as is."""
    out: list[str] = []
    for p in paths:
        if any(ch in p for ch in "*?[") and not Path(p).exists():
            matches = sorted(glob.glob(p, recursive=True))
            if not matches:
                raise WorkbookError(f"{p}: no files match")
            out.extend(matches)
        else:
            out.append(p)
    return out


def _codes(value: str | None) -> set[str] | None:
    if not value:
        return None
    return {v.strip() for v in value.split(",") if v.strip()}


def _common(p: argparse.ArgumentParser) -> None:
    p.add_argument("--format", "-f", choices=sorted(RENDERERS), default="text", help="output format (default: text)")
    p.add_argument("--output", "-o", metavar="PATH", help="write the report to a file instead of the screen")
    p.add_argument(
        "--fail-on",
        choices=["error", "warning", "info", "never"],
        default="error",
        help="exit with code 1 when a finding has this severity or higher (default: error)",
    )
    p.add_argument("--select", metavar="RULES", help="only run these rules, comma-separated (codes or names)")
    p.add_argument("--ignore", metavar="RULES", help="skip these rules, comma-separated (codes or names)")
    p.add_argument("--max-per-rule", type=int, default=20, metavar="N", help="text format: show at most N findings per rule; 0 shows all (default: 20)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlsx-audit",
        description="Audit Excel workbooks for broken references, external links, hard-coded overrides and formula changes.",
    )
    parser.add_argument("--version", action="version", version=f"xlsx-audit {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p_lint = sub.add_parser("lint", help="check one or more workbooks")
    p_lint.add_argument("files", nargs="+", metavar="FILE", help=".xlsx or .xlsm file")
    _common(p_lint)

    p_diff = sub.add_parser("diff", help="compare two versions of a workbook")
    p_diff.add_argument("old", metavar="OLD", help="earlier version")
    p_diff.add_argument("new", metavar="NEW", help="later version")
    p_diff.add_argument("--no-values", action="store_true", help="leave out changes to typed values (inputs)")
    _common(p_diff)

    sub.add_parser("rules", help="list all rules")
    return parser


def _write(text: str, output: str | None) -> None:
    if output:
        Path(output).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)


def _exit_code(reports: list[Report], fail_on: str) -> int:
    if fail_on == "never":
        return EXIT_OK
    floor = severity_rank(fail_on)
    hit = any(severity_rank(f.severity) >= floor for r in reports for f in r.findings)
    return EXIT_FINDINGS if hit else EXIT_OK


def _render(args, reports: list[Report]) -> str:
    if args.format == "text":
        return render_text(reports, args.max_per_rule if args.max_per_rule > 0 else 10**9)
    if args.format == "markdown":
        return render_markdown(reports)
    return RENDERERS[args.format](reports)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.command == "rules":
        lines = [f"{'Code':<6} {'Name':<28} {'Severity':<8} What it finds"]
        for rule in list(LINT_RULES.values()) + list(DIFF_RULES.values()):
            lines.append(f"{rule.code:<6} {rule.slug:<28} {rule.severity:<8} {rule.summary}")
        _write("\n".join(lines) + "\n", None)
        return EXIT_OK

    try:
        select, ignore = _codes(args.select), _codes(args.ignore)
        if args.command == "lint":
            reports = []
            for f in _expand(args.files):
                wb = load(f)
                reports.append(Report("lint", [f], lint(wb, select, ignore), wb.stats))
        else:
            old, new = load(args.old), load(args.new)
            result = diff(old, new, include_values=not args.no_values, lint_select=select, lint_ignore=ignore)
            reports = [Report("diff", [args.old, args.new], result.findings, result.stats)]
    except WorkbookError as exc:
        sys.stderr.write(f"xlsx-audit: {exc}\n")
        return EXIT_ERROR
    except ValueError as exc:  # unknown rule names
        sys.stderr.write(f"xlsx-audit: {exc}\n")
        return EXIT_ERROR

    _write(_render(args, reports), args.output)
    return _exit_code(reports, args.fail_on)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())

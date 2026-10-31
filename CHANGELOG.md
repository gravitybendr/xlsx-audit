# Changelog

## Unreleased

### Fixed

- Diff: references past the end of the data now move when rows or columns are inserted or deleted above them, as in Excel. With data in rows 2 to 21, inserting a row turns `=SUM(A2:A100)` into `=SUM(A2:A101)`, and 0.1.0 reported that as a changed formula (XD020). Rows typed into the empty rows below the data still move nothing. Found while making the README images; tests added.

### Documentation

- README rewritten: who it's for, what you get, examples with images of real output, a comparison with other tools, FAQ. Image and link URLs are absolute so they work on PyPI.
- Logo (`docs/brand/logo.svg`, PNGs) and a GitHub social preview image.
- Demo workbooks in `docs/demo/` and `scripts/make_images.py`, which runs xlsx-audit on them and renders the README images from its output.
- New PyPI summary and more keywords in `pyproject.toml`.

## 0.1.0 (30 October 2026)

First version.

- `xlsx-audit lint`: ten checks on one workbook (XA001-XA010): error values traced to their source, `#REF!` in formulas and names, links to other workbooks, numbers typed over formulas, formulas that break a pattern, volatile functions, hidden sheets, typed numbers inside formulas, and formulas that include their own cell.
- `xlsx-audit diff`: changes between two versions (XD001-XD041), with rows or columns lined up first, formulas compared after following moved references across sheets, renamed sheets matched by content, and lint findings reported as new or resolved.
- Text, JSON, SARIF 2.1.0 and Markdown output; CI exit codes with `--fail-on`; `--select` and `--ignore` by rule code or name; wildcards expanded on Windows.
- Python API: `audit()`, `compare()`, `load()`, `lint()`, `diff()`.
- Tested on Python 3.9 to 3.14 on Linux, and 3.12 on Windows and macOS; a 45,000-cell workbook lints in about 1 second and diffs in about 1.6 seconds.

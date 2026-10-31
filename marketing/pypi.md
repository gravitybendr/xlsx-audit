# PyPI listing

Project page: <https://pypi.org/project/xlsx-audit/>

PyPI builds the page from `pyproject.toml` and `README.md` **of the release**. Version 0.1.0 is live with the old README and summary; everything below shows up with the next release. PyPI has no icon field: the logo appears through the README.

## Name

`xlsx-audit` (the package name; PyPI shows it as the title).

## Summary (pyproject `description`, 179 characters, one line)

Lint and diff Excel workbooks: find broken references, numbers typed over formulas and links to other files, and see what changed between two versions. Command line, CI or Python.

## Long description

`README.md`. Its opening, as it reads on the page:

> **A linter and diff for Excel workbooks.** Find broken references, numbers typed over formulas and links to other files, and see exactly what changed between two versions. From the command line, in CI, or from Python. Your files never leave your machine.

All README images and links use absolute `https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/...` and `https://github.com/...` URLs, because relative paths break on PyPI. PyPI's renderer keeps `<img width height alt>`, `<sub>`, `<code>` and `<details>`, which the README uses; it strips `<picture>`, which the README doesn't use.

## Images (in README order)

All are made by `scripts/make_images.py` from the demo workbooks in `docs/demo/` and the CLI output saved in `docs/demo/output/`.

| File | Caption |
|---|---|
| `docs/images/hero.png` | The demo workbook `model_v1.xlsx` (made-up numbers), drawn from its cells with formulas shown. The highlighted cells and the list come from `xlsx-audit lint --format json` on that file. |
| `docs/images/lint-output.png` | `xlsx-audit lint model_v1.xlsx`: 11 findings, 3 errors, 5 warnings, 3 info. Severity colours are added in the image; the command prints plain ASCII text. |
| `docs/images/diff.png` | Old and new versions of the Sales sheet (columns A-D). The highlights and the list are the output of `xlsx-audit diff --format json`: an inserted row, a formula pasted over with a value, a fixed total and a renamed sheet. The 21 cells that only moved aren't reported. |
| `docs/images/markdown-report.png` | `xlsx-audit lint model_v1.xlsx --format markdown`, rendered as HTML, for pull-request comments and job summaries. |
| `docs/images/diff-output.png` | `xlsx-audit diff model_v1.xlsx model_v2.xlsx`: 6 changes. Not used in the README (the text is shown instead); available for posts and slides. |

## Keywords (pyproject `keywords`)

excel, xlsx, xlsm, spreadsheet, audit, formula, lint, linter, diff, excel diff, spreadsheet diff, financial model, model review, formula audit, hardcoded values, external links, openpyxl, sarif, ci

## Classifiers

Unchanged from 0.1.0: Beta; Console; Developers and Financial and Insurance Industry; OS Independent; Python 3.9 to 3.14; Topic :: Office/Business :: Financial :: Spreadsheet; Topic :: Software Development :: Quality Assurance.

## Short copy for sharing

- One line: *xlsx-audit: a linter and diff for Excel workbooks, for the command line, CI and Python.*
- Two lines: *Finds numbers typed over formulas, broken references and links to other files in .xlsx workbooks. Its diff follows inserted rows, so you see the six real changes, not every cell that moved.*
- Search phrases people use: "excel formula audit", "xlsx lint", "spreadsheet diff", "compare two excel files", "financial model review", "find hardcoded values in excel", "excel #REF! finder".

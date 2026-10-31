# <img src="https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/docs/brand/logo-128.png" width="40" height="40" alt=""> xlsx-audit

**A linter and diff for Excel workbooks.** Find broken references, numbers typed over formulas and links to other files, and see exactly what changed between two versions. From the command line, in CI, or from Python. Your files never leave your machine.

[![PyPI](https://img.shields.io/pypi/v/xlsx-audit)](https://pypi.org/project/xlsx-audit/)
[![Python](https://img.shields.io/pypi/pyversions/xlsx-audit)](https://pypi.org/project/xlsx-audit/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](https://github.com/gravitybendr/xlsx-audit/blob/main/LICENSE)

![xlsx-audit lint on a demo sales model: the Sales sheet with formulas shown, 7 flagged cells, a flagged hidden sheet and a list of 11 findings](https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/docs/images/hero.png)

<sub>The demo workbook <code>docs/demo/model_v1.xlsx</code> (made-up numbers), drawn from its cells with formulas shown. The highlighted cells and the list come from <code>xlsx-audit lint --format json</code> on that file. <code>scripts/make_images.py</code> makes the image.</sub>

## Who it's for

- **Analysts and model reviewers** checking a financial model before it goes to a client, a lender or a board.
- **Teams that keep workbooks in Git** and want a pull-request check, the way code has a linter.
- **Developers** whose apps generate or consume `.xlsx` files and need to test them.

## The problem

A workbook can look right and be wrong. Someone types `535` over one cell in a column of formulas, and the total still adds up. A formula picks up the price from the row below. A sheet that fed the model was deleted, and `#REF!` hides behind `IFERROR`. When the next version arrives, Excel can't tell you what changed, and a cell-by-cell compare drowns you in noise as soon as one row was inserted.

## What you get

- **Ten checks on one workbook**: broken references and names, error values traced to the cell they come from, links to other workbooks, numbers typed over formulas, formulas that break the pattern of their neighbours, formulas that include their own cell, hidden and very hidden sheets, volatile functions, and numbers typed inside formulas.
- **A diff that understands spreadsheets.** Rows or columns are lined up first and references are followed to where their cells moved, so an inserted row is one change, not hundreds. Formulas pasted over with values, changed formulas, renamed sheets, new links, and problems that are new or fixed are each reported once.
- **Output for people and for tools**: readable text, JSON, SARIF for GitHub code scanning, and Markdown for pull-request comments and job summaries.
- **Exit codes for CI**, with `--fail-on` to choose the level that fails a build.
- **A Python API** (`audit()`, `compare()`) for tests and scripts.
- **Read-only and offline.** It reads the file and reports. No network requests and no Excel needed, on Windows, macOS or Linux.

## Quick start

Python 3.9 or newer.

```
pip install xlsx-audit
xlsx-audit lint model.xlsx
xlsx-audit diff model_v1.xlsx model_v2.xlsx
```

Or run it without installing, with [uv](https://docs.astral.sh/uv/) or [pipx](https://pipx.pypa.io/):

```
uvx xlsx-audit lint model.xlsx
pipx run xlsx-audit lint model.xlsx
```

To try it on the two demo workbooks used on this page, clone the repo and run it in [`docs/demo/`](https://github.com/gravitybendr/xlsx-audit/tree/main/docs/demo):

```
git clone https://github.com/gravitybendr/xlsx-audit
cd xlsx-audit/docs/demo
xlsx-audit lint model_v1.xlsx
xlsx-audit diff model_v1.xlsx model_v2.xlsx
```

## Examples

All output below is what xlsx-audit 0.1.0 printed for the demo workbooks. The saved output is in [`docs/demo/output/`](https://github.com/gravitybendr/xlsx-audit/tree/main/docs/demo/output).

### Check a workbook

`model_v1.xlsx` is a small sales model with one of each common mistake:

![Terminal output of xlsx-audit lint model_v1.xlsx: 11 findings, 3 errors, 5 warnings, 3 info](https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/docs/images/lint-output.png)

<sub>Severity colours are added in the image; the command prints plain ASCII text.</sub>

<details>
<summary>The same output as text</summary>

```
$ xlsx-audit lint model_v1.xlsx
xlsx-audit 0.1.0  model_v1.xlsx
3 sheets, 32 cells, 10 formulas
11 findings: 3 errors, 5 warnings, 3 info

ERROR    XA002 broken-reference (1)
  Sales!D2  Formula refers to deleted cells: =#REF!*2

ERROR    XA008 broken-name (1)
  (workbook)  Defined name OldRange points to #REF!

ERROR    XA010 self-reference (1)
  Sales!C8  Formula includes its own cell (C2:C8), a circular reference: =SUM(C2:C8)

WARNING  XA003 external-link (2)
  (workbook)  Workbook links to another file: file:///C:/Finance/Prices.xlsx
  Sales!E2    Formula uses another workbook (file:///C:/Finance/Prices.xlsx): =[1]Prices!A1*1.15

WARNING  XA004 hardcoded-in-formula-range (1)
  Sales!C4  Typed value 535 inside a run of 4 matching formulas (expected =A4*B4*Inputs!$B$1)

WARNING  XA005 inconsistent-formula (1)
  Sales!C5  =A5*B6*Inputs!$B$1 breaks the pattern of the 4 matching formulas around it (expected =A5*B5*Inputs!$B$1)

WARNING  XA007 hidden-sheet (1)
  Notes  Sheet Notes is very hidden: it can only be shown with VBA

INFO     XA001 error-value (1)
  Sales!D3  Shows #REF! because it uses D2, which also shows an error

INFO     XA006 volatile-function (1)
  Sales!F2  Uses INDIRECT, which recalculates on every change and hides what it reads

INFO     XA009 hardcoded-number (1)
  Sales!E2  Typed number 1.15 inside =[1]Prices!A1*1.15
```

</details>

### Compare two versions

In `model_v2.xlsx`, someone inserted a row, pasted a value over a formula, fixed the total and renamed a sheet:

![Old and new versions of the Sales sheet side by side, with the inserted row, the overwritten formula and the changed total highlighted, and the 6 changes xlsx-audit reported](https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/docs/images/diff.png)

<sub>Drawn from the cells of both demo workbooks (columns A-D of the Sales sheet). The highlights and the list are the output of <code>xlsx-audit diff --format json</code>.</sub>

Row 2 was inserted, so 21 other cells on Sales moved down one row. None of them is reported; only the six real changes are:

```
$ xlsx-audit diff model_v1.xlsx model_v2.xlsx
xlsx-audit 0.1.0  model_v1.xlsx -> model_v2.xlsx
old: 3 sheets, 32 cells, 10 formulas
new: 3 sheets, 33 cells, 9 formulas
6 changes: 2 warnings, 4 info

WARNING  XD021 formula-overwritten (1)
  Sales!C4  Formula =A3*B3*Inputs!$B$1 was replaced by the value 42.8 (was C3 in the old version)

WARNING  XD040 new-issue (1)
  Sales!C4  New XA004 hardcoded-in-formula-range: Typed value 42.8 inside a run of 3 matching formulas (expected =A4*B4*Inputs!$B$1)

INFO     XD003 sheet-renamed (1)
  'Internal notes'  Sheet Notes renamed to Internal notes

INFO     XD010 rows-inserted (1)
  Sales  Row 2 inserted

INFO     XD020 formula-changed (1)
  Sales!C9  Formula changed: =SUM(C2:C8) -> =SUM(C3:C8) (was C8 in the old version)

INFO     XD041 resolved-issue (1)
  Sales!C8  Resolved XA010 self-reference: Formula includes its own cell (C2:C8), a circular reference: =SUM(C2:C8)
```

### Report in a pull request

`--format markdown` writes a table you can post as a pull-request comment or add to a GitHub Actions job summary:

![The Markdown report for model_v1.xlsx, rendered: a heading, the counts and a table of 11 findings](https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/docs/images/markdown-report.png)

<sub><code>xlsx-audit lint model_v1.xlsx --format markdown</code>, rendered as HTML.</sub>

## Usage

### Check workbooks

```
xlsx-audit lint model.xlsx                      # one file
xlsx-audit lint reports/*.xlsx                  # several
xlsx-audit lint model.xlsx --select XA004,XA005 # only some rules
xlsx-audit lint model.xlsx --ignore hardcoded-number,volatile-function
```

Rules can be named by code (`XA004`) or name (`hardcoded-in-formula-range`). `xlsx-audit rules` lists them all.

### Compare two versions

```
xlsx-audit diff model_v1.xlsx model_v2.xlsx
xlsx-audit diff old.xlsx new.xlsx --no-values   # only structure and formulas, not input changes
```

The diff lines up rows, or columns when that explains the changes better, before comparing cells. Formulas are compared after moving every reference to where its cells ended up, on any sheet. So `=SUM(C2:C8)` that Excel turned into `=SUM(C3:C9)` when you inserted a row is unchanged, and so is a formula on another sheet that points into the moved rows. Renamed sheets are matched by their content.

It also lints both versions and reports lint findings that are **new** in the second one (XD040) or **resolved** (XD041).

### Output formats

| `--format` | Use it for |
|---|---|
| `text` (default) | Reading in a terminal. Long lists are shortened; `--max-per-rule 0` shows everything. |
| `json` | Scripts and dashboards. Every finding, with sheet, cell, rule, severity and details. |
| `sarif` | GitHub code scanning and other SARIF viewers. |
| `markdown` | Pull-request comments and CI job summaries. |

`--output report.json` writes to a file instead of the screen.

### Exit codes, for CI

| Code | Meaning |
|---|---|
| 0 | Nothing at or above the `--fail-on` level (default: `error`) |
| 1 | At least one finding at or above `--fail-on` |
| 2 | The file couldn't be read, or an option was wrong |

`--fail-on warning` makes warnings fail the build too; `--fail-on never` always exits 0.

## Use it in GitHub Actions

Check every workbook in a pull request and show the results in the code scanning tab:

```yaml
name: Workbooks
on: [pull_request]
jobs:
  lint:
    runs-on: ubuntu-latest
    permissions:
      security-events: write
    steps:
      - uses: actions/checkout@v7
      - uses: actions/setup-python@v7
        with:
          python-version: "3.12"
      - run: pip install xlsx-audit
      - run: xlsx-audit lint models/*.xlsx --format sarif --output xlsx-audit.sarif --fail-on never
      - uses: github/codeql-action/upload-sarif@v4
        with:
          sarif_file: xlsx-audit.sarif
      - run: xlsx-audit lint models/*.xlsx --format markdown >> "$GITHUB_STEP_SUMMARY"
```

## Use it with pre-commit

```yaml
repos:
  - repo: https://github.com/gravitybendr/xlsx-audit
    rev: v0.1.0
    hooks:
      - id: xlsx-audit
```

## Use it from Python

```python
import xlsx_audit

for f in xlsx_audit.audit("model.xlsx"):
    print(f.severity, f.code, f.location, f.message)

for change in xlsx_audit.compare("model_v1.xlsx", "model_v2.xlsx"):
    print(change.code, change.location, change.old, "->", change.new)
```

Each `Finding` has `code`, `slug`, `severity`, `sheet`, `cell`, `message` and `detail` (for example the expected formula). For a diff it also has `old`, `new` and `old_cell`. `xlsx_audit.load()`, `lint()` and `diff()` give you the lower-level steps.

## Rules

### Checks on one workbook

| Code | Name | Severity | What it finds |
|---|---|---|---|
| XA001 | error-value | error, warning or info | A cell shows an Excel error. `#REF!`, `#NAME?` and `#NULL!` are errors; `#DIV/0!`, `#VALUE!` and `#NUM!` are warnings; `#N/A` is info. An error that only comes from another error cell is info and names its source. |
| XA002 | broken-reference | error | A formula contains `#REF!`: it points at cells that were deleted, even if `IFERROR` hides the result. |
| XA003 | external-link | warning | The workbook, a formula or a defined name depends on another workbook file. |
| XA004 | hardcoded-in-formula-range | warning | A typed number inside a run of cells that all use the same formula. Info when it sits at the end of the run. |
| XA005 | inconsistent-formula | warning | A formula that breaks the pattern of the identical formulas around it. The message shows the formula it was expected to have. |
| XA006 | volatile-function | info | `INDIRECT`, `OFFSET`, `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `CELL` or `INFO`. They recalculate on every change, and `INDIRECT` and `OFFSET` hide what they read. |
| XA007 | hidden-sheet | info, or warning if very hidden | A hidden sheet. Very hidden sheets can only be shown through VBA. |
| XA008 | broken-name | error | A defined name that points to `#REF!`. |
| XA009 | hardcoded-number | info | A typed number inside a formula, such as `=B2*1.07`. Common counts like 12, 100 and 365 are ignored. |
| XA010 | self-reference | error | A formula whose own cell is inside a range it reads, such as a total that includes itself. |

### Changes between two versions

| Code | Name | Severity |
|---|---|---|
| XD001, XD002, XD003 | sheet added, removed, renamed | info, warning, info |
| XD010, XD011 | rows inserted, deleted | info |
| XD012, XD013 | columns inserted, deleted | info |
| XD020 | formula changed | info |
| XD021 | formula overwritten with a value | warning |
| XD022, XD023 | formula added, removed | info, warning |
| XD024, XD025, XD026 | value changed, added, removed | info |
| XD030 | defined name changed | info |
| XD031, XD032 | link to another workbook added, removed | warning, info |
| XD040, XD041 | lint finding new, resolved | the finding's own severity, info |

## How the pattern checks work

Formulas are compared in relative (R1C1) form: `=A2*B2` in C2 and `=A3*B3` in C3 are the same formula copied down. In each unbroken run of cells down a column or across a row, the formula most cells share (at least three copies) is the pattern. Inside the stretch it covers, a typed number is flagged (XA004), and so is a formula that appears at most twice (XA005). A rare variant is likely a slip; a variant that repeats is likely a second block on purpose. Cells at the edges of a run, such as a header or a first row that starts a running total, aren't flagged.

## How it compares

Other tools do parts of this job well. Pick by where you work and what you need.

| Tool | Good at | Where xlsx-audit is different |
|---|---|---|
| Excel's own error checking (the green triangles, *Formulas > Error Checking*) | Flags inconsistent formulas and error values while you edit. Built in. | Runs without Excel, on many files at once, in CI. Also reports links to other workbooks, very hidden sheets and broken names, in one report per file. |
| Microsoft Inquire and Spreadsheet Compare | Workbook analysis and a visual compare of two files, from inside Office. | Those come with some Office editions on Windows only. xlsx-audit runs wherever Python does, including CI. |
| Commercial auditing tools, such as PerfectXL or Operis OAK | Deep, visual model review, with reports for sign-off. | xlsx-audit is free and scriptable, with fewer checks and no visual report. |
| Excel version control, such as xltrail or git-xl | Workbook history in Git, including VBA. | xlsx-audit doesn't diff VBA, but it lints each version and reports problems that are new or fixed. |
| openpyxl or pandas | Reading cells so you can write your own checks. | The checks, the diff and the report formats are already written and tested. |
| formulas or pycel | Calculating workbooks in Python. | xlsx-audit doesn't calculate. It checks structure and uses the results Excel saved. |

## FAQ

**Does it change my workbook?**
No. It only reads the file and never writes to it. With `--output` it writes the report file, nothing else.

**Do I need Excel installed?**
No. It reads the `.xlsx` file directly, on Windows, macOS or Linux.

**Does it calculate formulas?**
No. Error values (XA001) come from the results Excel stored when it last saved the file. Every other check reads the formulas themselves.

**How fast is it?**
On a synthetic workbook with 45,000 cells and 25,000 formulas, `lint` took about 1.1 seconds on a Windows PC, and `diff` against a copy with one inserted row took about 2.9 seconds.

**It flags something I did on purpose. How do I quiet it?**
Use `--ignore` with the rule code or name, or `--select` to run only the rules you want. There is no per-cell suppression or config file yet.

**Can I use it on `.xls` or `.xlsb` files?**
Not directly. Save them as `.xlsx` first.

**Is it free?**
Yes, under the MIT licence.

## Limits

- **It reads what Excel saved.** Error values (XA001) come from the results Excel stored when the file was last saved. Files written by other programs, including `openpyxl`, often have no stored results; every other check still works on them.
- **`.xlsx` and `.xlsm` only.** Save `.xls` and `.xlsb` files as `.xlsx` first. Password-protected workbooks can't be read.
- **Not checked:** VBA code, Power Query, pivot tables, charts and conditional formatting.
- **Circular references** are found when a formula's own cell is inside a range it reads, not across chains of cells.
- **The diff** lines up rows or columns for each sheet, not both. If one edit inserted rows and columns in the same sheet, some moved cells show up as changed.

## Privacy

Everything runs locally. xlsx-audit makes no network requests and sends nothing anywhere.

## Licence

MIT. See [LICENSE](https://github.com/gravitybendr/xlsx-audit/blob/main/LICENSE).

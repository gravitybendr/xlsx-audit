### xlsx-audit: `model_v1.xlsx`

**3** error | **5** warning | **3** info

| Severity | Rule | Where | What |
|---|---|---|---|
| error | XA008 broken-name | `(workbook)` | Defined name OldRange points to #REF! |
| error | XA002 broken-reference | `Sales!D2` | Formula refers to deleted cells: =#REF!*2 |
| error | XA010 self-reference | `Sales!C8` | Formula includes its own cell (C2:C8), a circular reference: =SUM(C2:C8) |
| warning | XA003 external-link | `(workbook)` | Workbook links to another file: file:///C:/Finance/Prices.xlsx |
| warning | XA003 external-link | `Sales!E2` | Formula uses another workbook (file:///C:/Finance/Prices.xlsx): =[1]Prices!A1*1.15 |
| warning | XA004 hardcoded-in-formula-range | `Sales!C4` | Typed value 535 inside a run of 4 matching formulas (expected =A4*B4*Inputs!$B$1) |
| warning | XA005 inconsistent-formula | `Sales!C5` | =A5*B6*Inputs!$B$1 breaks the pattern of the 4 matching formulas around it (expected =A5*B5*Inputs!$B$1) |
| warning | XA007 hidden-sheet | `Notes` | Sheet Notes is very hidden: it can only be shown with VBA |
| info | XA009 hardcoded-number | `Sales!E2` | Typed number 1.15 inside =[1]Prices!A1*1.15 |
| info | XA006 volatile-function | `Sales!F2` | Uses INDIRECT, which recalculates on every change and hides what it reads |
| info | XA001 error-value | `Sales!D3` | Shows #REF! because it uses D2, which also shows an error |

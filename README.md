# Physics Spreadsheet Check

A student preview of potential spreadsheet grades using seven physics lab rubrics.

Choose a lab and an Excel `.xlsx` workbook. The Python rubric grader runs locally in a browser worker using Pyodide and openpyxl 3.1.5. Files and results are not sent to Canvas or saved by this site. PDF files receive instructions to export the original spreadsheet as `.xlsx`.

## Hosting

GitHub Pages serves this repository's `main` branch from the repository root. All asset paths are relative; no build or server is needed. Internet access is needed to load the grading runtime and packages on first use.

## Rubrics

Lab 1, Lab 2, Lab 3, Labs 4–5, Labs 6–7, Rotation, and Springs. The scorer is copied from the instructor notebook, including the nearby-header policy: within two columns, a 0.25-point deduction per affected table, capped at 1 point per submission. Merged headings spanning the data receive full placement credit.

Scores are previews. Class instructions and instructor review can affect final grades. Future rubric edits must be synchronized with `grader.py`.

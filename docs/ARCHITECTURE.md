# Architecture — lele_auto

> Project structure, technical choices and their rationale. Read before any structural change.

## Overview
Stock-sheet proposals for Lele (pastry chef), delivered as xlsx files, plus an internal site that
previews them with an Excel-like rendering for Pascal's review.

## Folder / module map
- `tools/build_proposals.py` — generates the 10 xlsx proposals into `share/out/propositions_stock/`.
- `tools/build_proposals_2.py` — generates proposals 11-20 (reuses the helpers of build_proposals.py).
- `tools/verify_proposals.py` — evaluates every formula with the `formulas` engine, reports errors and sample values.
- `tools/render_site.py` — renders each xlsx to `site/propositions/NN.html` (+ `site/index.html`, `site/files/`).
- `tools/.venv/` — local venv with `formulas` + `openpyxl` (run the verify/render scripts with it).
- `site/assets/` — shared viewer CSS/JS (sheet tabs, name box, formula bar).

## Tech choices & rationale
- No LibreOffice Calc on lili: formulas are evaluated with the Python `formulas` engine.
- Conditional formatting is evaluated by copying each rule, translated per cell, into a helper sheet
  of a temporary workbook copy, then computing it with the same engine.
- openpyxl cannot read charts back: charts shown in the viewer are declared in `CHARTS` in render_site.py.
- Pipeline after any change: build_proposals(_2).py -> verify_proposals.py -> render_site.py.
- The `formulas` engine mishandles the `"<>"` (non-blank) criterion in COUNTIFS: use `">=0"` or SUMPRODUCT instead.

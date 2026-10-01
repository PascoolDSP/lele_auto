# Decisions log — lele_auto

> One entry per important decision. Format: `YYYY-MM-DD — decision — reason`.
> Append new decisions; never rewrite history.

- 2026-09-17 — Project bootstrapped with newproj — standard lili layout (served dir, share/, docs/, stable port).
- 2026-09-24 — Deliver spreadsheets (xlsx, Google Sheets compatible), not an app — it is what Lele asked for and her reference is a 6 EUR Sheets template; volunteer work, keep scope minimal.
- 2026-09-24 — No recipe-based stock deduction — adds setup cost and silent drift; out of scope for what she asked.
- 2026-09-24 — Formulas verified with the Python `formulas` engine (tools/.venv) — LibreOffice Calc is not installed on lili.
- 2026-09-24 — No cross-sheet references inside conditional formatting — Google Sheets does not support them.
- 2026-09-25 — Batch 2 stays within Lele's brief (no recipe deduction, no Apps Script, no Google Forms) — scripts and buttons do not run in the Sheets mobile app, and xlsx must work in both Excel and Sheets.
- 2026-09-25 — No FILTER/SPARKLINE: lists use INDEX/MATCH on a rank helper, trends use REPT bars — portable across Excel, Sheets and LibreOffice.

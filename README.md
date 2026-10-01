# lele_auto

Stock-tracking spreadsheets (Excel / Google Sheets) for a pastry chef: products and suppliers,
a month split into 4 weeks, stock counted / quantity ordered / quantity received, automatic totals.

- `tools/build_proposals.py`, `tools/build_proposals_2.py` — generate the 20 xlsx proposals.
- `tools/verify_proposals.py` — evaluates every formula (Python `formulas` engine) and reports errors.
- `tools/render_site.py` — renders each workbook as an Excel-like HTML preview in `site/`.
- `site/files/` — the generated workbooks.

See `docs/` for architecture, decisions and roadmap.

"""Render the xlsx proposals as static HTML pages that mimic Excel (for Pascal's review only).

Values come from the `formulas` engine (tools/.venv) since LibreOffice Calc is not installed.
Conditional formatting is evaluated by writing each rule, translated per cell, into a helper
sheet of a temporary copy of the workbook.
"""
import datetime as dt
import html
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path

import formulas
from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "share" / "out" / "propositions_stock"
SITE = ROOT / "site"
HELPER = "__cf"

# openpyxl does not read charts back: declare the one chart we generate (see build_proposals.p07)
CHARTS = {
    "07": [{"sheet": "Tableau de bord", "anchor": "A23", "title": "Dépenses du mois (€)",
            "cats": "B11:B20", "vals": "E11:E20", "color": "#3D3D3D", "w": 529, "h": 283}],
}


# ---------------------------------------------------------------- evaluation
def dxf_css(dxf):
    """CSS for a conditional-format differential style (fill, font colour, strike, bold)."""
    parts = []
    if dxf.fill is not None:
        c = dxf.fill.fgColor.rgb if dxf.fill.fgColor is not None else None
        if not isinstance(c, str) or c.endswith("000000"):
            c = dxf.fill.bgColor.rgb if dxf.fill.bgColor is not None else None
        if isinstance(c, str) and not c.endswith("000000"):
            parts.append(f"background:#{c[-6:]}")
    if dxf.font is not None:
        if dxf.font.color is not None and isinstance(dxf.font.color.rgb, str):
            parts.append(f"color:#{dxf.font.color.rgb[-6:]}")
        if dxf.font.strike:
            parts.append("text-decoration:line-through")
        if dxf.font.b:
            parts.append("font-weight:700")
    return ";".join(parts)


def qualify(formula, sheet):
    """Prefix every unqualified cell reference with the sheet name."""
    tok = Tokenizer(formula)
    out = []
    for t in tok.items:
        v = t.value
        if t.type == Token.OPERAND and t.subtype == Token.RANGE and "!" not in v:
            v = f"'{sheet}'!{v}"
        out.append(v)
    return "=" + "".join(out)


def evaluate(path):
    """Return (values, cf_results): values[(sheet, 'A1')] and cf_results[(sheet, 'A1')] = fill hex."""
    wb = load_workbook(path)
    helper = wb.create_sheet(HELPER)
    jobs = []  # (sheet, coord, helper_row, fill)
    row = 1
    for ws in wb.worksheets:
        if ws.title == HELPER:
            continue
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                if not rule.formula or rule.dxf is None:
                    continue
                fill = dxf_css(rule.dxf)
                if not fill:
                    continue
                for rng in cf.sqref.ranges:
                    c1, r1, c2, r2 = range_boundaries(rng.coord)
                    origin = f"{get_column_letter(c1)}{r1}"
                    for r in range(r1, r2 + 1):
                        for c in range(c1, c2 + 1):
                            target = f"{get_column_letter(c)}{r}"
                            f = Translator("=" + rule.formula[0], origin=origin).translate_formula(target)
                            helper[f"A{row}"] = qualify(f, ws.title)
                            jobs.append((ws.title, target, row, fill))
                            row += 1
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / path.name
        wb.save(tmp_path)
        sol = formulas.ExcelModel().loads(str(tmp_path)).finish().calculate()
    values = {}
    rx = re.compile(r"^'?\[[^\]]+\](.+?)'?!([A-Z]+[0-9]+)$")
    for key, rng in sol.items():
        m = rx.match(key)
        if not m:
            continue
        try:
            v = rng.value[0][0]
        except Exception:
            continue
        values[(m.group(1).upper(), m.group(2))] = v
    cf_results = {}
    for sheet, coord, hrow, fill in jobs:
        v = values.get((HELPER.upper(), f"A{hrow}"))
        if str(v).upper() == "TRUE" or (is_number(v) and float(v) != 0):
            cf_results.setdefault((sheet, coord), fill)  # first matching rule wins
    return values, cf_results


# ---------------------------------------------------------------- formatting
def fr_number(x, decimals=None):
    if decimals is None:
        s = f"{x:.10g}"
        if "e" in s:
            s = f"{x:.6f}".rstrip("0").rstrip(".")
    else:
        s = f"{x:,.{decimals}f}".replace(",", " ")
    return s.replace(".", ",")


def fmt_value(v, number_format):
    if v is None:
        return ""
    s = str(v)
    if s in ("#EMPTY", "EMPTY"):
        return ""
    if isinstance(v, bool):
        return "VRAI" if v else "FAUX"
    if isinstance(v, str):
        return v
    try:
        x = float(v)
    except (TypeError, ValueError):
        return s
    nf = number_format or "General"
    if "YYYY" in nf.upper():
        if x <= 0:
            return ""
        d = dt.date(1899, 12, 30) + dt.timedelta(days=int(x))
        return d.strftime("%d/%m/%Y")
    if nf.startswith("#,##0.00"):
        return fr_number(x, 2) + (" €" if "€" in nf else "")
    if nf == "0":
        return fr_number(round(x), 0)
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return fr_number(x)


def is_number(v):
    if isinstance(v, bool) or v is None:
        return False
    try:
        float(v)
        return not isinstance(v, str)
    except (TypeError, ValueError):
        return False


def rgb(color):
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return None
    return "#" + color.rgb[-6:]


# ---------------------------------------------------------------- rendering
class StylePool:
    def __init__(self):
        self.pool = {}

    def cls(self, css):
        if css not in self.pool:
            self.pool[css] = f"s{len(self.pool)}"
        return self.pool[css]

    def css(self):
        return "\n".join(f".{c}{{{s}}}" for s, c in self.pool.items())


def col_px(ws, idx):
    d = ws.column_dimensions.get(get_column_letter(idx))
    if d is not None and d.hidden:
        return 0
    w = d.width if d is not None and d.width else 8.43
    return int(w * 7 + 5)


def row_px(ws, r):
    d = ws.row_dimensions.get(r)
    h = d.height if d is not None and d.height else 15
    return round(h * 4 / 3)


def cell_css(c, fill_override=None):
    parts = []
    f = c.font
    if f is not None:
        parts.append(f"font-family:{f.name or 'Arial'},sans-serif")
        parts.append(f"font-size:{f.sz or 11}pt")
        if f.b:
            parts.append("font-weight:700")
        if f.i:
            parts.append("font-style:italic")
        col = rgb(f.color)
        if col:
            parts.append(f"color:{col}")
    fill = fill_override
    if not fill and c.fill is not None and c.fill.fill_type == "solid":
        fill = rgb(c.fill.fgColor)
    if fill:
        parts.append(f"background:{fill if fill.startswith('#') else '#' + fill}")
    for side in ("left", "right", "top", "bottom"):
        b = getattr(c.border, side)
        if b is not None and b.style:
            parts.append(f"border-{side}:1px solid {rgb(b.color) or '#000'}")
    a = c.alignment
    h = a.horizontal if a is not None else None
    if h in ("center", "centerContinuous"):
        parts.append("text-align:center")
    elif h == "right":
        parts.append("text-align:right")
    elif h == "left":
        parts.append("text-align:left")
    v = a.vertical if a is not None else None
    parts.append({"center": "vertical-align:middle", "top": "vertical-align:top"}.get(v, "vertical-align:bottom"))
    if a is not None and a.wrap_text:
        parts.append("white-space:normal")
    return ";".join(parts)


def render_sheet(ws, values, cfr, pool, charts):
    max_r, max_c = ws.max_row, ws.max_column
    max_c = max(max_c, 8)
    max_r = max(max_r + 2, 30)
    merged = {}
    covered = set()
    for m in ws.merged_cells.ranges:
        c1, r1, c2, r2 = m.bounds
        merged[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) != (r1, c1):
                    covered.add((r, c))
    fr_row, fr_col = 1, 1
    if ws.freeze_panes:
        fc, frr, _, _ = range_boundaries(ws.freeze_panes + ":" + ws.freeze_panes)
        fr_row, fr_col = frr, fc
    widths = [col_px(ws, c) for c in range(1, max_c + 1)]
    heights = {r: row_px(ws, r) for r in range(1, max_r + 1)}
    RH_W, CH_H = 42, 22
    left_off = {c: RH_W + sum(widths[:c - 1]) for c in range(1, fr_col)}
    top_off = {r: CH_H + sum(heights[k] for k in range(1, r)) for r in range(1, fr_row)}
    sheet_key = ws.title.upper()

    out = ['<table class="xl" style="width:%dpx"><colgroup><col style="width:%dpx">' % (RH_W + sum(widths), RH_W)]
    out += [f'<col style="width:{w}px">' for w in widths]
    out.append('</colgroup><thead><tr><th class="corner"></th>')
    for c in range(1, max_c + 1):
        extra = f' style="left:{left_off[c]}px" class="fc"' if c in left_off else ""
        label = "" if widths[c - 1] == 0 else get_column_letter(c)
        out.append(f"<th{extra}>{label}</th>")
    out.append("</tr></thead><tbody>")
    for r in range(1, max_r + 1):
        frozen_row = r in top_off
        out.append(f'<tr style="height:{heights[r]}px">')
        st = f' style="top:{top_off[r]}px"' if frozen_row else ""
        out.append(f'<th class="rh{" fr" if frozen_row else ""}"{st}>{r}</th>')
        for c in range(1, max_c + 1):
            if (r, c) in covered:
                continue
            if widths[c - 1] == 0:
                out.append('<td class="hid"></td>')
                continue
            cell = ws.cell(row=r, column=c)
            coord = cell.coordinate
            raw = cell.value
            formula = raw if isinstance(raw, str) and raw.startswith("=") else None
            v = values.get((sheet_key, coord)) if formula else raw
            text = fmt_value(v, cell.number_format)
            css = cell_css(cell)
            if (ws.title, coord) in cfr:
                css += ";" + cfr[(ws.title, coord)]
            align_set = cell.alignment is not None and cell.alignment.horizontal
            if not align_set and is_number(v):
                css += ";text-align:right"
            classes = [pool.cls(css)]
            if "background:" in css:
                classes.append("f")
            wrap = cell.alignment is not None and cell.alignment.wrap_text
            if text and not wrap and (r, c) not in merged and not is_number(v):
                nxt = ws.cell(row=r, column=c + 1) if c < max_c else None
                if nxt is None or (nxt.value in (None, "") and (r, c + 1) not in covered):
                    classes.append("ov")
            attrs = []
            span = merged.get((r, c))
            if span:
                if span[0] > 1:
                    attrs.append(f'rowspan="{span[0]}"')
                if span[1] > 1:
                    attrs.append(f'colspan="{span[1]}"')
            sticky = []
            if frozen_row:
                classes.append("fr")
                sticky.append(f"top:{top_off[r]}px")
            if c in left_off:
                classes.append("fc")
                sticky.append(f"left:{left_off[c]}px")
            if sticky:
                attrs.append(f'style="{";".join(sticky)}"')
            if formula:
                attrs.append(f'data-f="{html.escape(formula, quote=True)}"')
            if cell.comment is not None:
                classes.append("cm")
                attrs.append(f'title="{html.escape(cell.comment.text, quote=True)}"')
            attrs.insert(0, f'class="{" ".join(classes)}"')
            out.append(f'<td {" ".join(attrs)}>{html.escape(text)}</td>')
        out.append("</tr>")
    out.append("</tbody></table>")
    for ch in charts:
        out.append(render_chart(ws, ch, values, widths, heights, RH_W, CH_H))
    return "".join(out)


def render_chart(ws, ch, values, widths, heights, rh_w, ch_h):
    c1, r1, _, _ = range_boundaries(ch["anchor"] + ":" + ch["anchor"])
    x = rh_w + sum(widths[:c1 - 1])
    y = ch_h + sum(heights[k] for k in range(1, r1))
    key = ws.title.upper()
    cc1, cr1, _, cr2 = range_boundaries(ch["cats"])
    vc1, vr1, _, vr2 = range_boundaries(ch["vals"])
    rows = []
    for k in range(cr2 - cr1 + 1):
        cat = values.get((key, f"{get_column_letter(cc1)}{cr1 + k}"), "")
        val = values.get((key, f"{get_column_letter(vc1)}{vr1 + k}"), "")
        if is_number(val) and str(cat) not in ("", "#EMPTY"):
            rows.append((str(cat), float(val)))
    w, h = ch["w"], ch["h"]
    pad_l, pad_t, pad_r, pad_b = 130, 40, 20, 24
    vmax = max([v for _, v in rows] + [1])
    bar_h = (h - pad_t - pad_b) / max(len(rows), 1)
    svg = [f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" font-family="Calibri,Arial,sans-serif">',
           f'<rect width="{w}" height="{h}" fill="#fff" stroke="#d9d9d9"/>',
           f'<text x="{w / 2}" y="24" text-anchor="middle" font-size="14" fill="#595959">{html.escape(ch["title"])}</text>']
    for i in range(5):
        gx = pad_l + (w - pad_l - pad_r) * i / 4
        svg.append(f'<line x1="{gx}" y1="{pad_t}" x2="{gx}" y2="{h - pad_b}" stroke="#e5e5e5"/>')
        svg.append(f'<text x="{gx}" y="{h - 8}" text-anchor="middle" font-size="10" fill="#595959">{fr_number(vmax * i / 4, 0)}</text>')
    for i, (cat, v) in enumerate(rows):
        by = pad_t + i * bar_h + bar_h * 0.18
        bw = (w - pad_l - pad_r) * v / vmax
        svg.append(f'<rect x="{pad_l}" y="{by:.1f}" width="{bw:.1f}" height="{bar_h * 0.64:.1f}" fill="{ch["color"]}"/>')
        svg.append(f'<text x="{pad_l - 6}" y="{by + bar_h * 0.45:.1f}" text-anchor="end" font-size="10" fill="#595959">{html.escape(cat)}</text>')
    svg.append("</svg>")
    return f'<div class="chart" style="left:{x}px;top:{y}px">{"".join(svg)}</div>'


# ---------------------------------------------------------------- pages
def page(title, body, head_extra="", depth=0):
    up = "../" * depth
    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>{html.escape(title)}</title>
<meta name="description" content="Aperçu des propositions de tableaux de stock pour Lele">
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' rx='12' fill='%23217346'/><path d='M18 18h28v28H18z' fill='none' stroke='white' stroke-width='4'/><path d='M18 32h28M32 18v28' stroke='white' stroke-width='4'/></svg>">
<link rel="stylesheet" href="{up}assets/app.css">
{head_extra}
</head>
<body>
{body}
<script src="{up}smooth-scroll.js" data-auto></script>
</body>
</html>
"""


def idea_of(wb):
    ws = wb["Mode d'emploi"]
    return ws["A1"].value, ws["B4"].value


def build():
    (SITE / "propositions").mkdir(parents=True, exist_ok=True)
    (SITE / "files").mkdir(parents=True, exist_ok=True)
    files = sorted(SRC.glob("*.xlsx"))
    only = sys.argv[1:]
    cards = []
    for i, path in enumerate(files):
        num = path.name[:2]
        wb = load_workbook(path)
        name, idea = idea_of(wb)
        cards.append((num, name, idea, path.name))
        if only and num not in only:
            continue
        shutil.copy(path, SITE / "files" / path.name)
        values, cfr = evaluate(path)
        pool = StylePool()
        sheets_html, tabs = [], []
        for k, ws in enumerate(wb.worksheets):
            charts = [c for c in CHARTS.get(num, []) if c["sheet"] == ws.title]
            sheets_html.append(
                f'<section class="sheet" data-i="{k}"{"" if k == 0 else " hidden"}>'
                f'<div class="grid" data-lenis-prevent tabindex="0" aria-label="Feuille {html.escape(ws.title)}">'
                f'{render_sheet(ws, values, cfr, pool, charts)}</div></section>')
            tabs.append(f'<button class="tab{" on" if k == 0 else ""}" data-i="{k}" role="tab" '
                        f'aria-selected="{"true" if k == 0 else "false"}">{html.escape(ws.title)}</button>')
        prev_n = files[i - 1].name[:2] if i > 0 else None
        next_n = files[i + 1].name[:2] if i + 1 < len(files) else None
        nav = (f'<a class="nav" href="{prev_n}.html" aria-label="Proposition précédente">&#8249;</a>' if prev_n else '<span class="nav off">&#8249;</span>') + \
              (f'<a class="nav" href="{next_n}.html" aria-label="Proposition suivante">&#8250;</a>' if next_n else '<span class="nav off">&#8250;</span>')
        body = f"""<div class="app">
<header class="titlebar">
  <a class="home" href="../index.html" aria-label="Retour aux propositions">&#8592; Propositions</a>
  <span class="fname">{html.escape(path.name)}</span>
  <span class="tb-right">{nav}<a class="dl" href="../files/{html.escape(path.name)}" download>Télécharger</a></span>
</header>
<div class="fbar"><span class="namebox" id="namebox">A1</span><span class="fx">fx</span><span class="fcontent" id="fcontent"></span></div>
<main class="sheets">{"".join(sheets_html)}</main>
<nav class="tabs" role="tablist">{"".join(tabs)}</nav>
<footer class="status"><span>Prêt</span><span class="note">Aperçu fidèle, valeurs calculées — les cases ne sont pas modifiables ici</span></footer>
</div>
<script src="../assets/viewer.js"></script>"""
        (SITE / "propositions" / f"{num}.html").write_text(
            page(f"{name} — aperçu", body, f"<style>{pool.css()}</style>", depth=1), encoding="utf-8")
        print("rendered", num, len(values), "values,", len(cfr), "cf hits")
    def card_html(n, name, idea, fname):
        return (f'<a class="card" href="propositions/{n}.html"><span class="num">{n}</span>'
                f'<span class="cname">{html.escape(name.split("·", 1)[-1].strip())}</span>'
                f'<span class="idea">{html.escape(idea)}</span>'
                f'<span class="file">{html.escape(fname)}</span></a>')

    series = [("Première série", "Variantes de mise en page de sa demande.", [c for c in cards if int(c[0]) <= 10]),
              ("Deuxième série", "Inspirée des modèles existants : cadencier, ordre de la réserve, fiche papier, DLC, pertes…",
               [c for c in cards if int(c[0]) > 10])]
    grid = "".join(
        f'<section class="series"><h2>{t}</h2><p class="sdesc">{d}</p><div class="cards">'
        + "".join(card_html(*c) for c in cs) + "</div></section>"
        for t, d, cs in series if cs)
    body = f"""<div class="index">
<header class="ihead">
  <p class="kicker">lele_auto · usage interne</p>
  <h1>Propositions de tableaux de stock</h1>
  <p class="lead">{len(cards)} variantes de la même demande : produits et fournisseurs, 4 semaines par mois, stock / commandé / reçu, totaux automatiques. Chaque aperçu reproduit le rendu Excel, avec les valeurs calculées.</p>
</header>
<main>{grid}</main>
</div>"""
    (SITE / "index.html").write_text(page("Propositions de tableaux de stock — Lele", body), encoding="utf-8")


if __name__ == "__main__":
    build()

"""Generate the 10 stock-sheet proposals for Lele (Excel / Google Sheets compatible).

Every proposal implements the same brief (products + suppliers, 4 weeks per month,
stock counted / quantity ordered / quantity received, automatic totals) with a
different layout, so Lele can pick one and give feedback.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parent.parent / "share" / "out" / "propositions_stock"
N = 40            # product rows available
FONT = "Arial"
INPUT_FILL = PatternFill("solid", fgColor="FFF6D5")
RED_FILL = PatternFill("solid", fgColor="F8C9C4")
ORANGE_FILL = PatternFill("solid", fgColor="FBDDB0")
GREEN_FILL = PatternFill("solid", fgColor="D3EBD0")
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
EXAMPLE_FONT = Font(name=FONT, color="1F4FD1")
MONTHS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
          "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# name, supplier, unit, min stock, target stock, unit price (EUR) — example data
PRODUCTS = [
    ("Farine T55", "Metro", "kg", 5, 15, 0.9),
    ("Farine T45", "Metro", "kg", 3, 8, 1.1),
    ("Sucre semoule", "Metro", "kg", 5, 15, 1.2),
    ("Sucre glace", "Metro", "kg", 2, 4, 1.8),
    ("Beurre doux", "Crèmerie", "kg", 4, 8, 8.5),
    ("Crème liquide 35%", "Crèmerie", "L", 6, 15, 4.2),
    ("Lait entier", "Crèmerie", "L", 5, 10, 1.1),
    ("Oeufs", "Crèmerie", "pièce", 60, 150, 0.25),
    ("Chocolat blanc", "Transgourmet", "kg", 2, 4, 14.0),
    ("Chocolat noir 70%", "Transgourmet", "kg", 2, 4, 13.0),
    ("Levure chimique", "Transgourmet", "kg", 0.5, 1, 6.0),
    ("Vanille (gousse)", "Transgourmet", "pièce", 5, 10, 2.5),
]
# (stock counted, ordered, received) for week 1 and week 2 — example data
WEEK_DATA = [
    [(8, 10, 7), (2, 5, 5), (6, 10, 10), (1, 3, 3), (3, 5, 5), (4, 12, 12),
     (3, 6, 6), (40, 90, 90), (1.5, 3, 2), (2, 2, 2), (0.4, 1, 1), (6, 0, 0)],
    [(5, 10, 10), (3, 3, 3), (7, 5, 5), (2, 0, 0), (4, 5, 5), (5, 10, 10),
     (4, 6, 6), (50, 60, 60), (2, 2, 2), (2, 2, 2), (0.6, 0, 0), (6, 0, 0)],
]
SUPPLIERS = [("Metro", "Commercial Metro", "01 23 45 67 89", "vendredi", "lundi"),
             ("Crèmerie", "Crèmerie du quartier", "01 98 76 54 32", "vendredi", "lundi"),
             ("Transgourmet", "Transgourmet", "01 11 22 33 44", "jeudi", "lundi")]


# ---------------------------------------------------------------- helpers
def lighten(hex_color, ratio):
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    mix = lambda c: int(c + (255 - c) * ratio)
    return f"{mix(r):02X}{mix(g):02X}{mix(b):02X}"


class Theme:
    def __init__(self, main):
        self.main = main
        self.fill = PatternFill("solid", fgColor=main)
        self.soft = PatternFill("solid", fgColor=lighten(main, 0.80))
        self.softer = PatternFill("solid", fgColor=lighten(main, 0.92))


def cell(ws, ref, value=None, bold=False, color="000000", size=10, fill=None,
         align="left", border=True, fmt=None, wrap=False, font=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font or Font(name=FONT, bold=bold, color=color, size=size)
    if fill:
        c.fill = fill
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        c.border = BORDER
    if fmt:
        c.number_format = fmt
    return c


def header(ws, ref, text, theme, span_to=None, soft=False):
    if span_to:
        ws.merge_cells(f"{ref}:{span_to}")
    fill = theme.soft if soft else theme.fill
    color = "000000" if soft else "FFFFFF"
    cell(ws, ref, text, bold=True, color=color, fill=fill, align="center", wrap=True)
    if span_to:
        for row in ws[f"{ref}:{span_to}"]:
            for c in row:
                c.fill = fill
                c.border = BORDER


def title(ws, text, subtitle, theme, width_cols=8):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, bold=True, size=16, color=theme.main)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, italic=True, size=9, color="666666")
    ws.row_dimensions[1].height = 26


def inp(ws, ref, value=None, example=False, fmt="General"):
    c = cell(ws, ref, value, fill=INPUT_FILL, align="center", fmt=fmt)
    if example and value is not None:
        c.font = EXAMPLE_FONT
    return c


def calc(ws, ref, formula, fmt="General", bold=False, fill=None):
    return cell(ws, ref, formula, align="center", fmt=fmt, bold=bold, fill=fill)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def nonneg_validation(ws, rng):
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                        allow_blank=True, showErrorMessage=True,
                        error="Entre un nombre positif (ou laisse vide).",
                        errorTitle="Valeur invalide")
    ws.add_data_validation(dv)
    dv.add(rng)


def negative_red(ws, rng, first_ref):
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f"AND(ISNUMBER({first_ref}),{first_ref}<0)"], fill=RED_FILL))


def guide(wb, theme, name, idea, steps, extra_tabs):
    ws = wb.active
    ws.title = "Mode d'emploi"
    ws.sheet_view.showGridLines = False
    title(ws, name, "Proposition de tableau de gestion des stocks — Excel et Google Sheets", theme)
    rows = [("L'idée", idea), ("", "")]
    rows += [("Comment faire", steps[0])] + [("", s) for s in steps[1:]]
    rows += [("", ""), ("Les onglets", extra_tabs[0])] + [("", t) for t in extra_tabs[1:]]
    rows += [("", ""),
             ("Couleurs", "Cases jaunes = à remplir. Cases blanches = calculées toutes seules, ne pas y toucher."),
             ("", "Chiffres en bleu = exemple pour montrer le fonctionnement, à effacer."),
             ("", "Rouge = il manque quelque chose (livraison incomplète, stock trop bas)."),
             ("", ""),
             ("Google Sheets", "Fichier > Importer > Importer ce fichier. Tout fonctionne pareil, y compris sur téléphone.")]
    r = 4
    for label, text in rows:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(name=FONT, bold=True, color=theme.main)
        ws[f"B{r}"] = text
        ws[f"B{r}"].font = Font(name=FONT)
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    widths(ws, {"A": 16, "B": 100})
    ws["A4"].alignment = Alignment(vertical="top")


def products_sheet(wb, theme, extra=()):
    """extra: subset of ('mini', 'cible', 'prix')."""
    ws = wb.create_sheet("Produits")
    ws.sheet_view.showGridLines = False
    title(ws, "Mes produits", "Une ligne par produit. Tout le reste du fichier se remplit à partir de cette liste.", theme)
    cols = [("Produit", 26), ("Fournisseur", 18), ("Unité", 9)]
    labels = {"mini": ("Stock minimum", 13), "cible": ("Stock idéal", 12), "prix": ("Prix unitaire (€)", 14)}
    cols += [labels[e] for e in extra]
    for i, (h, w) in enumerate(cols, start=1):
        header(ws, f"{L(i)}4", h, theme)
        ws.column_dimensions[L(i)].width = w
    for i in range(N):
        r = 5 + i
        p = PRODUCTS[i] if i < len(PRODUCTS) else None
        inp(ws, f"A{r}", p[0] if p else None, example=True).alignment = Alignment(horizontal="left", vertical="center")
        inp(ws, f"B{r}", p[1] if p else None, example=True).alignment = Alignment(horizontal="left", vertical="center")
        inp(ws, f"C{r}", p[2] if p else None, example=True)
        for j, e in enumerate(extra):
            val = {"mini": 3, "cible": 4, "prix": 5}[e]
            fmt = '#,##0.00 "€"' if e == "prix" else "General"
            inp(ws, f"{L(4 + j)}{r}", p[val] if p else None, example=True, fmt=fmt)
    ws.freeze_panes = "A5"
    unit_dv = DataValidation(type="list", formula1='"kg,g,L,pièce,paquet,boîte,sac"', allow_blank=True)
    ws.add_data_validation(unit_dv)
    unit_dv.add(f"C5:C{N + 4}")
    if extra:
        nonneg_validation(ws, f"D5:{L(3 + len(extra))}{N + 4}")
    return ws


def product_ref(col, i):
    """Formula pulling column `col` of product i from the Produits sheet."""
    return f'=IF(Produits!$A${5 + i}="","",Produits!${col}${5 + i})'


def product_columns(ws, first_row, theme, with_supplier=True, with_unit=True):
    """Write product / supplier / unit link columns starting at A."""
    cols = ["A"] + (["B"] if with_supplier else []) + (["C"] if with_unit else [])
    for i in range(N):
        r = first_row + i
        cell(ws, f"A{r}", product_ref("A", i), bold=True)
        k = 1
        if with_supplier:
            cell(ws, f"{L(1 + k)}{r}", product_ref("B", i), color="555555"); k += 1
        if with_unit:
            cell(ws, f"{L(1 + k)}{r}", product_ref("C", i), color="555555", align="center")
    return len(cols)


def example(week, i):
    if week < len(WEEK_DATA) and i < len(PRODUCTS):
        return WEEK_DATA[week][i]
    return (None, None, None)


def non_blank_sum(refs):
    joined = ",".join(refs)
    return f'=IF(COUNT({joined})=0,"",SUM({joined}))'


def last_filled(refs):
    """Last non-empty value of a list of refs (last week filled wins)."""
    f = '""'
    for ref in refs:
        f = f'IF({ref}<>"",{ref},{f})'
    return "=" + f


def save(wb, filename):
    OUT.mkdir(parents=True, exist_ok=True)
    wb.calculation.fullCalcOnLoad = True  # no cached values: force recalculation on open
    wb.save(OUT / filename)
    print("wrote", filename)


# ---------------------------------------------------------------- week grid
def week_grid(ws, theme, month_label, fields, first_col, header_row=4, date_row=True,
              examples=True, extra_fields=None):
    """Generic month grid: product columns + 4 week blocks of `fields`.

    fields: list of (label, kind) where kind in
      'stock','cmd','rec' (inputs) or a callable(refs, r) -> formula.
    Returns dict: week -> {label: column letter}.
    """
    cols = {}
    col = first_col
    sub_row = header_row + (2 if date_row else 1)
    first = sub_row + 1
    for w in range(4):
        start = col
        end = col + len(fields) - 1
        header(ws, f"{L(start)}{header_row}", f"Semaine {w + 1}", theme,
               span_to=f"{L(end)}{header_row}" if end > start else None)
        if date_row:
            ws.merge_cells(f"{L(start)}{header_row + 1}:{L(end)}{header_row + 1}")
            c = inp(ws, f"{L(start)}{header_row + 1}", None, fmt="DD/MM/YYYY")
            c.comment = Comment("Date du comptage (facultatif)", "lele_auto")
            for row in ws[f"{L(start)}{header_row + 1}:{L(end)}{header_row + 1}"]:
                for cc in row:
                    cc.fill = INPUT_FILL
                    cc.border = BORDER
        cols[w] = {}
        for j, (label, kind) in enumerate(fields):
            letter = L(col + j)
            cols[w][label] = letter
            header(ws, f"{letter}{sub_row}", label, theme, soft=True)
            ws.column_dimensions[letter].width = 10.5
        col = end + 1
    for w in range(4):
        for i in range(N):
            r = first + i
            ex = example(w, i) if examples else (None, None, None)
            refs = {lab: f"{cols[w][lab]}{r}" for lab in cols[w]}
            for label, kind in fields:
                ref = refs[label]
                if kind in ("stock", "cmd", "rec"):
                    val = {"stock": ex[0], "cmd": ex[1], "rec": ex[2]}[kind]
                    inp(ws, ref, val, example=True)
                else:
                    formula, fmt = kind(refs, r) if callable(kind) else (kind, "General")
                    calc(ws, ref, formula, fmt=fmt)
            if w % 2 == 1:
                pass
    last = first + N - 1
    for w in range(4):
        for label, kind in fields:
            if kind in ("stock", "cmd", "rec"):
                letter = cols[w][label]
                nonneg_validation(ws, f"{letter}{first}:{letter}{last}")
    return cols, first, col


def dispo(refs, r, s="Stock", rec="Reçu"):
    a, b = refs[s], refs[rec]
    return (f'=IF(AND({a}="",{b}=""),"",N({a})+N({b}))', "General")


def ecart(refs, r, cmd="Commandé", rec="Reçu"):
    a, b = refs[cmd], refs[rec]
    return (f'=IF(OR({a}="",{b}=""),"",{b}-{a})', 'General')


def month_header_block(ws, theme, start_col, header_row, sub_row, labels):
    end = start_col + len(labels) - 1
    header(ws, f"{L(start_col)}{header_row}", "Total du mois", theme,
           span_to=f"{L(end)}{sub_row - 1}" if sub_row - 1 > header_row or end > start_col else None)
    for j, lab in enumerate(labels):
        header(ws, f"{L(start_col + j)}{sub_row}", lab, theme, soft=True)
        ws.column_dimensions[L(start_col + j)].width = 12


# ================================================================ proposals
def p01():
    t = Theme("B5523B")
    wb = Workbook()
    guide(wb, t, "1 · Le cahier classique",
          "Exactement ton cahier, en numérique : un onglet par mois, tes produits en lignes, et pour chaque semaine "
          "trois cases — ce qu'il reste, ce que tu commandes, ce que tu reçois. Les totaux du mois se font tout seuls.",
          ["1. Remplis l'onglet Produits une seule fois (produit, fournisseur, unité).",
           "2. Le vendredi : colonne Stock (ce qu'il reste) et colonne Commandé.",
           "3. À la livraison : colonne Reçu. Si c'est moins que commandé, la case Écart devient rouge.",
           "4. Nouveau mois : clic droit sur l'onglet du mois > Dupliquer, puis efface les cases jaunes."],
          ["Produits — ta liste de produits et fournisseurs.",
           "Octobre 2026 — le mois en cours, 4 semaines côte à côte."])
    products_sheet(wb, t)
    ws = wb.create_sheet("Octobre 2026")
    ws.sheet_view.showGridLines = False
    title(ws, "Octobre 2026", "Stock = ce qu'il reste au comptage · Commandé = ce que tu demandes · Reçu = ce qui est arrivé", t)
    for c, h, w in (("A", "Produit", 22), ("B", "Fournisseur", 14), ("C", "Unité", 7)):
        header(ws, f"{c}4", h, t, span_to=f"{c}6")
        ws.column_dimensions[c].width = w
    product_columns(ws, 7, t)
    fields = [("Stock", "stock"), ("Commandé", "cmd"), ("Reçu", "rec"),
              ("Écart", lambda refs, r: ecart(refs, r))]
    cols, first, nxt = week_grid(ws, t, "Octobre", fields, 4)
    labels = ["Commandé", "Reçu", "Écart"]
    month_header_block(ws, t, nxt, 4, 6, labels)
    last = first + N - 1
    for i in range(N):
        r = first + i
        cmd = non_blank_sum([f"{cols[w]['Commandé']}{r}" for w in range(4)])
        rec = non_blank_sum([f"{cols[w]['Reçu']}{r}" for w in range(4)])
        calc(ws, f"{L(nxt)}{r}", cmd, bold=True)
        calc(ws, f"{L(nxt + 1)}{r}", rec, bold=True)
        a, b = f"{L(nxt)}{r}", f"{L(nxt + 1)}{r}"
        calc(ws, f"{L(nxt + 2)}{r}", f'=IF(OR({a}="",{b}=""),"",{b}-{a})', fmt='General', bold=True)
    for w in range(4):
        e = cols[w]["Écart"]
        negative_red(ws, f"{e}{first}:{e}{last}", f"{e}{first}")
    e = L(nxt + 2)
    negative_red(ws, f"{e}{first}:{e}{last}", f"{e}{first}")
    ws.freeze_panes = f"D{first}"
    save(wb, "01_cahier_classique.xlsx")


def weekly_sheet(wb, t, name, week, subtitle):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    title(ws, name, subtitle, t)
    cell(ws, "A3", "Date du comptage :", bold=True, border=False)
    inp(ws, "B3", None, fmt="DD/MM/YYYY")
    heads = [("Produit", 24), ("Fournisseur", 15), ("Unité", 8), ("Stock", 11), ("Commandé", 11),
             ("Reçu", 11), ("Écart", 10), ("Total dispo", 12)]
    for i, (h, w) in enumerate(heads, start=1):
        header(ws, f"{L(i)}5", h, t)
        ws.column_dimensions[L(i)].width = w
    product_columns(ws, 6, t)
    for i in range(N):
        r = 6 + i
        ex = example(week, i)
        inp(ws, f"D{r}", ex[0], example=True)
        inp(ws, f"E{r}", ex[1], example=True)
        inp(ws, f"F{r}", ex[2], example=True)
        calc(ws, f"G{r}", f'=IF(OR(E{r}="",F{r}=""),"",F{r}-E{r})', fmt='General')
        calc(ws, f"H{r}", f'=IF(AND(D{r}="",F{r}=""),"",N(D{r})+N(F{r}))')
    last = 5 + N
    nonneg_validation(ws, f"D6:F{last}")
    negative_red(ws, f"G6:G{last}", "G6")
    ws.freeze_panes = "B6"
    return ws


def p02():
    t = Theme("5E7D5B")
    wb = Workbook()
    guide(wb, t, "2 · Une page par semaine",
          "Chaque semaine a son propre onglet, avec seulement quelques colonnes : c'est lisible même sur téléphone. "
          "Un onglet Récap additionne les 4 semaines tout seul.",
          ["1. Remplis l'onglet Produits une seule fois.",
           "2. Chaque semaine, ouvre l'onglet de la semaine : Stock et Commandé le vendredi, Reçu à la livraison.",
           "3. L'onglet Récap du mois se met à jour tout seul : total commandé, total reçu, dernier stock compté.",
           "4. Nouveau mois : Fichier > Faire une copie, puis efface les cases jaunes des 4 semaines."],
          ["Produits — ta liste.", "Semaine 1 à Semaine 4 — une page par semaine.",
           "Récap du mois — les totaux, calculés."])
    products_sheet(wb, t)
    for w in range(4):
        weekly_sheet(wb, t, f"Semaine {w + 1}", w,
                     "Stock = ce qu'il reste · Commandé = ce que tu demandes · Reçu = ce qui est arrivé")
    ws = wb.create_sheet("Récap du mois")
    ws.sheet_view.showGridLines = False
    title(ws, "Récap du mois", "Calculé à partir des 4 semaines — rien à remplir ici", t)
    heads = [("Produit", 24), ("Unité", 8), ("Commandé S1", 11), ("Commandé S2", 11), ("Commandé S3", 11),
             ("Commandé S4", 11), ("Total commandé", 13), ("Total reçu", 12), ("Manque", 10), ("Dernier stock", 13)]
    for i, (h, w) in enumerate(heads, start=1):
        header(ws, f"{L(i)}4", h, t)
        ws.column_dimensions[L(i)].width = w
    for i in range(N):
        r = 5 + i
        s = 6 + i
        cell(ws, f"A{r}", product_ref("A", i), bold=True)
        cell(ws, f"B{r}", product_ref("C", i), color="555555", align="center")
        for w in range(4):
            calc(ws, f"{L(3 + w)}{r}", f"=IF('Semaine {w + 1}'!E{s}=\"\",\"\",'Semaine {w + 1}'!E{s})")
        calc(ws, f"G{r}", non_blank_sum([f"'Semaine {w + 1}'!E{s}" for w in range(4)]), bold=True)
        calc(ws, f"H{r}", non_blank_sum([f"'Semaine {w + 1}'!F{s}" for w in range(4)]), bold=True)
        calc(ws, f"I{r}", f'=IF(OR(G{r}="",H{r}=""),"",H{r}-G{r})', fmt='General')
        calc(ws, f"J{r}", last_filled([f"'Semaine {w + 1}'!D{s}" for w in range(4)]), bold=True)
    negative_red(ws, f"I5:I{4 + N}", "I5")
    ws.freeze_panes = "B5"
    save(wb, "02_une_page_par_semaine.xlsx")


def p03():
    t = Theme("2F4B7C")
    wb = Workbook()
    rows = 300
    guide(wb, t, "3 · Le journal",
          "Au lieu d'un grand tableau, tu ajoutes simplement une ligne à chaque fois : la date, le produit "
          "(choisi dans une liste), et les quantités. Le fichier range tout et fait les totaux du mois tout seul. "
          "Le plus simple à remplir sur téléphone.",
          ["1. Remplis l'onglet Produits une seule fois.",
           "2. Onglet Saisies : une nouvelle ligne = une date + un produit. Choisis le produit dans la liste déroulante.",
           "3. Tu peux remplir Stock et Commandé le vendredi, puis compléter Reçu sur la même ligne à la livraison.",
           "4. Onglet Synthèse : tape le mois et l'année en haut, tous les totaux s'affichent. Rien à dupliquer, "
           "un seul fichier pour toute l'année."],
          ["Produits — ta liste.", "Saisies — le journal, une ligne par produit et par comptage.",
           "Synthèse — choisis un mois, tout est calculé."])
    products_sheet(wb, t)
    ws = wb.create_sheet("Saisies")
    ws.sheet_view.showGridLines = False
    title(ws, "Saisies", "Une ligne par produit et par comptage. La semaine et le mois se calculent tout seuls.", t)
    heads = [("Date", 12), ("Produit", 24), ("Fournisseur", 15), ("Unité", 8), ("Stock", 10), ("Commandé", 11),
             ("Reçu", 10), ("Écart", 9), ("Semaine", 9), ("Mois", 7), ("Année", 7)]
    for i, (h, w) in enumerate(heads, start=1):
        header(ws, f"{L(i)}4", h, t)
        ws.column_dimensions[L(i)].width = w
    ex_rows = []
    for w, day in ((0, 2), (1, 9)):
        for i, p in enumerate(PRODUCTS):
            ex_rows.append((f"=DATE(2026,10,{day})", p[0]) + WEEK_DATA[w][i])
    first, last = 5, 4 + rows
    rng = lambda c: f"${c}${first}:${c}${last}"
    for k in range(rows):
        r = first + k
        ex = ex_rows[k] if k < len(ex_rows) else (None,) * 5
        inp(ws, f"A{r}", ex[0], example=True, fmt="DD/MM/YYYY")
        inp(ws, f"B{r}", ex[1], example=True).alignment = Alignment(horizontal="left", vertical="center")
        calc(ws, f"C{r}", f'=IF(B{r}="","",IFERROR(INDEX(Produits!$B$5:$B${4 + N},MATCH(B{r},Produits!$A$5:$A${4 + N},0)),""))')
        calc(ws, f"D{r}", f'=IF(B{r}="","",IFERROR(INDEX(Produits!$C$5:$C${4 + N},MATCH(B{r},Produits!$A$5:$A${4 + N},0)),""))')
        inp(ws, f"E{r}", ex[2], example=True)
        inp(ws, f"F{r}", ex[3], example=True)
        inp(ws, f"G{r}", ex[4], example=True)
        calc(ws, f"H{r}", f'=IF(OR(F{r}="",G{r}=""),"",G{r}-F{r})', fmt='General')
        calc(ws, f"I{r}", f'=IF(A{r}="","","S"&MIN(4,INT((DAY(A{r})-1)/7)+1))')
        calc(ws, f"J{r}", f'=IF(A{r}="","",MONTH(A{r}))')
        calc(ws, f"K{r}", f'=IF(A{r}="","",YEAR(A{r}))')
    dv = DataValidation(type="list", formula1=f"=Produits!$A$5:$A${4 + N}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B{first}:B{last}")
    nonneg_validation(ws, f"E{first}:G{last}")
    negative_red(ws, f"H{first}:H{last}", f"H{first}")
    ws.freeze_panes = "B5"
    ws["I3"] = "Semaine du mois : jours 1-7 = S1, 8-14 = S2, 15-21 = S3, 22 et + = S4"
    ws["I3"].font = Font(name=FONT, italic=True, size=8, color="666666")

    sy = wb.create_sheet("Synthèse")
    sy.sheet_view.showGridLines = False
    title(sy, "Synthèse du mois", "Tape le mois et l'année dans les cases jaunes : tout le tableau se recalcule.", t)
    cell(sy, "A3", "Mois (1 à 12)", bold=True)
    inp(sy, "B3", 10)
    cell(sy, "C3", "Année", bold=True)
    inp(sy, "D3", 2026)
    heads = [("Produit", 24), ("Unité", 8), ("Commandé S1", 11), ("Commandé S2", 11), ("Commandé S3", 11),
             ("Commandé S4", 11), ("Total commandé", 13), ("Total reçu", 11), ("Manque", 9),
             ("Dernier stock", 12), ("Compté le", 12)]
    for i, (h, w) in enumerate(heads, start=1):
        header(sy, f"{L(i)}5", h, t)
        sy.column_dimensions[L(i)].width = w
    S = "Saisies!"
    for i in range(N):
        r = 6 + i
        cell(sy, f"A{r}", product_ref("A", i), bold=True)
        cell(sy, f"B{r}", product_ref("C", i), color="555555", align="center")
        crit = f'{S}{rng("B")},$A{r},{S}{rng("J")},$B$3,{S}{rng("K")},$D$3'
        for w in range(4):
            calc(sy, f"{L(3 + w)}{r}",
                 f'=IF($A{r}="","",IF(COUNTIFS({crit},{S}{rng("I")},"S{w + 1}",{S}{rng("F")},">=0")=0,"",'
                 f'SUMIFS({S}{rng("F")},{crit},{S}{rng("I")},"S{w + 1}")))')
        calc(sy, f"G{r}", f'=IF($A{r}="","",IF(COUNTIFS({crit},{S}{rng("F")},">=0")=0,"",SUMIFS({S}{rng("F")},{crit})))', bold=True)
        calc(sy, f"H{r}", f'=IF($A{r}="","",IF(COUNTIFS({crit},{S}{rng("G")},">=0")=0,"",SUMIFS({S}{rng("G")},{crit})))', bold=True)
        calc(sy, f"I{r}", f'=IF(OR(G{r}="",H{r}=""),"",H{r}-G{r})', fmt='General')
        calc(sy, f"K{r}", f'=IF($A{r}="","",IF(COUNTIFS({crit},{S}{rng("E")},">=0")=0,"",'
                          f'_xlfn.MAXIFS({S}{rng("A")},{crit},{S}{rng("E")},">=0")))', fmt="DD/MM/YYYY")
        calc(sy, f"J{r}", f'=IF(K{r}="","",SUMIFS({S}{rng("E")},{S}{rng("B")},$A{r},{S}{rng("A")},K{r}))', bold=True)
    negative_red(sy, f"I6:I{5 + N}", "I6")
    sy.freeze_panes = "B6"
    wb.move_sheet("Synthèse", offset=-1)
    save(wb, "03_journal.xlsx")


def p04():
    t = Theme("7A4069")
    wb = Workbook()
    guide(wb, t, "4 · Par fournisseur",
          "Un onglet par fournisseur, avec ses produits. Quand tu fais ta commande, tu as sous les yeux tout ce que "
          "tu dois demander à ce fournisseur-là, et ses coordonnées. Pratique si tu passes tes commandes par téléphone.",
          ["1. Onglet Fournisseurs : nom, contact, jour de commande, jour de livraison.",
           "2. Dans chaque onglet fournisseur, écris ses produits dans la colonne Produit (cases jaunes).",
           "3. Chaque semaine : Stock, Commandé, puis Reçu à la livraison. Le total du mois se calcule.",
           "4. Nouveau fournisseur : clic droit sur un onglet fournisseur > Dupliquer, puis renomme-le."],
          ["Fournisseurs — carnet d'adresses et jours de commande/livraison.",
           "Metro, Crèmerie, Transgourmet — un onglet par fournisseur avec ses produits.",
           "Récap fournisseurs — combien de produits, combien de livraisons incomplètes, par fournisseur."])
    ws = wb.create_sheet("Fournisseurs")
    ws.sheet_view.showGridLines = False
    title(ws, "Mes fournisseurs", "Carnet d'adresses", t)
    heads = [("Fournisseur", 16), ("Contact / commercial", 24), ("Téléphone", 16), ("E-mail", 24),
             ("Jour de commande", 16), ("Jour de livraison", 16), ("Notes", 30)]
    for i, (h, w) in enumerate(heads, start=1):
        header(ws, f"{L(i)}4", h, t)
        ws.column_dimensions[L(i)].width = w
    for k in range(10):
        r = 5 + k
        s = SUPPLIERS[k] if k < len(SUPPLIERS) else (None,) * 5
        for j, v in enumerate((s[0], s[1], s[2], None, s[3], s[4], None)):
            inp(ws, f"{L(j + 1)}{r}", v, example=True).alignment = Alignment(horizontal="left", vertical="center")
    dv = DataValidation(type="list", formula1='"lundi,mardi,mercredi,jeudi,vendredi,samedi,dimanche"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E5:F14")
    rows_per = 20
    for sname, _, tel, jc, jl in SUPPLIERS:
        sh = wb.create_sheet(sname)
        sh.sheet_view.showGridLines = False
        title(sh, sname, f"Commande le {jc} · Livraison le {jl} · Tél. {tel}", t)
        header(sh, "A4", "Produit", t, span_to="A6")
        header(sh, "B4", "Unité", t, span_to="B6")
        sh.column_dimensions["A"].width = 24
        sh.column_dimensions["B"].width = 8
        fields = [("Stock", "stock"), ("Commandé", "cmd"), ("Reçu", "rec")]
        prods = [(i, p) for i, p in enumerate(PRODUCTS) if p[1] == sname]
        cols = {}
        col = 3
        for w in range(4):
            header(sh, f"{L(col)}4", f"Semaine {w + 1}", t, span_to=f"{L(col + 2)}4")
            sh.merge_cells(f"{L(col)}5:{L(col + 2)}5")
            inp(sh, f"{L(col)}5", None, fmt="DD/MM/YYYY")
            for row in sh[f"{L(col)}5:{L(col + 2)}5"]:
                for cc in row:
                    cc.fill = INPUT_FILL
                    cc.border = BORDER
            cols[w] = {}
            for j, (lab, _) in enumerate(fields):
                cols[w][lab] = L(col + j)
                header(sh, f"{L(col + j)}6", lab, t, soft=True)
                sh.column_dimensions[L(col + j)].width = 10.5
            col += 3
        header(sh, f"{L(col)}4", "Total du mois", t, span_to=f"{L(col + 2)}5")
        for j, lab in enumerate(("Commandé", "Reçu", "Manque")):
            header(sh, f"{L(col + j)}6", lab, t, soft=True)
            sh.column_dimensions[L(col + j)].width = 11
        for k in range(rows_per):
            r = 7 + k
            gi, p = prods[k] if k < len(prods) else (None, None)
            inp(sh, f"A{r}", p[0] if p else None, example=True).alignment = Alignment(horizontal="left", vertical="center")
            inp(sh, f"B{r}", p[2] if p else None, example=True)
            for w in range(4):
                ex = example(w, gi) if gi is not None else (None, None, None)
                inp(sh, f"{cols[w]['Stock']}{r}", ex[0], example=True)
                inp(sh, f"{cols[w]['Commandé']}{r}", ex[1], example=True)
                inp(sh, f"{cols[w]['Reçu']}{r}", ex[2], example=True)
            calc(sh, f"{L(col)}{r}", non_blank_sum([f"{cols[w]['Commandé']}{r}" for w in range(4)]), bold=True)
            calc(sh, f"{L(col + 1)}{r}", non_blank_sum([f"{cols[w]['Reçu']}{r}" for w in range(4)]), bold=True)
            a, b = f"{L(col)}{r}", f"{L(col + 1)}{r}"
            calc(sh, f"{L(col + 2)}{r}", f'=IF(OR({a}="",{b}=""),"",{b}-{a})', fmt='General', bold=True)
        last = 6 + rows_per
        nonneg_validation(sh, f"C7:{L(col - 1)}{last}")
        m = L(col + 2)
        negative_red(sh, f"{m}7:{m}{last}", f"{m}7")
        sh.freeze_panes = "B7"
    rc = wb.create_sheet("Récap fournisseurs")
    rc.sheet_view.showGridLines = False
    title(rc, "Récap fournisseurs", "Calculé — rien à remplir", t)
    for i, (h, w) in enumerate((("Fournisseur", 18), ("Produits suivis", 15), ("Produits livrés incomplets ce mois", 30)), start=1):
        header(rc, f"{L(i)}4", h, t)
        rc.column_dimensions[L(i)].width = w
    m_col = L(3 + 4 * 3 + 2)
    for k, s in enumerate(SUPPLIERS):
        r = 5 + k
        cell(rc, f"A{r}", s[0], bold=True)
        calc(rc, f"B{r}", f"=COUNTA('{s[0]}'!A7:A{6 + rows_per})")
        calc(rc, f"C{r}", f"=COUNTIF('{s[0]}'!{m_col}7:{m_col}{6 + rows_per},\"<0\")")
    rc.conditional_formatting.add("C5:C7", FormulaRule(formula=["C5>0"], fill=RED_FILL))
    save(wb, "04_par_fournisseur.xlsx")


def p05():
    t = Theme("B8860B")
    wb = Workbook()
    guide(wb, t, "5 · La liste de courses",
          "Pour chaque produit, tu indiques une fois pour toutes ton stock minimum et ton stock idéal. "
          "Le vendredi, tu comptes ce qu'il reste : le fichier te propose la quantité à commander et signale "
          "en orange ce qui est passé sous le minimum. Tu restes libre de commander autre chose.",
          ["1. Onglet Produits : produit, fournisseur, unité, stock minimum, stock idéal.",
           "2. Le vendredi : colonne Stock. La colonne Suggéré se remplit (stock idéal − stock restant).",
           "3. Tu écris ce que tu commandes vraiment dans Commandé (tu peux suivre la suggestion ou pas).",
           "4. À la livraison : Reçu. Nouveau mois : duplique l'onglet du mois."],
          ["Produits — ta liste, avec stock minimum et stock idéal.",
           "Octobre 2026 — 4 semaines ; orange = sous le minimum, rouge = livraison incomplète."])
    products_sheet(wb, t, extra=("mini", "cible"))
    ws = wb.create_sheet("Octobre 2026")
    ws.sheet_view.showGridLines = False
    title(ws, "Octobre 2026", "Suggéré = stock idéal − stock restant (arrondi). Orange = sous le minimum.", t)
    for c, h, w in (("A", "Produit", 22), ("B", "Fournisseur", 14), ("C", "Unité", 7)):
        header(ws, f"{c}4", h, t, span_to=f"{c}6")
        ws.column_dimensions[c].width = w
    product_columns(ws, 7, t)
    # min stock copied here so conditional formatting stays on-sheet (Google Sheets requirement)
    header(ws, "D4", "Stock mini", t, span_to="D6")
    ws.column_dimensions["D"].width = 11
    for i in range(N):
        calc(ws, f"D{7 + i}", product_ref("D", i)).font = Font(name=FONT, color="555555")

    def suggested(refs, r):
        i = r - 7
        s = refs["Stock"]
        return (f'=IF(OR({s}="",Produits!$E${5 + i}=""),"",MAX(0,CEILING(Produits!$E${5 + i}-{s},1)))', "General")

    fields = [("Stock", "stock"), ("Suggéré", suggested), ("Commandé", "cmd"), ("Reçu", "rec")]
    cols, first, nxt = week_grid(ws, t, "Octobre", fields, 5)
    last = first + N - 1
    for w in range(4):
        s = cols[w]["Stock"]
        ws.conditional_formatting.add(
            f"{s}{first}:{s}{last}",
            FormulaRule(formula=[f'AND(ISNUMBER({s}{first}),ISNUMBER($D{first}),{s}{first}<=$D{first})'], fill=ORANGE_FILL))
        cm, rc = cols[w]["Commandé"], cols[w]["Reçu"]
        ws.conditional_formatting.add(
            f"{rc}{first}:{rc}{last}",
            FormulaRule(formula=[f"AND(ISNUMBER({rc}{first}),ISNUMBER({cm}{first}),{rc}{first}<{cm}{first})"], fill=RED_FILL))
        sg = cols[w]["Suggéré"]
        for i in range(N):
            ws[f"{sg}{first + i}"].font = Font(name=FONT, italic=True, color="8A6D00")
    month_header_block(ws, t, nxt, 4, 6, ["Commandé", "Reçu"])
    for i in range(N):
        r = first + i
        calc(ws, f"{L(nxt)}{r}", non_blank_sum([f"{cols[w]['Commandé']}{r}" for w in range(4)]), bold=True)
        calc(ws, f"{L(nxt + 1)}{r}", non_blank_sum([f"{cols[w]['Reçu']}{r}" for w in range(4)]), bold=True)
    ws.freeze_panes = f"E{first}"
    save(wb, "05_liste_de_courses.xlsx")


def p06():
    t = Theme("1F7A7A")
    wb = Workbook()
    guide(wb, t, "6 · Contrôle des livraisons",
          "Met l'accent sur ce que tu reçois vraiment. Pour chaque produit, le fichier compare commandé et reçu "
          "et affiche OK, MANQUE ou EN TROP. Un onglet Réclamations garde la trace des appels au fournisseur.",
          ["1. Remplis l'onglet Produits une seule fois.",
           "2. Chaque semaine : Stock, Commandé, puis Reçu à la livraison. Le Statut s'affiche tout seul.",
           "3. En cas de MANQUE : ajoute une ligne dans Réclamations (date, produit, ce que tu as fait).",
           "4. Nouveau mois : duplique l'onglet du mois."],
          ["Produits — ta liste.", "Octobre 2026 — stock, commandé, reçu, statut par semaine.",
           "Réclamations — suivi des livraisons incomplètes."])
    products_sheet(wb, t)
    ws = wb.create_sheet("Octobre 2026")
    ws.sheet_view.showGridLines = False
    title(ws, "Octobre 2026", "Statut : OK = tout reçu · MANQUE = moins que commandé · EN TROP = plus que commandé", t)
    for c, h, w in (("A", "Produit", 22), ("B", "Fournisseur", 14), ("C", "Unité", 7)):
        header(ws, f"{c}4", h, t, span_to=f"{c}6")
        ws.column_dimensions[c].width = w
    product_columns(ws, 7, t)

    def status(refs, r):
        a, b = refs["Commandé"], refs["Reçu"]
        return (f'=IF(OR({a}="",{b}="",N({a})=0),"",IF({b}<{a},"MANQUE",IF({b}>{a},"EN TROP","OK")))', "General")

    fields = [("Stock", "stock"), ("Commandé", "cmd"), ("Reçu", "rec"),
              ("Écart", lambda refs, r: ecart(refs, r)), ("Statut", status)]
    cols, first, nxt = week_grid(ws, t, "Octobre", fields, 4)
    last = first + N - 1
    for w in range(4):
        st = cols[w]["Statut"]
        ws.column_dimensions[st].width = 10
        rng = f"{st}{first}:{st}{last}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{st}{first}="MANQUE"'], fill=RED_FILL))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{st}{first}="EN TROP"'], fill=ORANGE_FILL))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{st}{first}="OK"'], fill=GREEN_FILL))
    month_header_block(ws, t, nxt, 4, 6, ["Livraisons incomplètes"])
    ws.column_dimensions[L(nxt)].width = 14
    for i in range(N):
        r = first + i
        refs = ",".join(f"{cols[w]['Statut']}{r}" for w in range(4))
        calc(ws, f"{L(nxt)}{r}", f'=IF(A{r}="","",COUNTIF({cols[0]["Statut"]}{r}:{cols[3]["Statut"]}{r},"MANQUE"))', bold=True)
    ws.conditional_formatting.add(f"{L(nxt)}{first}:{L(nxt)}{last}",
                                  FormulaRule(formula=[f"AND(ISNUMBER({L(nxt)}{first}),{L(nxt)}{first}>0)"], fill=RED_FILL))
    ws.freeze_panes = f"D{first}"
    rc = wb.create_sheet("Réclamations")
    rc.sheet_view.showGridLines = False
    title(rc, "Réclamations fournisseurs", "Une ligne par livraison incomplète", t)
    heads = [("Date", 12), ("Produit", 24), ("Fournisseur", 15), ("Commandé", 11), ("Reçu", 10), ("Manque", 10),
             ("Ce que j'ai fait", 30), ("Réglé ?", 10)]
    for i, (h, w) in enumerate(heads, start=1):
        header(rc, f"{L(i)}4", h, t)
        rc.column_dimensions[L(i)].width = w
    for k in range(60):
        r = 5 + k
        ex = k == 0
        inp(rc, f"A{r}", "=DATE(2026,10,5)" if ex else None, example=True, fmt="DD/MM/YYYY")
        inp(rc, f"B{r}", "Farine T55" if ex else None, example=True).alignment = Alignment(horizontal="left", vertical="center")
        calc(rc, f"C{r}", f'=IF(B{r}="","",IFERROR(INDEX(Produits!$B$5:$B${4 + N},MATCH(B{r},Produits!$A$5:$A${4 + N},0)),""))')
        inp(rc, f"D{r}", 10 if ex else None, example=True)
        inp(rc, f"E{r}", 7 if ex else None, example=True)
        calc(rc, f"F{r}", f'=IF(OR(D{r}="",E{r}=""),"",D{r}-E{r})')
        inp(rc, f"G{r}", "Appelé Metro, avoir promis" if ex else None, example=True).alignment = Alignment(horizontal="left", vertical="center")
        inp(rc, f"H{r}", "Non" if ex else None, example=True)
    dv = DataValidation(type="list", formula1=f"=Produits!$A$5:$A${4 + N}", allow_blank=True)
    rc.add_data_validation(dv)
    dv.add("B5:B64")
    dv2 = DataValidation(type="list", formula1='"Oui,Non"', allow_blank=True)
    rc.add_data_validation(dv2)
    dv2.add("H5:H64")
    rc.conditional_formatting.add("H5:H64", FormulaRule(formula=['H5="Non"'], fill=RED_FILL))
    rc.conditional_formatting.add("H5:H64", FormulaRule(formula=['H5="Oui"'], fill=GREEN_FILL))
    rc.freeze_panes = "A5"
    save(wb, "06_controle_livraisons.xlsx")


def simple_month(wb, t, name, examples, title_text=None):
    """Month sheet: product cols + 4 weeks x (Stock, Commandé, Reçu) + month totals + last stock.

    Returns (sheet, first_row, dict of total column letters).
    """
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    title(ws, title_text or name, "Stock = ce qu'il reste · Commandé = ce que tu demandes · Reçu = ce qui est arrivé", t)
    for c, h, w in (("A", "Produit", 22), ("B", "Fournisseur", 14), ("C", "Unité", 7)):
        header(ws, f"{c}4", h, t, span_to=f"{c}6")
        ws.column_dimensions[c].width = w
    product_columns(ws, 7, t)
    fields = [("Stock", "stock"), ("Commandé", "cmd"), ("Reçu", "rec")]
    cols, first, nxt = week_grid(ws, t, name, fields, 4, examples=examples)
    labels = ["Commandé", "Reçu", "Écart", "Dernier stock"]
    month_header_block(ws, t, nxt, 4, 6, labels)
    tot = {lab: L(nxt + j) for j, lab in enumerate(labels)}
    for i in range(N):
        r = first + i
        calc(ws, f"{tot['Commandé']}{r}", non_blank_sum([f"{cols[w]['Commandé']}{r}" for w in range(4)]), bold=True)
        calc(ws, f"{tot['Reçu']}{r}", non_blank_sum([f"{cols[w]['Reçu']}{r}" for w in range(4)]), bold=True)
        a, b = f"{tot['Commandé']}{r}", f"{tot['Reçu']}{r}"
        calc(ws, f"{tot['Écart']}{r}", f'=IF(OR({a}="",{b}=""),"",{b}-{a})', fmt='General', bold=True)
        calc(ws, f"{tot['Dernier stock']}{r}", last_filled([f"{cols[w]['Stock']}{r}" for w in range(4)]), bold=True)
    last = first + N - 1
    e = tot["Écart"]
    negative_red(ws, f"{e}{first}:{e}{last}", f"{e}{first}")
    ws.freeze_panes = f"D{first}"
    return ws, first, tot


def p07():
    t = Theme("3D3D3D")
    wb = Workbook()
    guide(wb, t, "7 · Tableau de bord",
          "Même saisie que le cahier classique, plus une page de résumé : combien tu as dépensé ce mois-ci, "
          "combien de livraisons étaient incomplètes, quels produits sont sous le minimum, et un graphique "
          "des 10 produits qui te coûtent le plus.",
          ["1. Onglet Produits : produit, fournisseur, unité, stock minimum, prix unitaire.",
           "2. Onglet du mois : Stock, Commandé, Reçu chaque semaine.",
           "3. Onglet Tableau de bord : tout se calcule, rien à remplir.",
           "4. Nouveau mois : duplique l'onglet du mois et change la référence en haut du tableau de bord "
           "(ou garde un fichier par mois)."],
          ["Tableau de bord — les chiffres clés et le graphique.",
           "Produits — ta liste, avec minimum et prix.", "Octobre 2026 — la saisie de la semaine."])
    products_sheet(wb, t, extra=("mini", "prix"))
    month, first, tot = simple_month(wb, t, "Octobre 2026", True)
    db = wb.create_sheet("Tableau de bord", 1)
    db.sheet_view.showGridLines = False
    title(db, "Tableau de bord — Octobre 2026", "Calculé à partir de l'onglet Octobre 2026 — rien à remplir", t)
    M = "'Octobre 2026'!"
    # helper table (columns K..Q), kept visible but grey, to feed KPIs and the top 10
    hh = [("Produit", 22), ("Reçu", 9), ("Prix", 9), ("Dépense (€)", 12), ("Dernier stock", 12), ("Sous le mini", 11), ("Tri", 6)]
    for j, (h, w) in enumerate(hh):
        header(db, f"{L(11 + j)}4", h, t, soft=True)
        db.column_dimensions[L(11 + j)].width = w
    cell(db, "K3", "Calculs intermédiaires", bold=True, color="888888", border=False)
    for i in range(N):
        r = 5 + i
        mr = first + i
        calc(db, f"K{r}", f"={M}A{mr}")
        calc(db, f"L{r}", f"=N({M}{tot['Reçu']}{mr})")
        calc(db, f"M{r}", f"=N(Produits!E{5 + i})", fmt='#,##0.00')
        calc(db, f"N{r}", f"=L{r}*M{r}", fmt='#,##0.00')
        calc(db, f"O{r}", f"={M}{tot['Dernier stock']}{mr}")
        calc(db, f"P{r}", f'=IF(AND(K{r}<>"",O{r}<>"",Produits!D{5 + i}<>""),IF(N(O{r})<=N(Produits!D{5 + i}),1,0),0)')
        calc(db, f"Q{r}", f'=IF(K{r}="",-1,N{r}+ROW()/1000000)')
        for c in "KLMNOPQ":
            db[f"{c}{r}"].font = Font(name=FONT, color="888888", size=9)
    for c in "KLMNOPQ":  # helper columns: needed by formulas, hidden from Lele
        db.column_dimensions[c].hidden = True
    last = 4 + N
    kpis = [("Dépensé ce mois-ci", f"=SUM(N5:N{last})", '#,##0.00 "€"'),
            ("Livraisons incomplètes", f"=COUNTIF({M}{tot['Écart']}{first}:{tot['Écart']}{first + N - 1},\"<0\")", "0"),
            ("Produits sous le minimum", f"=SUM(P5:P{last})", "0"),
            ("Produits suivis", f"=COUNTA(Produits!A5:A{4 + N})", "0")]
    for k, (lab, f, fmt) in enumerate(kpis):
        col = L(1 + 2 * k)
        db.merge_cells(f"{col}4:{L(2 + 2 * k)}4")
        db.merge_cells(f"{col}5:{L(2 + 2 * k)}6")
        cell(db, f"{col}4", lab, bold=True, color="FFFFFF", fill=t.fill, align="center", size=9)
        cell(db, f"{col}5", f, bold=True, size=18, align="center", fmt=fmt, fill=t.softer)
        for ref in (f"{L(2 + 2 * k)}4",):
            db[ref].fill = t.fill
    db.conditional_formatting.add("C5", FormulaRule(formula=["C5>0"], fill=RED_FILL))
    db.conditional_formatting.add("E5", FormulaRule(formula=["E5>0"], fill=ORANGE_FILL))
    for c in "ABCDEFGH":
        db.column_dimensions[c].width = 12
    cell(db, "A9", "Top 10 des dépenses du mois", bold=True, size=12, color=t.main, border=False)
    for j, h in enumerate(("Rang", "Produit", "", "Reçu", "Dépense (€)")):
        if h:
            header(db, f"{L(1 + j)}10", h, t)
    db.merge_cells("B10:C10")
    for k in range(10):
        r = 11 + k
        big = f"LARGE($Q$5:$Q${last},{k + 1})"
        idx = f"MATCH({big},$Q$5:$Q${last},0)"
        cell(db, f"A{r}", k + 1, align="center")
        db.merge_cells(f"B{r}:C{r}")
        calc(db, f"B{r}", f'=IF({big}<=0,"",INDEX($K$5:$K${last},{idx}))')
        db[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center")
        calc(db, f"D{r}", f'=IF(B{r}="","",INDEX($L$5:$L${last},{idx}))')
        calc(db, f"E{r}", f'=IF(B{r}="","",INDEX($N$5:$N${last},{idx}))', fmt='#,##0.00 "€"')
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Dépenses du mois (€)"
    ch.style = 2
    ch.legend = None
    ch.y_axis.title = None
    data = Reference(db, min_col=5, min_row=10, max_row=20)
    cats = Reference(db, min_col=2, min_row=11, max_row=20)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.x_axis.scaling.orientation = "maxMin"
    ch.height, ch.width = 7.5, 14
    ch.series[0].graphicalProperties.solidFill = t.main
    db.add_chart(ch, "A23")
    save(wb, "07_tableau_de_bord.xlsx")


def p08():
    t = Theme("2E6B3F")
    wb = Workbook()
    guide(wb, t, "8 · Suivi sur l'année",
          "Un fichier pour toute l'année : un onglet par mois (déjà prêts), et un onglet Année qui montre, "
          "produit par produit, combien tu as reçu chaque mois, ta moyenne mensuelle et ton plus gros mois. "
          "C'est ce qui permet de voir combien tu rachètes en général de chaque produit.",
          ["1. Remplis l'onglet Produits une seule fois.",
           "2. Chaque semaine : dans l'onglet du mois en cours, Stock, Commandé, Reçu.",
           "3. L'onglet Année se remplit tout seul au fil des mois.",
           "4. Rien à dupliquer : les 12 mois sont déjà là."],
          ["Année 2026 — reçu par mois, moyenne, maximum.", "Produits — ta liste.",
           "Janvier à Décembre — un onglet par mois."])
    products_sheet(wb, t)
    firsts = {}
    tots = {}
    for m, mname in enumerate(MONTHS):
        ws, first, tot = simple_month(wb, t, mname, examples=(mname == "Octobre"), title_text=f"{mname} 2026")
        firsts[mname], tots[mname] = first, tot
    if True:  # a few September values so the yearly view is not empty in the example
        sep = wb["Septembre"]
        for i in range(len(PRODUCTS)):
            for w in range(2):
                s, c, r = WEEK_DATA[1 - w][i]
                base = 4 + w * 3
                sep[f"{L(base + 1)}{firsts['Septembre'] + i}"].value = c
                sep[f"{L(base + 2)}{firsts['Septembre'] + i}"].value = r
                sep[f"{L(base)}{firsts['Septembre'] + i}"].value = s
                for k in range(3):
                    sep[f"{L(base + k)}{firsts['Septembre'] + i}"].font = EXAMPLE_FONT
    yr = wb.create_sheet("Année 2026", 1)
    yr.sheet_view.showGridLines = False
    title(yr, "Année 2026 — quantités reçues par mois", "Calculé à partir des 12 onglets mensuels — rien à remplir", t)
    short = ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."]
    heads = [("Produit", 22), ("Unité", 7)] + [(m, 8) for m in short] + \
            [("Total", 10), ("Moyenne / mois", 16), ("Plus gros mois", 16)]
    for i, (h, w) in enumerate(heads, start=1):
        header(yr, f"{L(i)}4", h, t)
        yr.column_dimensions[L(i)].width = w
    for i in range(N):
        r = 5 + i
        cell(yr, f"A{r}", product_ref("A", i), bold=True)
        cell(yr, f"B{r}", product_ref("C", i), color="555555", align="center")
        for m, mname in enumerate(MONTHS):
            src = f"'{mname}'!{tots[mname]['Reçu']}{firsts[mname] + i}"
            calc(yr, f"{L(3 + m)}{r}", f'=IF(N({src})=0,"",N({src}))')
        rng = f"C{r}:N{r}"
        calc(yr, f"O{r}", f'=IF(COUNT({rng})=0,"",SUM({rng}))', bold=True)
        calc(yr, f"P{r}", f'=IF(COUNT({rng})=0,"",ROUND(AVERAGE({rng}),1))', bold=True)
        calc(yr, f"Q{r}", f'=IF(COUNT({rng})=0,"",MAX({rng}))')
    yr.freeze_panes = "C5"
    save(wb, "08_suivi_annuel.xlsx")


def p09():
    t = Theme("6B4226")
    wb = Workbook()
    guide(wb, t, "9 · Stock en euros",
          "En plus des quantités, le fichier calcule la valeur de ton stock en euros et ce que tu as dépensé chez "
          "chaque fournisseur. Utile le jour de l'inventaire : tu comptes, la valeur totale se fait toute seule.",
          ["1. Onglet Produits : produit, fournisseur, unité, prix unitaire.",
           "2. Onglet du mois : Stock, Commandé, Reçu chaque semaine. La valeur du stock s'affiche.",
           "3. Jour d'inventaire : onglet Inventaire, tape les quantités comptées, le total en euros se calcule.",
           "4. Onglet Dépenses : combien tu as dépensé chez chaque fournisseur ce mois-ci."],
          ["Produits — ta liste, avec les prix.", "Octobre 2026 — la saisie de la semaine + valeur.",
           "Inventaire — comptage complet et valeur totale.", "Dépenses — total par fournisseur."])
    products_sheet(wb, t, extra=("prix",))
    ws = wb.create_sheet("Octobre 2026")
    ws.sheet_view.showGridLines = False
    title(ws, "Octobre 2026", "Valeur = stock × prix unitaire (onglet Produits)", t)
    for c, h, w in (("A", "Produit", 22), ("B", "Fournisseur", 14), ("C", "Unité", 7)):
        header(ws, f"{c}4", h, t, span_to=f"{c}6")
        ws.column_dimensions[c].width = w
    product_columns(ws, 7, t)

    def value(refs, r):
        s = refs["Stock"]
        return (f'=IF({s}="","",{s}*N(Produits!$D${r - 2}))', '#,##0.00 "€"')

    fields = [("Stock", "stock"), ("Valeur", value), ("Commandé", "cmd"), ("Reçu", "rec")]
    cols, first, nxt = week_grid(ws, t, "Octobre", fields, 4)
    month_header_block(ws, t, nxt, 4, 6, ["Reçu", "Dépense (€)"])
    for i in range(N):
        r = first + i
        calc(ws, f"{L(nxt)}{r}", non_blank_sum([f"{cols[w]['Reçu']}{r}" for w in range(4)]), bold=True)
        calc(ws, f"{L(nxt + 1)}{r}", f'=IF({L(nxt)}{r}="","",{L(nxt)}{r}*N(Produits!$D${5 + i}))', fmt='#,##0.00 "€"', bold=True)
    last = first + N - 1
    header(ws, f"A{last + 1}", "Valeur totale du stock", t, span_to=f"C{last + 1}")
    for w in range(4):
        v = cols[w]["Valeur"]
        calc(ws, f"{v}{last + 1}", f"=SUM({v}{first}:{v}{last})", fmt='#,##0.00 "€"', bold=True, fill=t.soft)
    d = L(nxt + 1)
    calc(ws, f"{d}{last + 1}", f"=SUM({d}{first}:{d}{last})", fmt='#,##0.00 "€"', bold=True, fill=t.soft)
    ws.freeze_panes = f"D{first}"
    inv = wb.create_sheet("Inventaire")
    inv.sheet_view.showGridLines = False
    title(inv, "Inventaire", "Le jour de l'inventaire : compte tout, tape les quantités, la valeur se calcule.", t)
    cell(inv, "A3", "Date de l'inventaire :", bold=True, border=False)
    inp(inv, "B3", "=DATE(2026,12,31)", example=True, fmt="DD/MM/YYYY")
    heads = [("Produit", 24), ("Unité", 8), ("Quantité comptée", 16), ("Prix unitaire", 13), ("Valeur", 13)]
    for i, (h, w) in enumerate(heads, start=1):
        header(inv, f"{L(i)}5", h, t)
        inv.column_dimensions[L(i)].width = w
    for i in range(N):
        r = 6 + i
        cell(inv, f"A{r}", product_ref("A", i), bold=True)
        cell(inv, f"B{r}", product_ref("C", i), color="555555", align="center")
        inp(inv, f"C{r}", None)
        calc(inv, f"D{r}", f'=IF(A{r}="","",N(Produits!D{5 + i}))', fmt='#,##0.00 "€"')
        calc(inv, f"E{r}", f'=IF(C{r}="","",C{r}*N(D{r}))', fmt='#,##0.00 "€"')
    header(inv, f"A{6 + N}", "Valeur totale", t, span_to=f"D{6 + N}")
    calc(inv, f"E{6 + N}", f"=SUM(E6:E{5 + N})", fmt='#,##0.00 "€"', bold=True, fill=t.soft)
    nonneg_validation(inv, f"C6:C{5 + N}")
    inv.freeze_panes = "A6"
    dp = wb.create_sheet("Dépenses")
    dp.sheet_view.showGridLines = False
    title(dp, "Dépenses par fournisseur — Octobre 2026", "Quantités reçues × prix unitaire", t)
    header(dp, "A4", "Fournisseur", t)
    header(dp, "B4", "Dépensé ce mois-ci", t)
    widths(dp, {"A": 20, "B": 20})
    M = "'Octobre 2026'!"
    for k, s in enumerate(SUPPLIERS + [(None,)] * 3):
        r = 5 + k
        inp(dp, f"A{r}", s[0], example=True).alignment = Alignment(horizontal="left", vertical="center")
        calc(dp, f"B{r}", f'=IF(A{r}="","",SUMIF({M}$B${first}:$B${last},A{r},{M}${d}${first}:${d}${last}))', fmt='#,##0.00 "€"')
    header(dp, f"A{5 + len(SUPPLIERS) + 3}", "Total", t)
    calc(dp, f"B{5 + len(SUPPLIERS) + 3}", f"=SUM(B5:B{4 + len(SUPPLIERS) + 3})", fmt='#,##0.00 "€"', bold=True, fill=t.soft)
    save(wb, "09_stock_en_euros.xlsx")


def p10():
    t = Theme("A23B5A")
    wb = Workbook()
    guide(wb, t, "10 · Le carnet vertical (téléphone)",
          "Pensé pour le téléphone : pas de défilement de côté. Les 4 semaines sont les unes sous les autres, "
          "avec seulement 5 colonnes et une grande écriture. Le bilan du mois est tout en bas.",
          ["1. Remplis l'onglet Produits une seule fois.",
           "2. Fais défiler jusqu'à la semaine en cours : Stock, Commandé, Reçu.",
           "3. Le bilan du mois en bas de page se calcule tout seul.",
           "4. Nouveau mois : duplique l'onglet du mois."],
          ["Produits — ta liste.", "Octobre 2026 — les 4 semaines empilées + le bilan."])
    products_sheet(wb, t)
    ws = wb.create_sheet("Octobre 2026")
    ws.sheet_view.showGridLines = False
    title(ws, "Octobre 2026", "Fais défiler vers le bas pour les semaines suivantes", t)
    widths(ws, {"A": 30, "B": 10, "C": 11, "D": 10, "E": 9})
    big = Font(name=FONT, size=12)
    r = 4
    week_rows = []
    for w in range(4):
        ws.merge_cells(f"A{r}:C{r}")
        cell(ws, f"A{r}", f"SEMAINE {w + 1}", bold=True, color="FFFFFF", fill=t.fill, size=13)
        for c in "BC":
            ws[f"{c}{r}"].fill = t.fill
        ws.merge_cells(f"D{r}:E{r}")
        c = inp(ws, f"D{r}", None, fmt="DD/MM/YYYY")
        c.comment = Comment("Date du comptage (facultatif)", "lele_auto")
        ws[f"E{r}"].fill = INPUT_FILL
        ws.row_dimensions[r].height = 24
        r += 1
        for j, h in enumerate(("Produit", "Stock", "Commandé", "Reçu", "Écart")):
            header(ws, f"{L(1 + j)}{r}", h, t, soft=True)
        r += 1
        start = r
        for i in range(N):
            ex = example(w, i)
            c = cell(ws, f"A{r}", f'=IF(Produits!$A${5 + i}="","",Produits!$A${5 + i}&" ("&Produits!$C${5 + i}&")")', bold=True)
            c.font = Font(name=FONT, size=12, bold=True)
            for col, v in zip("BCD", ex):
                cc = inp(ws, f"{col}{r}", v, example=True)
                cc.font = Font(name=FONT, size=12, color=EXAMPLE_FONT.color if v is not None else "000000")
            cc = calc(ws, f"E{r}", f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})', fmt='General')
            cc.font = big
            ws.row_dimensions[r].height = 20
            r += 1
        week_rows.append(start)
        nonneg_validation(ws, f"B{start}:D{r - 1}")
        negative_red(ws, f"E{start}:E{r - 1}", f"E{start}")
        r += 1
    ws.merge_cells(f"A{r}:E{r}")
    cell(ws, f"A{r}", "BILAN DU MOIS", bold=True, color="FFFFFF", fill=t.fill, size=13)
    ws.row_dimensions[r].height = 24
    r += 1
    for j, h in enumerate(("Produit", "Dernier stock", "Commandé", "Reçu", "Manque")):
        header(ws, f"{L(1 + j)}{r}", h, t, soft=True)
    r += 1
    start = r
    for i in range(N):
        rows = [s + i for s in week_rows]
        cell(ws, f"A{r}", f"=A{rows[0]}", bold=True).font = Font(name=FONT, size=12, bold=True)
        calc(ws, f"B{r}", last_filled([f"B{x}" for x in rows])).font = big
        calc(ws, f"C{r}", non_blank_sum([f"C{x}" for x in rows])).font = big
        calc(ws, f"D{r}", non_blank_sum([f"D{x}" for x in rows])).font = big
        calc(ws, f"E{r}", f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})', fmt='General').font = big
        r += 1
    negative_red(ws, f"E{start}:E{r - 1}", f"E{start}")
    ws.freeze_panes = "B4"
    save(wb, "10_carnet_vertical_telephone.xlsx")


if __name__ == "__main__":
    for fn in (p01, p02, p03, p04, p05, p06, p07, p08, p09, p10):
        fn()

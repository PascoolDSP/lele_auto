"""Evaluate every formula of the generated proposals with the `formulas` engine and report errors."""
import sys
from pathlib import Path

import formulas

OUT = Path(__file__).resolve().parent.parent / "share" / "out" / "propositions_stock"


def check(path, samples):
    model = formulas.ExcelModel().loads(str(path)).finish()
    sol = model.calculate()
    errors = {}
    values = {}
    for key, rng in sol.items():
        try:
            v = rng.value
        except AttributeError:
            continue
        flat = [x for row in v for x in row] if hasattr(v, "__iter__") else [v]
        for x in flat:
            s = str(x)
            if s.startswith("#") and s not in ("#EMPTY",):
                errors.setdefault(s, []).append(key)
        values[key.upper()] = v
    print(f"== {path.name}: {len(sol)} cells, errors: {sum(len(e) for e in errors.values())}")
    for err, keys in errors.items():
        print("  ", err, keys[:6])
    for label, ref in samples:
        hit = [k for k in values if k.replace("'", "").endswith(ref.upper().replace("'", ""))]
        print(f"   {label}: {values[hit[0]].tolist() if hit else 'n/a'}")


SAMPLES = {
    "01": [("Farine total cmd/reçu/écart", "'OCTOBRE 2026'!T7"), ("reçu", "'OCTOBRE 2026'!U7"), ("écart", "'OCTOBRE 2026'!V7")],
    "02": [("Récap farine total cmd", "'RÉCAP DU MOIS'!G5"), ("dernier stock", "'RÉCAP DU MOIS'!J5")],
    "03": [("Synthèse farine S1", "SYNTHÈSE'!C6"), ("dernier stock", "SYNTHÈSE'!J6"), ("date", "SYNTHÈSE'!K6")],
    "04": [("Metro incomplets", "'RÉCAP FOURNISSEURS'!C5"), ("Transgourmet incomplets", "'RÉCAP FOURNISSEURS'!C7")],
    "05": [("Suggéré farine S1", "'OCTOBRE 2026'!F7")],
    "06": [("Statut farine S1", "'OCTOBRE 2026'!H7")],
    "07": [("Dépense mois", "'TABLEAU DE BORD'!A5"), ("Top1", "'TABLEAU DE BORD'!B11")],
    "08": [("Année farine oct", "'ANNÉE 2026'!L5"), ("moyenne", "'ANNÉE 2026'!P5")],
    "09": [("Valeur stock S1 total", "'OCTOBRE 2026'!E47")],
    "11": [("Farine reçu mois", "'OCTOBRE 2026'!P8"), ("écart", "'OCTOBRE 2026'!Q8")],
    "12": [("Farine habituelle", "METRO!AA7"), ("dernière", "METRO!AB7"), ("fois manqué", "METRO!AC7")],
    "13": [("Récap farine cmd", "'RÉCAP DU MOIS'!C5"), ("manque", "'RÉCAP DU MOIS'!E5")],
    "15": [("État crème", "'DATES LIMITES'!F5"), ("État lot fini", "'DATES LIMITES'!F9")],
    "16": [("Crème périmé", "'BILAN DES PERTES'!C11"), ("Choc raté", "'BILAN DES PERTES'!D14"), ("total choc", "'BILAN DES PERTES'!H14")],
    "17": [("Barre S1", "'OCTOBRE 2026'!T7"), ("Barre S2", "'OCTOBRE 2026'!U7"), ("Évol", "'OCTOBRE 2026'!X7")],
    "18": [("Essentiel 1", "'CHAQUE SEMAINE'!A7"), ("Essentiel 6", "'CHAQUE SEMAINE'!A12"), ("Autre 1", "'UNE FOIS PAR MOIS'!A6")],
    "19": [("Utilisé S1 farine", "'OCTOBRE 2026'!G7"), ("Utilisé S2", "'OCTOBRE 2026'!K7")],
    "20": [("À traiter S1", "'OCTOBRE 2026'!G3"), ("À traiter S3", "'OCTOBRE 2026'!O3")],
    "10": [("Bilan farine", "'OCTOBRE 2026'!C178")],
}

files = sorted(OUT.glob("*.xlsx"))
if len(sys.argv) > 1:
    files = [f for f in files if f.name[:2] in sys.argv[1:]]
for f in files:
    check(f, SAMPLES.get(f.name[:2], []))

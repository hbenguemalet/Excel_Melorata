#!/usr/bin/env python3
"""Migration unique : Artistes.xlsm (VBA + 45 boutons) -> Artistes.xlsx propre, piloté par bot_musique.py.

python optimiser.py chemin/vers/Artistes.xlsm
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
import bot_musique as bm

src = Path(sys.argv[1])
wb = load_workbook(src)                        # sans keep_vba : macros et boutons abandonnés
journal = []

# 1. feuilles cachées vides
for ws in list(wb.worksheets):
    if ws.sheet_state != "visible" and ws.max_row == 1 and ws["A1"].value is None:
        journal.append(f"Feuille cachée vide supprimée : {ws.title}")
        del wb[ws.title]

# 2. espaces parasites dans les noms de feuilles (+ report dans Collection)
for ws in wb.worksheets:
    if ws.title != ws.title.strip():
        journal.append(f"Nom de feuille nettoyé : '{ws.title}' -> '{ws.title.strip()}'")
        ws.title = ws.title.strip()
if "Collection" in wb.sheetnames:
    wc = wb["Collection"]
    for row in wc.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, str) and c.value != c.value.strip():
                c.value = c.value.strip()

# 3. normalisation de chaque feuille artiste
STYLE_FIX = {"rap": "Rap"}
for ws in bm.feuilles_artistes(wb):
    albums = []
    for c in range(1, ws.max_column + 1, bm.BLOC):          # année saisie dans la colonne Style
        st, an = ws.cell(1, c + 1).value, ws.cell(1, c + 2).value
        if isinstance(st, int) and 1900 < st < 2100 and an is None:
            ws.cell(1, c + 2).value, ws.cell(1, c + 1).value = st, None
            journal.append(f"{ws.title} / {ws.cell(1, c).value} : année {st} déplacée de la colonne Style vers Année")
    for a in bm.lire_feuille(ws):
        if not a.sons and len(a.titre) <= 1:                 # résidus type 'S', 'Z'
            journal.append(f"{ws.title} : cellule parasite '{a.titre}' supprimée")
            continue
        if a.style:
            a.style = STYLE_FIX.get(a.style.lower(), a.style.strip())
        albums.append(a)
    bm.ecrire_feuille(ws, bm.trier(albums))

out = src.with_suffix(".xlsx").with_name("Artistes.xlsx")
bm.mise_a_jour(wb)
bm.enregistrer(wb, out)
print("\n".join(journal))
print(f"→ {out}")

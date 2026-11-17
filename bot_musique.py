#!/usr/bin/env python3
"""
bot_musique.py — Gestion de mon classeur de notation d'albums.

Commandes :
  python bot_musique.py                         menu interactif
  python bot_musique.py ajouter "SZA" "SOS"     ajoute un album (données Apple Music)
  python bot_musique.py ajouter "SZA"           liste la discographie et laisse choisir
  python bot_musique.py maj                     trie les albums par année, régénère Sommaire + Collection, exporte Power BI
  python bot_musique.py export                  exporte seulement les tables Power BI (CSV)
  python bot_musique.py completer               récupère les années / styles manquants

Les données (titres, année, style) viennent de l'API publique iTunes Search,
qui interroge le même catalogue que l'application Apple Music. Aucune clé requise.
Seule dépendance : openpyxl  (pip install openpyxl)
"""
from __future__ import annotations

import argparse, csv, difflib, json, os, re, shutil, subprocess, sys, unicodedata, urllib.parse, urllib.request
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L

# --------------------------------------------------------------------------- configuration
DOSSIER = Path(__file__).resolve().parent
FICHIER = DOSSIER / "Artistes.xlsx"
DOSSIER_PBI = DOSSIER / "powerbi"
DOSSIER_SAUV = DOSSIER / "sauvegardes"
GITHUB_REPO = "https://github.com/hbenguemalet/Excel_Melorata"
GITHUB_AUTO = True              # publie automatiquement chaque mise à jour sur GitHub
PAYS = "fr"                     # catalogue Apple Music utilisé
FEUILLES_SYSTEME = {"Sommaire", "Collection"}
BLOC = 5                        # Titre | Note/10 | Featuring | Moyenne | (vide)
LIGNE_SONS = 4
MAX_LIGNE_MOY = 150             # plage bornée de la moyenne (au lieu de B:B)

# Genre Apple Music -> mon style (à compléter librement)
STYLES = {
    "r&b/soul": "R&B", "r&b": "R&B", "contemporary r&b": "R&B", "neo-soul": "R&B",
    "soul": "Soul", "hip-hop/rap": "Rap", "hip-hop": "Rap", "rap": "Rap",
    "rap français": "Rap", "french rap": "Rap", "pop": "Pop", "dance": "Pop", "k-pop": "K-Pop",
    "j-pop": "J-Pop", "alternative": "Alternative", "indie pop": "Alternative",
    "country": "Country", "gospel": "Gospel", "christian & gospel": "Gospel",
    "disco": "Disco", "funk": "Soul", "jazz": "Jazz", "zouk": "Zouk",
    "electronic": "Électro", "rock": "Rock",
}

# Couleurs reprises des macros d'origine
VERT_F, VERT, VERT_C, JAUNE, ORANGE, ROUGE = "009900", "92D050", "ADE255", "FFFF00", "FFC000", "FF0000"
ENTETES = {
    "top_art": "0070C0", "top_alb": "70AD47", "pire_art": "FF9900", "pire_alb": "C00000",
    "rap_fr": "848484", "rap_us": "B22222", "style": "9BC2E6", "pire_style": "E69BC2",
    "coll": "4472C4", "wish": "FF8C00",
}

# --------------------------------------------------------------------------- modèle
@dataclass
class Son:
    titre: str
    note: float | None = None
    featuring: str | None = None
    styles: tuple = (None, None, None)   # styles de cellule conservés lors des réécritures

    @property
    def favori(self) -> bool:
        return self.titre.rstrip().endswith("*")

    @property
    def titre_propre(self) -> str:
        return self.titre.rstrip().rstrip("*").strip()


@dataclass
class Album:
    titre: str
    style: str | None = None
    annee: int | None = None
    sons: list[Son] = field(default_factory=list)

    @property
    def moyenne(self) -> float | None:
        n = [s.note for s in self.sons if s.note is not None]
        return sum(n) / len(n) if n else None


def cle(txt: str) -> str:
    """Clé de comparaison : sans accents, ponctuation ni casse ('Ms.Lauryn Hill' == 'Ms. Lauryn Hill')."""
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", txt.lower())


def nombre(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", ".").strip())
        except ValueError:
            return None
    return None


def feuilles_artistes(wb):
    return [ws for ws in wb.worksheets if ws.title not in FEUILLES_SYSTEME and ws.sheet_state == "visible"]

# --------------------------------------------------------------------------- lecture / écriture d'une feuille artiste
def lire_feuille(ws) -> list[Album]:
    albums = []
    for c in range(1, ws.max_column + 1, BLOC):
        titre = ws.cell(1, c).value
        if titre is None or str(titre).strip() == "":
            continue
        annee = ws.cell(1, c + 2).value
        annee = int(annee) if isinstance(annee, (int, float)) and 1900 < annee < 2100 else None
        style = ws.cell(1, c + 1).value
        alb = Album(str(titre).strip() if not isinstance(titre, (int, float)) else str(titre),
                    str(style).strip() if style not in (None, "") else None, annee)
        for r in range(LIGNE_SONS, ws.max_row + 1):
            t = ws.cell(r, c).value
            if t is None or str(t).strip() == "":
                continue
            cells = [ws.cell(r, c + k) for k in range(3)]
            alb.sons.append(Son(str(t), nombre(cells[1].value),
                                str(cells[2].value).strip() if cells[2].value not in (None, "") else None,
                                tuple(copy(x._style) if x.has_style else None for x in cells)))
        albums.append(alb)
    return albums


def ecrire_feuille(ws, albums: list[Album]):
    ws.delete_rows(1, ws.max_row + 1)
    for k in list(ws.column_dimensions):
        del ws.column_dimensions[k]
    lien = ws.cell(2, 1, "← Sommaire")
    lien.hyperlink = "#'Sommaire'!A1"
    lien.font = Font(color="0563C1", underline="single", italic=True, size=9)
    for i, a in enumerate(albums):
        c = 1 + i * BLOC
        ws.cell(1, c, a.titre)
        ws.cell(1, c + 1, a.style)
        ws.cell(1, c + 2, a.annee)
        for k, h in enumerate(("Titre", "Note/10", "Featuring")):
            ws.cell(3, c + k, h)
        col = L(c + 1)
        m = ws.cell(3, c + 3, f'=IFERROR(AVERAGE({col}{LIGNE_SONS}:{col}{MAX_LIGNE_MOY}),"")')
        m.number_format = "0.00"
        for j, s in enumerate(a.sons):
            r = LIGNE_SONS + j
            for k, v in enumerate((s.titre, s.note, s.featuring)):
                cell = ws.cell(r, c + k, v)
                if s.styles[k] is not None:
                    cell._style = copy(s.styles[k])
        ws.column_dimensions[L(c)].width = 26
        ws.column_dimensions[L(c + 2)].width = 18


def trier(albums):
    """Tri chronologique (comme la macro d'origine). Si une année manque, l'ordre est conservé :
    lance `completer` pour récupérer les années manquantes depuis Apple Music."""
    if any(a.annee is None for a in albums):
        return list(albums)
    return sorted(albums, key=lambda a: a.annee)

# --------------------------------------------------------------------------- mise en forme utilitaire
def fill(hexa):
    return PatternFill("solid", fgColor=hexa)


def couleur_note(n):
    return VERT_F if n >= 9 else VERT if n >= 8 else VERT_C if n >= 7 else JAUNE if n >= 5.5 else ORANGE if n >= 4 else ROUGE


def entete(ws, r, c, titre, cols, couleur, blanc=True):
    ws.cell(r, c, titre)
    for k, h in enumerate(cols):
        ws.cell(r + 1, c + k, h)
    for rr in (r, r + 1):
        for k in range(len(cols)):
            x = ws.cell(rr, c + k)
            x.font = Font(bold=True, color="FFFFFF" if blanc else "000000")
            x.fill = fill(couleur)
            x.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + len(cols) - 1)


def lien_feuille(cell, feuille):
    cell.hyperlink = f"#'{feuille}'!A1"
    cell.font = Font(color="0563C1", underline="single")


def classement(ws, r, c, lignes, avec_album=True):
    """lignes : [(note, artiste, album)] déjà triées."""
    for i, (note, art, alb) in enumerate(lignes):
        ws.cell(r + i, c, f"{i + 1}e")
        n = ws.cell(r + i, c + 1, round(note, 4))
        n.number_format = "0.00"
        n.fill = fill(couleur_note(note))
        a = ws.cell(r + i, c + 2, art)
        lien_feuille(a, art)
        if avec_album:
            lien_feuille(ws.cell(r + i, c + 3, alb), art)

# --------------------------------------------------------------------------- collecte globale
def collecter(wb):
    data = {}
    for ws in feuilles_artistes(wb):
        data[ws.title] = lire_feuille(ws)
    return data


def remplacer_feuille(wb, nom, index):
    if nom in wb.sheetnames:
        del wb[nom]
    return wb.create_sheet(nom, index)

# --------------------------------------------------------------------------- Sommaire
def construire_sommaire(wb, data):
    ws = remplacer_feuille(wb, "Sommaire", 0)
    albums = [(a.moyenne, art, a.titre, a.style) for art, als in data.items() for a in als if a.moyenne is not None]
    artistes = {}
    for m, art, _, _ in albums:
        artistes.setdefault(art, []).append(m)
    moy_art = {k: sum(v) / len(v) for k, v in artistes.items()}
    styles = {}
    for m, art, alb, st in albums:
        if st:
            styles.setdefault(st, []).append((m, art, alb))

    ws["A1"] = "Classement musical"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"Mis à jour le {datetime.now():%d/%m/%Y %H:%M} — généré par bot_musique.py"
    ws["A2"].font = Font(italic=True, size=9, color="808080")
    ws["A3"], ws["B3"] = "Albums notés :", len(albums)
    ws["A4"], ws["B4"] = "Artistes notés :", len(moy_art)
    ws["A5"] = "Albums par style :"
    r = 6
    for st, v in sorted(styles.items(), key=lambda x: -len(x[1])):
        ws.cell(r, 1, f"{st}:"), ws.cell(r, 2, len(v))
        r += 1
    r += 1
    ws.cell(r, 1, "Liste des artistes").font = Font(bold=True, size=12)
    for art in data:
        r += 1
        lien_feuille(ws.cell(r, 1, art), art)

    top_art = sorted(((m, a, None) for a, m in moy_art.items()), key=lambda x: -x[0])
    top_alb = sorted(((m, a, al) for m, a, al, _ in albums), key=lambda x: -x[0])
    entete(ws, 4, 4, "Top 20 des meilleurs artistes", ("Place", "Note", "Artiste"), ENTETES["top_art"])
    classement(ws, 6, 4, top_art[:20], False)
    entete(ws, 29, 4, "Top 20 des pires artistes", ("Place", "Note", "Artiste"), ENTETES["pire_art"])
    classement(ws, 31, 4, top_art[::-1][:20], False)
    entete(ws, 4, 8, "Top 20 des meilleurs albums", ("Place", "Note", "Artiste", "Album"), ENTETES["top_alb"])
    classement(ws, 6, 8, top_alb[:20])
    entete(ws, 29, 8, "Top 20 des pires albums", ("Place", "Note", "Artiste", "Album"), ENTETES["pire_alb"])
    classement(ws, 31, 8, top_alb[::-1][:20])

    c = 13
    for st, coul in (("RAP FR", "rap_fr"), ("RAP US", "rap_us")):          # repris de la macro d'origine
        arts = {a for _, a, _, s in albums if s and s.upper() == st}
        if arts:
            entete(ws, 4, c, f"Top Artistes {st}", ("Place", "Note", "Artiste"), ENTETES[coul])
            classement(ws, 6, c, sorted(((moy_art[a], a, None) for a in arts), key=lambda x: -x[0])[:20], False)
            c += 4
    for st, v in sorted(styles.items(), key=lambda x: -len(x[1])):
        v = sorted(v, key=lambda x: -x[0])
        entete(ws, 4, c, f"Top 10 {st}", ("Place", "Note", "Artiste", "Album"), ENTETES["style"], blanc=False)
        classement(ws, 6, c, v[:10])
        rp = 6 + min(10, len(v)) + 2
        entete(ws, rp, c, f"Top 10 pires {st}", ("Place", "Note", "Artiste", "Album"), ENTETES["pire_style"], blanc=False)
        classement(ws, rp + 2, c, v[::-1][:10])
        c += 5

    # Meilleurs sons (titres terminés par *) par style
    favoris = {}
    for art, als in data.items():
        for a in als:
            for s in a.sons:
                if s.favori and a.style:
                    favoris.setdefault(a.style, []).append((art, s.titre_propre, a.titre))
    r = max(r, 52) + 4
    for st, sons in favoris.items():
        ws.cell(r, 1, f"Meilleurs Sons : {st}").font = Font(bold=True, size=12)
        for k, h in enumerate(("Artiste", "Titre de la chanson", "Album")):
            ws.cell(r + 1, 1 + k, h).font = Font(bold=True)
        for i, (art, t, alb) in enumerate(sons):
            lien_feuille(ws.cell(r + 2 + i, 1, art), art)
            ws.cell(r + 2 + i, 2, t), ws.cell(r + 2 + i, 3, alb)
        r += len(sons) + 4

    for col, w in {"A": 24, "B": 30, "C": 22, "D": 7, "E": 7, "F": 20, "G": 3, "H": 7, "I": 7, "J": 20, "K": 30}.items():
        ws.column_dimensions[col].width = w
    for cc in range(13, c + 1):
        ws.column_dimensions[L(cc)].width = (7, 7, 18, 28, 3)[(cc - 13) % 5]
    ws.freeze_panes = "A4"
    return ws

# --------------------------------------------------------------------------- Collection (CD / Vinyle)
def couleur_vx(cell):
    v = str(cell.value or "").strip().upper()
    cell.alignment = Alignment(horizontal="center")
    cell.font = Font(bold=v in "VXW" and v != "", color={"V": "008000", "X": "FF0000", "W": "FF8C00"}.get(v, "000000"))


def construire_collection(wb, data):
    ancien = {}
    if "Collection" in wb.sheetnames:
        wc = wb["Collection"]
        for r in range(2, wc.max_row + 1):
            a, b = wc.cell(r, 1).value, wc.cell(r, 2).value
            if a is not None and b is not None:
                ancien[(cle(a), cle(b))] = (wc.cell(r, 3).value, wc.cell(r, 4).value)
    ws = remplacer_feuille(wb, "Collection", 1)
    for k, h in enumerate(("Artiste", "Album", "CD", "Vinyle")):
        x = ws.cell(1, 1 + k, h)
        x.font, x.fill, x.alignment = Font(bold=True, color="FFFFFF"), fill(ENTETES["coll"]), Alignment(horizontal="center")
    possedes, wish, r = [], [], 2
    for art, als in data.items():
        for a in als:
            cd, vin = ancien.get((cle(art), cle(a.titre)), (None, None))
            cd, vin = (cd or "X"), (vin or "X")
            ws.cell(r, 1, art), ws.cell(r, 2, a.titre)
            couleur_vx(ws.cell(r, 3, cd)), couleur_vx(ws.cell(r, 4, vin))
            r += 1
            if "V" in (cd, vin):
                possedes.append((art, a.titre, "V" if cd == "V" else "", "V" if vin == "V" else ""))
            m = a.moyenne or 0
            w_cd, w_vin = m > 8.5 and cd != "V", m > 9 and vin != "V"
            if w_cd or w_vin:
                wish.append((art, a.titre, "W" if w_cd else "", "W" if w_vin else ""))
    bord = Border(*(Side("thin"),) * 4)
    for c0, titre, coul, cols, lignes in (
            (6, "Albums Possédés (V)", "coll", ("Artiste", "Album", "CD", "Vinyle"), possedes),
            (11, "Wishlist (W)", "wish", ("Artiste", "Album", "CD (>8.5)", "Vinyle (>9)"), wish)):
        entete(ws, 11, c0, titre, cols, ENTETES[coul])
        for i, ligne in enumerate(lignes):
            for k, v in enumerate(ligne):
                x = ws.cell(13 + i, c0 + k, v)
                x.border = bord
                if k >= 2:
                    couleur_vx(x)
    for col, w in zip("ABCDEFGHIJKLMN", (20, 34, 7, 8, 3, 20, 34, 7, 8, 3, 20, 34, 10, 11)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{max(r - 1, 1)}"

# --------------------------------------------------------------------------- export Power BI (modèle en étoile)
def separer_feats(txt):
    return [p.strip() for p in re.split(r",|&| et | and | x ", txt or "") if p.strip()]


def exporter(data, dossier=DOSSIER_PBI, collection=None):
    dossier.mkdir(exist_ok=True)
    collection = collection or {}
    t_art, t_alb, t_son, t_feat = [], [], [], []
    alb_id = son_id = 0
    for art_id, (art, als) in enumerate(data.items(), 1):
        moys = [a.moyenne for a in als if a.moyenne is not None]
        t_art.append([art_id, art, len(als), len(moys), round(sum(moys) / len(moys), 4) if moys else None])
        for a in als:
            alb_id += 1
            notes = [s.note for s in a.sons if s.note is not None]
            cd, vin = collection.get((cle(art), cle(a.titre)), ("X", "X"))
            t_alb.append([alb_id, art_id, a.titre, a.annee, a.style, len(a.sons), len(notes),
                          round(a.moyenne, 4) if a.moyenne is not None else None, cd == "V", vin == "V"])
            for p, s in enumerate(a.sons, 1):
                son_id += 1
                t_son.append([son_id, alb_id, art_id, p, s.titre_propre, s.note, s.featuring, s.favori])
                for f in separer_feats(s.featuring):
                    t_feat.append([son_id, f])
    tables = {
        "dim_artiste": (["artiste_id", "artiste", "nb_albums", "nb_albums_notes", "note_moyenne"], t_art),
        "dim_album": (["album_id", "artiste_id", "album", "annee", "style", "nb_titres", "nb_titres_notes",
                       "note_moyenne", "cd_possede", "vinyle_possede"], t_alb),
        "fait_son": (["son_id", "album_id", "artiste_id", "piste", "titre", "note", "featuring", "favori"], t_son),
        "pont_featuring": (["son_id", "artiste_featuring"], t_feat),
    }
    def fr(v):                                       # CSV au format français (Power BI / Excel FR)
        return ("1" if v else "0") if isinstance(v, bool) else str(v).replace(".", ",") if isinstance(v, float) else v
    for nom, (cols, rows) in tables.items():
        rows = [[fr(v) for v in r] for r in rows]
        with open(dossier / f"{nom}.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(cols)
            w.writerows(rows)
    return {k: len(v[1]) for k, v in tables.items()}


def lire_collection(wb):
    out = {}
    if "Collection" in wb.sheetnames:
        wc = wb["Collection"]
        for r in range(2, wc.max_row + 1):
            a, b = wc.cell(r, 1).value, wc.cell(r, 2).value
            if a is not None and b is not None:
                out[(cle(a), cle(b))] = (str(wc.cell(r, 3).value or "X").upper(), str(wc.cell(r, 4).value or "X").upper())
    return out

# --------------------------------------------------------------------------- ouverture / sauvegarde sûres
def git(*args, cwd=None):
    return subprocess.run(["git", *args], cwd=cwd or FICHIER.parent, capture_output=True, text=True)


def depot_git_ok():
    if not GITHUB_AUTO:
        return False
    if shutil.which("git") is None:
        print("ℹ Git n'est pas installé : publication GitHub ignorée (https://git-scm.com).")
        return False
    if git("rev-parse", "--is-inside-work-tree").returncode:
        print(f"ℹ Ce dossier n'est pas relié à GitHub. Fais une fois : git clone {GITHUB_REPO}")
        return False
    return True


def recuperer_github():
    """Récupère la dernière version avant de modifier (au cas où tu as changé quelque chose ailleurs)."""
    if depot_git_ok():
        r = git("pull", "--ff-only", "--quiet")
        if r.returncode:
            print("⚠ Impossible de récupérer la version GitHub :", r.stderr.strip()[-200:])


def publier_github(message):
    if not depot_git_ok():
        return
    git("add", "-A")
    if not git("status", "--porcelain").stdout.strip():
        return
    git("commit", "-m", message)
    r = git("push")
    print("☁ GitHub à jour" if r.returncode == 0 else f"⚠ Push GitHub échoué : {r.stderr.strip()[-300:]}")


def ouvrir(fichier=None):
    fichier = fichier or FICHIER
    recuperer_github()
    verrou = fichier.with_name("~$" + fichier.name)
    if verrou.exists():
        sys.exit("⚠ Le classeur est ouvert dans Excel : ferme-le puis relance.")
    return load_workbook(fichier)


def enregistrer(wb, fichier=None, message="Mise à jour du classement"):
    fichier = fichier or FICHIER
    dossier_sauv = fichier.parent / DOSSIER_SAUV.name
    dossier_sauv.mkdir(exist_ok=True)
    if fichier.exists():
        shutil.copy2(fichier, dossier_sauv / f"{fichier.stem}_{datetime.now():%Y%m%d_%H%M%S}{fichier.suffix}")
        sauv = sorted(dossier_sauv.glob(f"{fichier.stem}_*"))
        for vieux in sauv[:-10]:                        # garde les 10 dernières
            vieux.unlink()
    wb.calculation.fullCalcOnLoad = True                # Excel recalcule les moyennes à l'ouverture
    wb.save(fichier)
    publier_github(f"{message} ({datetime.now():%d/%m/%Y %H:%M})")


def mise_a_jour(wb, verbose=True):
    for ws in feuilles_artistes(wb):
        als = lire_feuille(ws)
        if [a.titre for a in als] != [a.titre for a in trier(als)]:
            ecrire_feuille(ws, trier(als))
    data = collecter(wb)
    construire_sommaire(wb, data)
    construire_collection(wb, data)
    wb.active = 0
    stats = exporter(data, FICHIER.parent / DOSSIER_PBI.name, lire_collection(wb))
    if verbose:
        print("✔ Sommaire, Collection et tables Power BI à jour :", ", ".join(f"{k}={v}" for k, v in stats.items()))
    return data

# --------------------------------------------------------------------------- Apple Music (API iTunes Search)
def itunes(endpoint, **params):
    params.setdefault("country", PAYS)
    url = f"https://itunes.apple.com/{endpoint}?{urllib.parse.urlencode(params)}"
    fixture = os.environ.get("BOT_MUSIQUE_FIXTURE")       # tests hors-ligne
    if fixture:
        return json.load(open(fixture, encoding="utf-8"))[url.split("?")[0].rsplit("/", 1)[1] + ":" + str(params.get("id", params.get("term")))]
    req = urllib.request.Request(url, headers={"User-Agent": "bot-musique/1.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.load(r)["results"]


RE_FEAT = re.compile(r"\s*[\(\[](?:feat\.?|ft\.?|featuring|with|avec)\s+([^\)\]]+)[\)\]]", re.I)


def nettoyer_titre(nom):
    feats = RE_FEAT.findall(nom)
    return RE_FEAT.sub("", nom).strip(), ", ".join(f.strip() for f in feats) or None


def nettoyer_album(nom):
    return re.sub(r"\s+-\s+(Single|EP)$", "", nom).strip()


def chercher_albums(artiste, album=None, tout=False):
    if album:
        res = itunes("search", term=f"{artiste} {album}", entity="album", limit=25)
    else:
        arts = itunes("search", term=artiste, entity="musicArtist", limit=5)
        if not arts:
            return []
        best = max(arts, key=lambda a: difflib.SequenceMatcher(None, cle(a["artistName"]), cle(artiste)).ratio())
        res = [r for r in itunes("lookup", id=best["artistId"], entity="album", limit=200)
               if r.get("wrapperType") == "collection"]
    res = [r for r in res if r.get("wrapperType") == "collection"
           and difflib.SequenceMatcher(None, cle(r["artistName"]), cle(artiste)).ratio() > 0.6]
    if not tout:
        res = [r for r in res if not re.search(r" - (Single|EP)$", r["collectionName"]) and r.get("trackCount", 0) >= 6]
    if album:
        res.sort(key=lambda r: -difflib.SequenceMatcher(None, cle(nettoyer_album(r["collectionName"])), cle(album)).ratio())
    else:
        res.sort(key=lambda r: r.get("releaseDate", ""))
    vus, uniques = set(), []
    for r in res:                                        # doublons (versions explicit / clean)
        k = (cle(r["collectionName"]), r.get("releaseDate", "")[:4])
        if k not in vus:
            vus.add(k), uniques.append(r)
    return uniques


def album_depuis_apple(collection) -> Album:
    pistes = [t for t in itunes("lookup", id=collection["collectionId"], entity="song")
              if t.get("wrapperType") == "track" and t.get("kind") == "song"]
    pistes.sort(key=lambda t: (t.get("discNumber", 1), t.get("trackNumber", 0)))
    principal = cle(collection["artistName"])
    sons = []
    for t in pistes:
        titre, feat = nettoyer_titre(t["trackName"])
        if not feat and cle(t.get("artistName", "")) != principal:     # ex. "Kanye West & JAY-Z"
            autres = [a for a in separer_feats(t.get("artistName", "")) if cle(a) != principal]
            feat = ", ".join(autres) or None
        sons.append(Son(titre, None, feat))
    genre = collection.get("primaryGenreName", "")
    return Album(nettoyer_album(collection["collectionName"]), STYLES.get(genre.lower(), genre or None),
                 int(collection["releaseDate"][:4]) if collection.get("releaseDate") else None, sons)


def nom_feuille(nom):
    return re.sub(r"[\[\]:*?/\\]", "", nom).strip()[:31]


def trouver_feuille(wb, artiste):
    noms = [ws.title for ws in feuilles_artistes(wb)]
    exact = {cle(n): n for n in noms}
    if cle(artiste) in exact:
        return exact[cle(artiste)]
    proche = difflib.get_close_matches(cle(artiste), list(exact), n=1, cutoff=0.85)
    return exact[proche[0]] if proche else None


def ajouter_album(wb, artiste, alb: Album):
    nom = trouver_feuille(wb, artiste)
    if nom is None:
        nom = nom_feuille(artiste)
        arts = feuilles_artistes(wb)
        index = next((wb.worksheets.index(w) for w in arts if w.title.casefold() > nom.casefold()), len(wb.worksheets))
        wb.create_sheet(nom, index)
        print(f"➕ Nouvelle feuille artiste : {nom}")
    ws = wb[nom]
    albums = lire_feuille(ws)
    if any(cle(a.titre) == cle(alb.titre) for a in albums):
        print(f"↷ « {alb.titre} » existe déjà dans {nom}, ignoré.")
        return False
    ecrire_feuille(ws, trier(albums + [alb]))
    print(f"✔ {nom} — {alb.titre} ({alb.annee}, {alb.style}) : {len(alb.sons)} titres ajoutés")
    return True


def choisir(options, multiple):
    for i, r in enumerate(options, 1):
        print(f"  {i:>2}. {nettoyer_album(r['collectionName'])}  ({r.get('releaseDate', '')[:4]}, "
              f"{r.get('trackCount', '?')} titres, {r.get('primaryGenreName', '')})")
    rep = input("Numéro(s) à ajouter (ex. 1  ou  1,3  ou  tous ; Entrée = annuler) : ").strip().lower()
    if not rep:
        return []
    if rep == "tous" and multiple:
        return options
    idx = [int(x) for x in re.findall(r"\d+", rep) if 0 < int(x) <= len(options)]
    return [options[i - 1] for i in (idx if multiple else idx[:1])]


def cmd_ajouter(artiste, album=None, choix=None, tout=False, maj=True):
    wb = ouvrir()
    try:
        res = chercher_albums(artiste, album, tout)
    except OSError as e:
        sys.exit(f"Impossible de joindre Apple Music ({e}). Vérifie ta connexion.")
    if not res:
        sys.exit("Aucun album trouvé. Essaie une autre orthographe ou --tout pour inclure EP/singles.")
    if choix:
        selection = [res[i - 1] for i in choix if 0 < i <= len(res)]
    elif album and len(res) >= 1 and cle(nettoyer_album(res[0]["collectionName"])) == cle(album):
        selection = [res[0]]
    else:
        print(f"Résultats Apple Music pour « {artiste}{' – ' + album if album else ''} » :")
        selection = choisir(res[:30], multiple=album is None)
    ajoutes = [r for r in selection if ajouter_album(wb, artiste, album_depuis_apple(r))]
    if ajoutes:
        if maj:
            mise_a_jour(wb)
        enregistrer(wb, message=f"Ajout {artiste} : " + ", ".join(nettoyer_album(r["collectionName"]) for r in ajoutes))
        print(f"💾 Enregistré dans {FICHIER.name}. Il ne reste plus qu'à noter les sons !")

def cmd_completer():
    """Complète les années / styles manquants des albums déjà présents (sans toucher aux notes)."""
    wb = ouvrir()
    modifs = 0
    for ws in feuilles_artistes(wb):
        albums = lire_feuille(ws)
        manquants = [a for a in albums if a.annee is None or not a.style]
        if not manquants:
            continue
        try:
            candidats = chercher_albums(ws.title, tout=True)
        except OSError as e:
            sys.exit(f"Impossible de joindre Apple Music ({e}).")
        for a in manquants:
            best = max(candidats, default=None, key=lambda r: difflib.SequenceMatcher(
                None, cle(nettoyer_album(r["collectionName"])), cle(a.titre.rstrip("*"))).ratio())
            if not best or difflib.SequenceMatcher(None, cle(nettoyer_album(best["collectionName"])),
                                                   cle(a.titre.rstrip("*"))).ratio() < 0.85:
                print(f"  ? {ws.title} — {a.titre} : introuvable, à compléter à la main")
                continue
            if a.annee is None and best.get("releaseDate"):
                a.annee = int(best["releaseDate"][:4])
            if not a.style:
                g = best.get("primaryGenreName", "")
                a.style = STYLES.get(g.lower(), g or None)
            modifs += 1
            print(f"  ✔ {ws.title} — {a.titre} : {a.annee}, {a.style}")
        ecrire_feuille(ws, trier(albums))
    if modifs:
        mise_a_jour(wb)
        enregistrer(wb, message=f"Années/styles complétés ({modifs} albums)")
    print(f"{modifs} album(s) complété(s).")


# --------------------------------------------------------------------------- interface
def menu():
    while True:
        try:
            if not menu_une_fois():
                break
        except SystemExit as e:
            print(e)


def menu_une_fois():
    print("\n🎵  Bot musique\n  1. Ajouter un album\n  2. Ajouter des albums d'un artiste (discographie)\n"
          "  3. Mettre à jour Sommaire / Collection / Power BI\n"
          "  4. Compléter les années / styles manquants\n  5. Quitter")
    c = input("Choix : ").strip()
    if c == "1":
        cmd_ajouter(input("Artiste : ").strip(), input("Album : ").strip())
    elif c == "2":
        cmd_ajouter(input("Artiste : ").strip())
    elif c == "3":
        wb = ouvrir()
        mise_a_jour(wb)
        enregistrer(wb)
    elif c == "4":
        cmd_completer()
    else:
        return False
    return True


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--fichier", type=Path, help="chemin du classeur (défaut : Artistes.xlsx à côté du script)")
    sp = p.add_subparsers(dest="cmd")
    a = sp.add_parser("ajouter", help="ajouter un album ou une discographie")
    a.add_argument("artiste")
    a.add_argument("album", nargs="?")
    a.add_argument("--choix", type=lambda s: [int(x) for x in s.split(",")], help="numéros à prendre sans demander")
    a.add_argument("--tout", action="store_true", help="inclure EP et singles")
    a.add_argument("--sans-maj", action="store_true", help="ne pas régénérer Sommaire/Collection")
    sp.add_parser("maj", help="trier, régénérer Sommaire + Collection, exporter Power BI")
    sp.add_parser("export", help="exporter uniquement les CSV Power BI")
    sp.add_parser("completer", help="récupérer années/styles manquants depuis Apple Music")
    args = p.parse_args()
    global FICHIER
    if args.fichier:
        FICHIER = args.fichier.resolve()
    if args.cmd == "ajouter":
        cmd_ajouter(args.artiste, args.album, args.choix, args.tout, not args.sans_maj)
    elif args.cmd == "maj":
        wb = ouvrir(FICHIER)
        mise_a_jour(wb)
        enregistrer(wb, FICHIER)
    elif args.cmd == "export":
        wb = load_workbook(FICHIER, read_only=False)
        print(exporter(collecter(wb), FICHIER.parent / DOSSIER_PBI.name, lire_collection(wb)))
    elif args.cmd == "completer":
        cmd_completer()
    else:
        menu()


if __name__ == "__main__":
    main()

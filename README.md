# 🎵 Notation d'albums — Excel + Python + Power BI

Projet perso de data analyse : je note chaque titre des albums que j'écoute (/10), le classeur calcule
les moyennes, des classements (artistes, albums, styles), ma collection CD/vinyles et une wishlist.
Un bot Python récupère automatiquement les tracklists depuis Apple Music et produit un modèle de données
propre pour Power BI.

## Chiffres
60 artistes · 231 albums · 1 982 titres · 534 featurings

## Architecture
```
Apple Music (API iTunes Search)
        │  bot_musique.py ajouter
        ▼
Artistes.xlsx  ── saisie des notes (1 feuille par artiste, 1 bloc de 5 colonnes par album)
        │  bot_musique.py maj
        ├──► Sommaire   (tops 20 / pires, top 10 par style, meilleurs sons ★)
        ├──► Collection (CD / vinyle possédés, wishlist auto : note > 8,5 CD, > 9 vinyle)
        └──► powerbi/*.csv  modèle en étoile
              dim_artiste ─┐
              dim_album ───┼── fait_son ── pont_featuring
```

## Utilisation
```bash
pip install -r requirements.txt
python bot_musique.py                      # menu (ou double-clic sur lancer_bot.bat)
python bot_musique.py ajouter "SZA" "SOS"  # ajoute l'album : titres, année, style, featurings
python bot_musique.py ajouter "SZA"        # liste la discographie, choix multiple
python bot_musique.py completer            # récupère les années / styles manquants
python bot_musique.py maj                  # trie par année, régénère Sommaire + Collection + CSV
```
Excel doit être fermé pendant l'exécution. Une sauvegarde horodatée est faite à chaque écriture (10 gardées).

## Conventions de saisie
- Ligne 1 : album | style | année — ligne 3 : en-têtes + moyenne — à partir de la ligne 4 : titres
- Un `*` à la fin d'un titre = favori (remonte dans « Meilleurs sons »)
- Collection : `V` = possédé, `X` = non possédé

## Ce que la migration depuis l'ancien .xlsm a corrigé (`optimiser.py`)
- 5 modules VBA + 45 boutons remplacés par un script Python versionnable (macros d'origine dans `legacy_vba/`)
- Moyennes `AVERAGE(B:B)` (colonne entière) → plages bornées avec `IFERROR` (plus de `#DIV/0!`)
- Notes saisies en texte (`"7.5"`) ignorées par Excel → converties en nombres (moyenne de *Rupture* : 9 → 8,44)
- Moyenne codée en dur et périmée (*Sonder Son*) → formule
- Années saisies dans la colonne Style, styles en doublon (`RAP`/`Rap`), espaces parasites dans les noms
- Fichier : 384 Ko → ~140 Ko

## Power BI
Obtenir les données → Dossier → `powerbi/` (ou chaque CSV), séparateur `;`, décimale `,`.
Relations : `dim_artiste[artiste_id]` 1→* `dim_album`, `dim_album[album_id]` 1→* `fait_son`,
`fait_son[son_id]` 1→* `pont_featuring`.

# IPCAL data dictionary

Voir `CLAUDE.md` pour le contexte complet du projet (Claude Code le charge
automatiquement au démarrage d'une session dans ce dossier).

## Démarrage rapide (une nouvelle année)

```bash
# 1. Déposer les fichiers bruts reçus (PDF + Excel) dans data/<année>/raw/
mkdir -p data/2023/raw
cp ~/Téléchargements/*2023* data/2023/raw/

# 1bis. Vérifier ce qui manque avant d'aller plus loin (utile aussi en routine,
#       sur toute la série, pour repérer un trou de collecte type "Bruxelles 2016")
python3 scripts/00_check_sources.py

# 2. Alléger les PDF (texte OCR uniquement)
python3 scripts/01_extract_pdf_text.py --annee 2023

# 3. Créer data/2023/manifest.json (voir exemple dans CLAUDE.md), puis :
python3 scripts/02_parse_pdf_structure.py --annee 2023

# 4. Construire les enregistrements (vérifier d'abord les index de colonnes Excel dans
#    l'Excel maître — voir CLAUDE.md section "Excel maître annuel" — et les surcharger
#    dans manifest.json si besoin, sous la clé "excel_columns")
python3 scripts/03_build_dictionary.py --annee 2023

# 5. Générer le classeur final (une année, ou toute la série via glob)
python3 scripts/04_write_excel.py data/2023/records.json -o IPCAL_data_dictionary_2023.xlsx
python3 scripts/04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_full.xlsx

# 6. (multi-année) Historique par variable + détection de dérive sémantique
python3 scripts/05_build_historique.py data/*/records.json -o data/historique.json
python3 scripts/06_write_historique_excel.py data/historique.json -o IPCAL_historique_variables.xlsx
```

`--annee <année>` (étapes 2-4) résout automatiquement les chemins (`raw/`, `manifest.json`,
`pdf_struct.json`, `records.json`, l'Excel maître) via `config.json` à la racine du dépôt —
pratique pour un run local, sans rien changer au code. Une option de chemin explicite
(`--excel`, `--struct`, `-o`, ...) garde toujours priorité sur ce que `--annee` déduirait.
Si votre disposition locale de dossiers diffère (pas de sous-dossier `raw/`, données
stockées ailleurs que dans `data/`, etc.), éditez `config.json` plutôt que les scripts —
voir le commentaire en tête de `scripts/_layout.py` pour le fonctionnement des gabarits.

Demandez simplement à Claude Code de faire tourner ces étapes pour une nouvelle année
— il a le contexte nécessaire dans `CLAUDE.md` pour adapter le pipeline si un fichier
sort un peu de la convention habituelle (nommage différent, colonnes Excel décalées,
vrai PDF au lieu d'un zip, etc.).

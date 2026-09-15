# IPCAL data dictionary

Voir `CLAUDE.md` pour le contexte complet du projet (Claude Code le charge
automatiquement au démarrage d'une session dans ce dossier).

## Démarrage rapide (une nouvelle année)

```bash
# 1. Déposer les fichiers bruts reçus (PDF/zip + Excel) dans data/<année>/raw/
mkdir -p data/2023/raw
cp ~/Téléchargements/*2023* data/2023/raw/

# 2. Alléger les PDF (texte OCR uniquement)
python3 scripts/01_extract_pdf_text.py --dir data/2023/raw

# 3. Créer data/2023/manifest.json (voir exemple dans CLAUDE.md), puis :
python3 scripts/02_parse_pdf_structure.py data/2023/manifest.json -o data/2023/pdf_struct.json

# 4. Construire les enregistrements (vérifier d'abord les index de colonnes Excel,
#    voir CLAUDE.md section "Excel maître annuel")
python3 scripts/03_build_dictionary.py \
  --excel data/2023/raw/IPCAL_2023.xlsx --sheet IPCAL_Codes \
  --struct data/2023/pdf_struct.json \
  --annee-revenus 2023 --exercice 2024 \
  -o data/2023/records.json

# 5. Générer le classeur final (une année, ou toute la série via glob)
python3 scripts/04_write_excel.py data/2023/records.json -o IPCAL_data_dictionary_2023.xlsx
python3 scripts/04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_full.xlsx
```

Demandez simplement à Claude Code de faire tourner ces étapes pour une nouvelle année
— il a le contexte nécessaire dans `CLAUDE.md` pour adapter le pipeline si un fichier
sort un peu de la convention habituelle (nommage différent, colonnes Excel décalées,
vrai PDF au lieu d'un zip, etc.).

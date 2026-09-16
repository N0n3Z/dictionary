# IPCAL — Data dictionary de la déclaration d'impôt des personnes physiques (Belgique)

## Objectif du projet
Produire un data dictionary structuré, multi-année, pour les données micro d'IPCAL
(déclaration à l'impôt des personnes physiques belge). Revenus 2017-2023 traités et
validés (7 années, voir `data/`) ; l'objectif maintenant est d'étendre à la série
historique complète (2014 → aujourd'hui, avec un cas particulier avant 2014 où il n'y
avait qu'un seul PDF, pas un par région) au fur et à mesure que les fichiers bruts
manquants (2014-2016, et l'année test revenus 2024/exercice 2025 déjà validée
manuellement mais pas encore réintégrée ici) sont déposés dans `data/<année>/raw/`.

## Sources (deux par année)
1. **PDF préparatoire** : un par région (BXL/RF/RW) × partie (1=classique, 2=indépendants),
   plus deux versions "non-résidents" (INR, simplifiées, parties 1 et 2). Avant 2014 :
   un seul PDF (pas de découpage régional).
   - **Piège découvert** : ces "PDF" sont en réalité des **archives ZIP** déguisées
     (vérifier via les magic bytes `PK\x03\x04`, pas l'extension). Elles contiennent
     `N.jpeg` (image de la page N) + `N.txt` (texte OCR de la page N) + `manifest.json`.
   - Le texte OCR ne représente qu'~2,5 % du poids total (le reste = images de pages).
     `scripts/01_extract_pdf_text.py` extrait uniquement le texte (x37-41 plus léger),
     avec un mode de secours `pdftotext -layout` si le fichier est un vrai PDF (années
     anciennes, probablement).
2. **Excel maître annuel** (`IPCAL_YYYY.xlsx` ou nom similaire) : liste complète des
   codes IPCAL, y compris les variables **calculées/administratives** absentes du PDF.
   - Feuille utile : celle qui contient les codes (`IPCAL_Codes` en 2024). Il existe
     souvent une seconde feuille (`Sheet1` en 2024) quasi vide — vérifier avant de
     l'ignorer purement et simplement, mais ne pas supposer qu'elle est toujours vide.
   - En-tête sur 3 lignes, données à partir de la ligne 4 (**à revérifier chaque année** :
     imprimer les 5 premières lignes avant de lancer le pipeline).
   - Colonnes clés (index 0-based observés en 2024, **peuvent décaler d'une année à
     l'autre — toujours vérifier**) : 2=IPCAL_A(N-1), 4=IPCAL_B(N-1), 5=Décl_A(N),
     6=IPCAL_A(N), 7=Décl_B(N), 8=IPCAL_B(N), 9=IPType, 14-17=hiérarchie NL (4 niveaux),
     18-21=hiérarchie FR (4 niveaux).

## Règles de codage (stables dans le temps, a priori)
- Code déclaration (4 chiffres) → code IPCAL : 1er chiffre remplacé par une lettre.
  `1→A, 2→B, 3→C, 4→D`. Cible = toujours les codes commençant par une lettre.
- **A↔B** = fédéral déclaré (titulaire/conjoint) ; **C↔D** = régional déclaré
  (titulaire/conjoint) ; **E/F, G/H, I/J, K/L** = variables calculées/administratives
  (pas de code déclaration), toujours en paires titulaire/conjoint.
- Convention titulaire/conjoint : préfixes `ACEGIK` = conjoint 1 (titulaire),
  `BDFHJL` = conjoint 2.
- Le chiffre de contrôle (2 derniers chiffres du code déclaration, ex. `1250-11`) se
  lit uniquement dans le PDF — l'Excel ne le contient pas.

## Logique métier essentielle : PDF prioritaire sur Excel
**Les libellés Excel peuvent être obsolètes** (constaté : "Marié en 2007" dans l'Excel
alors que le PDF 2025 affiche "mariés en 2024..."). Règle : pour tout code présent
dans un PDF, la hiérarchie (Cadre/Section/Rubrique) et le libellé du PDF priment sur
l'Excel. L'Excel ne sert de source que pour les codes qui n'apparaissent dans aucun PDF
(variables calculées/administratives, ~53% des codes en 2024).

Structure PDF reconstruite par regex sur le texte OCR (`02_parse_pdf_structure.py`) :
`Cadre <chiffre romain> - ...` → `<Lettre>. SECTION EN MAJUSCULES` → `<n>. Rubrique...`
→ libellé (généralement juste avant le code sur la même ligne pour les montants ;
avant les cases à cocher `□` pour les indicateurs). Attention aux rubriques qui
s'étendent sur 2 lignes (ex. "...au lieu / de travail") — gérées par une heuristique
de continuation (ligne suivante commençant par une minuscule, sans marqueur a)/b)/•).

## Granularité de sortie : une ligne par code, avec lien conjoint
Décision utilisateur (importante, ne pas revenir dessus sans consigne explicite) :
**une ligne par code IPCAL** (nécessaire pour les traitements sur données brutes en
aval), avec :
- `Conjoint` (1 ou 2) pour filtrer directement le conjoint 1 (usage principal après
  individualisation des données).
- `Code_IPCAL_conjoint` (+ `Code_declaration_conjoint*`) pour retrouver le code apparié
  sans dupliquer la ligne.
- `A_pendant_conjoint` (booléen) : la variable existe-t-elle réellement pour les deux
  conjoints, ou est-elle individuelle/commune (ex. totaux ménage) ?

## Disponibilité : colonnes booléennes (pas de texte à parser)
`Dispo_PDF`, `Dispo_PDF_resident`, `Dispo_PDF_INR`, `Dispo_Excel`, `Present_IPP`,
`Present_INR`, `Region_dependante`, `Nouveau_cette_annee`, plus une colonne
`Present_<clé_document>` par document source (ex. `Present_P1_BXL`, `Present_P2`,
`Present_INR_P1`...) générée dynamiquement selon les clés présentes dans le manifest.

## IPType (colonne Excel, régimes fiscaux)
90/91 = revenus immobiliers (à préciser — pas encore clarifié avec l'utilisateur) ;
92=Pays-Bas, 93=Allemagne, 94=Luxembourg (revenus étrangers exonérés, réserve de
progression) ; 95=CSSS exonéré ; 96=épargne française ; 97=autres pays exonérés IPP.

## Dérive sémantique multi-année (instrumentée — `scripts/05_build_historique.py` + `06_write_historique_excel.py`)
Une variable peut changer de sens au fil du temps sous le même code. Signal empirique
observé par l'utilisateur : une **rupture de disponibilité** d'un code pendant ≥1 an,
suivie d'une réapparition, indique souvent un changement de sémantique. Confirmé sur
2017-2023 : sur 6982 codes multi-années, 76 ont un trou de disponibilité, et les 76
correspondent à la relecture à un changement de libellé substantiel (similarité
difflib max = 0.559 après normalisation — voir seuil `SEUIL_SUSPECT` dans le script).
- `05_build_historique.py` : lit tous les `records.json`, regroupe par `Code_IPCAL`
  (vérifié : pas de renumérotation d'une année à l'autre dans ce jeu de données —
  `Code_IPCAL_precedent == Code_IPCAL` dans 100% des cas non-nouveaux — donc le code
  sert d'identifiant stable de variable ; à revoir si une future année renumérote),
  détecte les runs d'années manquantes dans le span [première, dernière année connue],
  calcule une similarité libellé+hiérarchie avant/après chaque coupure, et **scinde
  en plusieurs "générations" les codes dont une coupure est jugée suspecte** (une
  ligne par génération, plutôt qu'une ligne par code brut — voir ci-dessous).
  `python3 05_build_historique.py data/*/records.json -o data/historique.json`
- `06_write_historique_excel.py` : classeur 4 feuilles — `Codes_avec_rupture` (LA
  feuille à consulter avant tout mapping : uniquement les ~76 codes scindés, une
  ligne par génération avec sa plage de validité), `Historique` (une ligne par
  génération, libellé par année en colonnes, cellules grises hors plage de validité
  de la génération, oranges pour une coupure mineure interne), `Ruptures_semantiques`
  (une ligne par frontière entre générations, triée par similarité croissante =
  cas les plus suspects en premier), `Legende` (méthode et limites).
  `python3 06_write_historique_excel.py data/historique.json -o IPCAL_historique_variables.xlsx`
- **Pas d'identifiant inventé** (décision utilisateur, 2026-09-16, à ne pas revenir
  dessus sans consigne explicite) : `Code_IPCAL` reste la clé partout, y compris dans
  le `Dictionnaire` annuel, où `(Code_IPCAL, Annee_revenus)` désambiguïse déjà
  totalement. Seul un usage hors-année (mapping vers des agrégats nationaux) a besoin
  de désambiguïser un code scindé — d'où `Cle_mapping` : `Code_IPCAL` nu si une seule
  génération (~99% des codes, coût nul), sinon `"<Code_IPCAL>-<Validite_debut>"` (ex.
  `A0270-2021`). Le flag `Rupture_semantique` (True sur toutes les générations d'un
  code scindé) doit être vérifié avant tout mapping par `Code_IPCAL` seul.
- Limite connue : ne détecte que la dérive accompagnée d'une coupure de disponibilité
  (conforme au signal empirique documenté) ; un changement de sens sans coupure
  (code présent en continu mais qui change de signification) n'est pas couvert —
  un tel cas resterait une seule génération, avec `Rupture_semantique=False`, sans
  aucun signal d'alerte pour un futur mapping.

## Gotchas techniques
- Caractères de contrôle OCR corrompent l'écriture openpyxl : sanitizer
  `re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', v)` sur toute chaîne avant écriture
  (déjà fait dans `04_write_excel.py`).
- Les vrais booléens Python écrits directement dans une cellule openpyxl sont
  préservés comme booléens Excel — pas de contournement nécessaire.
- `zipfile.ZipFile(path)` lève `BadZipFile` si le fichier est un vrai PDF (pas une
  archive) → c'est le signal utilisé pour basculer sur le mode de secours `pdftotext`.

## Pipeline (scripts/, dans l'ordre)
0. **`00_check_sources.py`** — contrôle de couverture des sources brutes par année
   (PDF par région/partie, non-résidents, Excel maître) : présent+déclaré, manquant,
   déclaré-mais-fichier-absent, présent-mais-non-déclaré, ou non applicable (ex. pas
   de découpage régional avant 2014). Ne dépend d'aucune autre étape — tourne dès
   qu'une année est déposée, avant même 01. `python3 00_check_sources.py [--debut
   2014] [--fin <dernière année trouvée>] -o IPCAL_controle_sources.xlsx`
1. **`01_extract_pdf_text.py`** — allège les "PDF" (zip→texte, ou vrai PDF→pdftotext).
   `python3 01_extract_pdf_text.py --dir data/<année>/raw`
2. **`02_parse_pdf_structure.py`** — reconstruit Cadre/Section/Rubrique/Libellé par
   code, à partir d'un `manifest.json` par année (voir format dans le docstring du
   script et exemple ci-dessous).
   `python3 02_parse_pdf_structure.py data/<année>/manifest.json -o data/<année>/pdf_struct.json`
3. **`03_build_dictionary.py`** — fusionne Excel maître + structure PDF, produit les
   enregistrements (une ligne par code) pour une année.
   `python3 03_build_dictionary.py --excel data/<année>/IPCAL_<année>.xlsx --sheet IPCAL_Codes --struct data/<année>/pdf_struct.json --annee-revenus <année> --exercice <année+1> -o data/<année>/records.json`
4. **`04_write_excel.py`** — écrit le classeur final (une année ou plusieurs via glob).
   `python3 04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_2014_2024.xlsx`
5. **`05_build_historique.py`** — (multi-année uniquement) vue historique par
   variable + détection de dérive sémantique. `python3 05_build_historique.py
   data/*/records.json -o data/historique.json`
6. **`06_write_historique_excel.py`** — classeur Historique/Alertes_rupture/Legende.
   `python3 06_write_historique_excel.py data/historique.json -o IPCAL_historique_variables.xlsx`

## Convention de dossiers pour les nouvelles années
```
data/
  2024/
    raw/                              <- PDF + Excel bruts, tels que reçus, renommés :
      IPCAL_2024.xlsx                    Excel maître
      P1_BXL.pdf / P1_RF.pdf / P1_RW.pdf Partie 1, résidents, par région
      P2.pdf                              Partie 2, indépendants
      INR_P1.pdf / INR_P2.pdf            Non-résidents, parties 1/2
      P1.pdf                              Avant 2014 : PDF unique (pas de découpage régional)
    manifest.json                     <- mapping clé -> fichier .txt/.pdf, voir ex. ci-dessous
    pdf_struct.json                   <- sortie script 02
    records.json                      <- sortie script 03
  2023/
    ...
```
Règles de nommage dans `raw/` : toujours `<clé>.pdf` où `<clé>` est exactement la clé
utilisée dans `manifest.json` (pas de préfixe parasite type `111-`, une seule extension
`.pdf`) ; Excel maître toujours `IPCAL_<année_revenus>.xlsx` (jamais l'exercice, pour
éviter l'ambiguïté du style "2023-2024" — l'exercice se déduit, `année_revenus + 1`).

Exemple de `manifest.json` (bloc `excel_columns` optionnel — voir section suivante) :
```json
{
  "annee_revenus": 2024,
  "exercice_imposition": 2025,
  "documents": {
    "P1_BXL": "raw/P1_BXL.txt",
    "P1_RF": "raw/P1_RF.txt",
    "P1_RW": "raw/P1_RW.txt",
    "P2": "raw/P2.txt",
    "INR_P1": "raw/INR_P1.txt",
    "INR_P2": "raw/INR_P2.txt"
  }
}
```
Pour les années avant 2014 (un seul PDF, pas de découpage régional), adapter le
manifest avec une seule clé (ex. `"P1": "raw/P1.txt"`) — `03_build_dictionary.py`
lit désormais la liste et l'ordre des clés depuis `manifest.json` lui-même (au lieu
d'une liste câblée en dur), donc aucun changement de code n'est nécessaire pour ce cas.

## Paramétrage pour un run local (`config.json` + `manifest.json`)
Deux niveaux de configuration, tous deux optionnels (défauts = disposition ci-dessus) :
- **`config.json`** (racine du dépôt) : emplacement et disposition des dossiers/fichiers
  (gabarits avec `{annee}`/`{data_dir}`, résolus par `scripts/_layout.py`). À modifier
  si la disposition locale diffère (pas de sous-dossier `raw/`, données hors du dépôt,
  autre nom de fichier pour l'Excel maître, etc.) — voir le commentaire en tête du
  module pour le détail. Les scripts 01-03 acceptent `--annee <année>` (+ `--config
  <chemin>` en option) pour résoudre leurs chemins via cette config ; toute option de
  chemin explicite (`--excel`, `--struct`, `-o`, le manifest positionnel...) garde
  priorité sur ce que `--annee` déduirait.
- **`manifest.json` par année**, bloc `excel_columns` optionnel : feuille, nombre de
  lignes d'en-tête, indices de colonnes (voir docstring de `03_build_dictionary.py`
  pour le schéma complet), et labels de source PDF. À utiliser quand une année s'écarte
  des hypothèses par défaut (ex. 2021 : feuille nommée `Feuil1` au lieu de
  `IPCAL_Codes` — voir `data/2021/manifest.json`) plutôt que de passer des `--sheet`/
  `--header-rows` en ligne de commande, pour que l'ajustement reste versionné et
  reproductible sans avoir à s'en souvenir d'une session à l'autre.

## Points ouverts / à vérifier avec l'utilisateur avant de généraliser massivement
- Confirmer que les archives des années antérieures ont la même structure interne
  (zip page-par-page) ou si certaines sont de vrais PDF scannés sans couche texte
  (auquel cas `pdftotext` ne suffira pas — il faudra de l'OCR, cf. mode de secours
  non couvert par le script actuel dans ce cas extrême).
- Signification exacte des régimes IPType 90/91 (non clarifiée).
- Règle de région-dépendance pour les variables calculées : actuellement basée sur un
  mot-clé (gewest/régional/...) dans la hiérarchie Excel — imprécise, à raffiner si
  possible avec une source plus fiable.
- Type/Unité (`Type_infere`/`Unite_inferee`) sont des heuristiques sur les libellés,
  pas une vérité issue d'une source structurée — à valider ou remplacer si une
  meilleure source de typage existe.

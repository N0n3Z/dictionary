#!/usr/bin/env python3
"""
Contrôle de couverture des sources brutes, par année : pour chaque document attendu
(PDF partie 1 par région, partie 2, non-résidents, Excel maître), indique s'il est
présent sur disque ET déclaré dans manifest.json, absent, ou non applicable cette
année-là (ex. pas de découpage régional avant 2014).

But : repérer d'un coup d'œil un trou de collecte (ex. "il manque la partie
Bruxelles en 2016") AVANT de lancer le pipeline dessus — contrairement aux colonnes
Dispo_PDF_* du data dictionary, qui ne disent rien d'un document jamais reçu (elles
répondent "ce code est-il dans les PDF qu'on a", pas "a-t-on tous les PDF qu'on
devrait avoir").

Ne dépend que de config.json + manifest.json + l'arborescence data/<année>/raw/ —
tourne indépendamment du reste du pipeline (pas besoin d'avoir déjà lancé 01-04).

États par cellule :
  OK              fichier présent sur disque ET déclaré dans manifest.json
  MANQUANT        attendu cette année, ni fichier ni déclaration -> trou de collecte
  FICHIER_ABSENT  déclaré dans manifest.json mais le fichier référencé n'existe pas
  NON_DECLARE     fichier présent dans raw/ mais absent de manifest.json (oubli)
  N/A             non applicable cette année (ex. clés régionales avant 2014)
  ANNEE_ABSENTE   data/<année>/ n'existe pas du tout

Usage :
    python3 00_check_sources.py                       # 2014..dernière année trouvée
    python3 00_check_sources.py --debut 2014 --fin 2024
    python3 00_check_sources.py -o IPCAL_controle_sources.xlsx
"""
import json, os, re, argparse, glob
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from _layout import load_config, paths_for_year

PRE2014_KEYS = ['P1']
POST2014_CORE = ['P1_BXL', 'P1_RF', 'P1_RW', 'P2']
POST2014_SUPP = ['INR_P1', 'INR_P2']
ALL_DOC_KEYS = PRE2014_KEYS + POST2014_CORE + POST2014_SUPP  # ordre d'affichage
PSEUDO_EXCEL_KEY = 'Excel_maitre'

LABELS = {
    'P1': 'Partie 1 (unique, avant découpage régional)',
    'P1_BXL': 'Partie 1 — Bruxelles', 'P1_RF': 'Partie 1 — Flandre', 'P1_RW': 'Partie 1 — Wallonie',
    'P2': 'Partie 2 (indépendants)',
    'INR_P1': 'Non-résidents, partie 1', 'INR_P2': 'Non-résidents, partie 2',
    PSEUDO_EXCEL_KEY: 'Excel maître',
}


def applicable_keys(annee):
    return set(PRE2014_KEYS) if annee < 2014 else set(POST2014_CORE) | set(POST2014_SUPP)


def is_core(key):
    return key in POST2014_CORE or key in PRE2014_KEYS


def check_year(annee, config_path=None):
    paths = paths_for_year(annee, config_path)
    year_dir = paths['year_dir']
    result = {'annee': annee, 'annee_absente': not os.path.isdir(year_dir), 'cellules': {}}
    if result['annee_absente']:
        for k in ALL_DOC_KEYS + [PSEUDO_EXCEL_KEY]:
            result['cellules'][k] = 'ANNEE_ABSENTE'
        return result

    manifest = {}
    if os.path.exists(paths['manifest']):
        manifest = json.load(open(paths['manifest'], encoding='utf-8'))
    declared = manifest.get('documents', {})
    raw_dir = paths['raw_dir']
    applicable = applicable_keys(annee)

    for key in ALL_DOC_KEYS:
        if key not in applicable:
            result['cellules'][key] = 'N/A'
            continue
        on_disk = any(os.path.exists(os.path.join(raw_dir, key + ext)) for ext in ('.pdf', '.txt'))
        is_declared = key in declared
        decl_file_exists = is_declared and os.path.exists(os.path.join(year_dir, declared[key]))
        if is_declared and (decl_file_exists or on_disk):
            result['cellules'][key] = 'OK'
        elif is_declared and not decl_file_exists:
            result['cellules'][key] = 'FICHIER_ABSENT'
        elif on_disk and not is_declared:
            result['cellules'][key] = 'NON_DECLARE'
        else:
            result['cellules'][key] = 'MANQUANT'

    excel_path = paths.get('excel_master')
    excel_on_disk = bool(excel_path and os.path.exists(excel_path))
    if not excel_on_disk and os.path.isdir(raw_dir):
        excel_on_disk = bool(glob.glob(os.path.join(raw_dir, '*.xlsx')))
    result['cellules'][PSEUDO_EXCEL_KEY] = 'OK' if excel_on_disk else 'MANQUANT'
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--debut', type=int, default=2014, help="première année revenus à contrôler (défaut 2014, objectif documenté dans CLAUDE.md)")
    ap.add_argument('--fin', type=int, help="dernière année revenus à contrôler (défaut : dernière année trouvée dans data/, ou --debut si aucune)")
    ap.add_argument('--config', help="chemin vers config.json (défaut : config.json à la racine du dépôt)")
    ap.add_argument('-o', '--out', default='IPCAL_controle_sources.xlsx')
    args = ap.parse_args()

    cfg = load_config(args.config)
    data_dir = cfg['data_dir']
    found_years = sorted(int(os.path.basename(p)) for p in glob.glob(os.path.join(data_dir, '*'))
                          if os.path.isdir(p) and re.fullmatch(r'\d{4}', os.path.basename(p)))
    fin = args.fin if args.fin is not None else (max(found_years) if found_years else args.debut)

    rows = [check_year(y, args.config) for y in range(args.debut, fin + 1)]

    # ---- Rapport console ----
    problems = []
    for r in rows:
        if r['annee_absente']:
            problems.append(f"  {r['annee']} : dossier data/{r['annee']}/ absent")
            continue
        manquants = [LABELS[k] for k, v in r['cellules'].items()
                     if v == 'MANQUANT' and (is_core(k) or k == PSEUDO_EXCEL_KEY)]
        fichiers_absents = [LABELS[k] for k, v in r['cellules'].items() if v == 'FICHIER_ABSENT']
        non_declares = [LABELS[k] for k, v in r['cellules'].items() if v == 'NON_DECLARE']
        supp_manquants = [LABELS[k] for k, v in r['cellules'].items() if v == 'MANQUANT' and k in POST2014_SUPP]
        if manquants:
            problems.append(f"  {r['annee']} : MANQUANT (coeur) -> {', '.join(manquants)}")
        if fichiers_absents:
            problems.append(f"  {r['annee']} : déclaré mais fichier absent -> {', '.join(fichiers_absents)}")
        if non_declares:
            problems.append(f"  {r['annee']} : fichier présent mais non déclaré dans manifest.json -> {', '.join(non_declares)}")
        if supp_manquants:
            problems.append(f"  {r['annee']} : absent (complémentaire, non-résidents) -> {', '.join(supp_manquants)}")

    print(f'Contrôle des sources, {args.debut}-{fin} ({len(rows)} années) :')
    if problems:
        print(f'{len(problems)} anomalie(s) :')
        for p in problems:
            print(p)
    else:
        print('  Aucune anomalie : toutes les sources coeur attendues sont présentes et déclarées.')

    write_excel(rows, args.out)
    print(f'Écrit : {args.out}')


def write_excel(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Controle_sources'

    HEAD = PatternFill('solid', fgColor='1F4E78')
    HEADF = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    CELLF = Font(name='Arial', size=9, bold=True)
    thin = Side(style='thin', color='D9D9D9')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    FILL = {
        'OK': PatternFill('solid', fgColor='C6E0B4'),
        'MANQUANT': PatternFill('solid', fgColor='F8696B'),
        'FICHIER_ABSENT': PatternFill('solid', fgColor='F8696B'),
        'NON_DECLARE': PatternFill('solid', fgColor='FFD966'),
        'N/A': PatternFill('solid', fgColor='E7E6E6'),
        'ANNEE_ABSENTE': PatternFill('solid', fgColor='808080'),
    }
    TEXT = {'OK': 'OK', 'MANQUANT': 'MANQUANT', 'FICHIER_ABSENT': 'FICHIER ABSENT',
            'NON_DECLARE': 'NON DÉCLARÉ', 'N/A': '—', 'ANNEE_ABSENTE': 'ANNÉE ABSENTE'}

    cols = ['Annee_revenus'] + ALL_DOC_KEYS + [PSEUDO_EXCEL_KEY]
    header_labels = ['Année revenus'] + [LABELS[k] for k in ALL_DOC_KEYS] + [LABELS[PSEUDO_EXCEL_KEY]]
    ws.append(header_labels)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    for i, r in enumerate(rows):
        rr = i + 2
        cell = ws.cell(row=rr, column=1, value=r['annee'])
        cell.font = Font(name='Arial', bold=True, size=9); cell.border = BORDER
        cell.alignment = Alignment(horizontal='center')
        for c, key in enumerate(ALL_DOC_KEYS + [PSEUDO_EXCEL_KEY], start=2):
            state = r['cellules'][key]
            cell = ws.cell(row=rr, column=c, value=TEXT[state])
            cell.font = CELLF; cell.border = BORDER; cell.fill = FILL[state]
            cell.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(rows) + 1}"
    ws.column_dimensions['A'].width = 13
    for i in range(2, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
    ws.row_dimensions[1].height = 46

    lg = wb.create_sheet('Legende')
    lg.column_dimensions['A'].width = 20; lg.column_dimensions['B'].width = 100
    bf = Font(name='Arial', bold=True, size=10); nf = Font(name='Arial', size=10)
    tf = Font(name='Arial', bold=True, size=13, color='1F4E78')
    r = 1
    lg.cell(row=r, column=1, value='Contrôle des sources brutes — IPCAL').font = tf; r += 2
    entries = [
        ('OK', "Fichier présent sur disque et déclaré dans manifest.json."),
        ('MANQUANT', "Attendu cette année (coeur : régions/partie 2/Excel), ni fichier ni déclaration -- trou de collecte à combler."),
        ('FICHIER ABSENT', "Déclaré dans manifest.json mais le fichier référencé n'existe pas -- chemin cassé ou fichier jamais déposé."),
        ('NON DÉCLARÉ', "Le fichier existe dans raw/ mais manifest.json ne le mentionne pas -- il sera ignoré par le pipeline tant que le manifest n'est pas mis à jour."),
        ('—  (N/A)', "Non applicable cette année (ex. clés régionales avant 2014, où un seul PDF existait -- voir CLAUDE.md)."),
        ('ANNÉE ABSENTE', "data/<année>/ n'existe pas du tout dans le dépôt."),
    ]
    for label, desc in entries:
        c1 = lg.cell(row=r, column=1, value=label); c1.font = bf; c1.alignment = Alignment(wrap_text=True, vertical='top')
        c2 = lg.cell(row=r, column=2, value=desc); c2.font = nf; c2.alignment = Alignment(wrap_text=True, vertical='top')
        r += 1
    r += 1
    lg.cell(row=r, column=1, value='Non-résidents (INR)').font = bf; r += 1
    lg.cell(row=r, column=2, value="Classées \"complémentaires\", pas \"coeur\" : on ne sait pas avec certitude si elles existent pour chaque année (seule 2023 en a été fournie à ce jour). Un MANQUANT sur ces colonnes est donc informatif, pas nécessairement une anomalie de collecte.").font = nf
    lg.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical='top')

    wb.save(out_path)


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""
Écrit le classeur Excel final (Dictionnaire / Legende / Diagnostics) à partir d'un
ou plusieurs fichiers records.json produits par 03_build_dictionary.py.

Usage (une année) :
    python3 04_write_excel.py data/2024/records.json -o IPCAL_data_dictionary_2024.xlsx

Usage (série multi-années, un fichier records.json par année) :
    python3 04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_2014_2024.xlsx
"""
import json, argparse, glob, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re as _re

ILLEGAL = _re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

def san(v):
    return ILLEGAL.sub('', v) if isinstance(v, str) else v

COLS = [
    'Annee_revenus', 'Exercice_imposition', 'Code_IPCAL', 'Prefixe', 'Conjoint', 'Portee',
    'A_pendant_conjoint', 'Code_IPCAL_conjoint',
    'Code_declaration', 'Code_declaration_complet', 'Code_declaration_conjoint', 'Code_declaration_conjoint_complet',
    'Source_hierarchie', 'Cadre', 'Categorie', 'Sous_categorie', 'Libelle_FR', 'Source_libelle', 'Chemin_hierarchique',
    'Cadre_PDF', 'Section_PDF', 'Rubrique_PDF', 'Libelle_PDF', 'Libelle_PDF_complet',
    'Cadre_Excel_FR', 'Categorie_Excel_FR', 'SousCat_Excel_FR', 'Detail_Excel_FR', 'Libelle_Excel_FR', 'Libelle_Excel_NL',
    'Cadre_Excel_NL', 'Categorie_Excel_NL', 'SousCat_Excel_NL', 'Detail_Excel_NL',
    'Type_infere', 'Unite_inferee', 'Nature_variable', 'Region_dependante', 'Regions_disponibilite',
    'Dispo_PDF', 'Dispo_PDF_resident', 'Dispo_PDF_INR', 'Dispo_Excel', 'Disponibilite_source', 'Partie_declaration',
    'Present_IPP', 'Present_INR', 'Champ_IPP_INR',
    # Present_P1_BXL, Present_P1_RF, ... sont ajoutées dynamiquement (voir plus bas)
    'IPType', 'IPType_libelle',
    'Code_IPCAL_precedent', 'Nouveau_cette_annee', 'Date_debut_disponibilite', 'Rupture_semantique',
    'Ligne_source_Excel',
]
BOOL_COLS = {'A_pendant_conjoint', 'Region_dependante', 'Dispo_PDF', 'Dispo_PDF_resident', 'Dispo_PDF_INR',
             'Dispo_Excel', 'Present_IPP', 'Present_INR', 'Nouveau_cette_annee'}


def load_records(patterns):
    records = []
    for pat in patterns:
        for path in sorted(glob.glob(pat)) if any(c in pat for c in '*?[') else [pat]:
            recs = json.load(open(path, encoding='utf-8'))
            records.extend(recs)
            print(f'  {path}: {len(recs)} lignes')
    return records


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('records', nargs='+', help='un ou plusieurs records.json (globs acceptés)')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    print('Chargement des enregistrements :')
    records = load_records(args.records)
    print(f'Total : {len(records)} lignes')
    if not records:
        raise SystemExit('Aucun enregistrement chargé — vérifier les chemins.')

    # colonnes Present_<source> dynamiques (peuvent varier selon les années/sources dispo)
    present_keys = sorted({k for r in records for k in r if k.startswith('Present_')
                            and k not in ('Present_IPP', 'Present_INR')})
    cols = COLS[:47] + present_keys + COLS[47:]  # insérés après Champ_IPP_INR
    bool_cols = BOOL_COLS | set(present_keys)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Dictionnaire'
    HEAD = PatternFill('solid', fgColor='1F4E78')
    HEADF = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    CELLF = Font(name='Arial', size=9)
    thin = Side(style='thin', color='D9D9D9')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt = PatternFill('solid', fgColor='F2F6FB')

    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    # tri : année, puis code IPCAL (lisibilité de la série multi-années)
    records_sorted = sorted(records, key=lambda r: (r.get('Annee_revenus', 0), r.get('Code_IPCAL', '')))

    for i, rec in enumerate(records_sorted):
        ws.append([san(rec.get(k, '')) for k in cols])
        rr = i + 2
        for c, k in enumerate(cols, start=1):
            cell = ws.cell(row=rr, column=c); cell.font = CELLF; cell.border = BORDER
            if k in bool_cols:
                cell.alignment = Alignment(horizontal='center')
            if i % 2 == 1:
                cell.fill = alt
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(records_sorted) + 1}"

    W = {'Code_IPCAL': 11, 'Code_IPCAL_conjoint': 13, 'Conjoint': 9, 'Prefixe': 8, 'Portee': 20,
         'A_pendant_conjoint': 10, 'Cadre': 40, 'Categorie': 30, 'Sous_categorie': 38, 'Libelle_FR': 46,
         'Chemin_hierarchique': 60, 'Cadre_PDF': 40, 'Section_PDF': 30, 'Rubrique_PDF': 40,
         'Libelle_PDF': 40, 'Libelle_PDF_complet': 50, 'Nature_variable': 32, 'IPType_libelle': 44,
         'Regions_disponibilite': 24, 'Champ_IPP_INR': 22}
    for i, k in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = W.get(k, 13)
    ws.row_dimensions[1].height = 40

    # ---- Diagnostics (léger, calculé uniquement à partir des records) ----
    dg = wb.create_sheet('Diagnostics')
    from collections import Counter
    dg.column_dimensions['A'].width = 46; dg.column_dimensions['B'].width = 70
    bf = Font(name='Arial', bold=True, size=10); nf = Font(name='Arial', size=10)
    years = sorted({r['Annee_revenus'] for r in records})
    stats = [
        ('STATISTIQUES', ''),
        ('Années couvertes', ', '.join(str(y) for y in years)),
        ('Lignes-codes (total)', len(records)),
        ('Hiérarchie/libellé issus du PDF', Counter(r['Source_hierarchie'] for r in records)['PDF']),
        ('Hiérarchie/libellé issus de l\'Excel', Counter(r['Source_hierarchie'] for r in records)['Excel']),
        ('Région-dépendants', Counter(r['Region_dependante'] for r in records)[True]),
        ('Présents aussi à l\'INR', Counter(r['Champ_IPP_INR'] for r in records)['IPP + INR']),
        ('', ''),
    ]
    r = 1
    for a, b in stats:
        dg.cell(row=r, column=1, value=a).font = Font(name='Arial', bold=(b == ''), size=10)
        dg.cell(row=r, column=2, value=b).font = nf; r += 1
    dg.cell(row=r, column=1, value='Par année :').font = bf; r += 1
    dg.cell(row=r, column=1, value='Année').font = bf; dg.cell(row=r, column=2, value='Nb lignes').font = bf; r += 1
    for y in years:
        n = sum(1 for rec in records if rec['Annee_revenus'] == y)
        dg.cell(row=r, column=1, value=y).font = nf; dg.cell(row=r, column=2, value=n).font = nf; r += 1

    # ---- Legende (statique, à adapter si vous changez le schéma) ----
    lg = wb.create_sheet('Legende')
    lg.column_dimensions['A'].width = 32; lg.column_dimensions['B'].width = 112
    tf = Font(name='Arial', bold=True, size=13, color='1F4E78')
    sf = Font(name='Arial', bold=True, size=11, color='1F4E78')
    def put(row, a, b):
        ca = lg.cell(row=row, column=1, value=a); ca.font = bf; ca.alignment = Alignment(wrap_text=True, vertical='top')
        cb = lg.cell(row=row, column=2, value=b); cb.font = nf; cb.alignment = Alignment(wrap_text=True, vertical='top')
    r = 1
    lg.cell(row=r, column=1, value='Data dictionary IPCAL — IPP').font = tf; r += 2
    lg.cell(row=r, column=1, value='STRUCTURE').font = sf; r += 1
    put(r, 'Granularité', "Une ligne par code IPCAL. Conjoint 1 (préfixes A/C/E/G/I/K) et conjoint 2 (B/D/F/H/J/L) sur deux lignes reliées par Code_IPCAL_conjoint."); r += 1
    put(r, 'Source_hierarchie / Source_libelle', "PDF prioritaire quand le code y figure (libellés à jour) ; sinon repli sur l'Excel (codes calculés/administratifs)."); r += 1
    put(r, 'IPType', "90/91=revenus immobiliers (à préciser) ; 92=Pays-Bas, 93=Allemagne, 94=Luxembourg (exonérés) ; 95=CSSS exonéré ; 96=épargne française ; 97=autres pays exonérés IPP."); r += 1
    put(r, 'Voir CLAUDE.md', "Pour le détail complet des règles de construction, consulter CLAUDE.md à la racine du dépôt.")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    wb.save(args.out)
    print(f'Écrit: {args.out} | {len(cols)} colonnes, {len(records_sorted)} lignes')


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""
Écrit le classeur "vue historique par variable sémantique" à partir du JSON produit
par 05_build_historique.py : une ligne par Code_IPCAL, avec le libellé de chaque
année en colonnes (vue large, pour un scan visuel rapide de l'évolution), plus une
feuille d'alertes listant les ruptures de disponibilité suspectées de correspondre
à un changement de sens (à revoir manuellement -- voir Legende pour la méthode et
ses limites).

Usage :
    python3 06_write_historique_excel.py data/historique.json -o IPCAL_historique_variables.xlsx
"""
import json, argparse, os, re as _re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ILLEGAL = _re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

def san(v):
    return ILLEGAL.sub('', v) if isinstance(v, str) else v


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('historique', help='JSON produit par 05_build_historique.py')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    data = json.load(open(args.historique, encoding='utf-8'))
    years = data['annees']
    variables = data['variables']
    print(f'{len(variables)} variables, {len(years)} années ({years[0]}-{years[-1]})')

    wb = openpyxl.Workbook()
    HEAD = PatternFill('solid', fgColor='1F4E78')
    HEADF = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    CELLF = Font(name='Arial', size=9)
    thin = Side(style='thin', color='D9D9D9')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt = PatternFill('solid', fgColor='F2F6FB')
    gap_fill = PatternFill('solid', fgColor='FDE9D9')       # année absente dans le span
    alert_fill = PatternFill('solid', fgColor='F8CBAD')     # rupture sémantique suspectée

    # ---- Feuille 1 : Historique (vue large, une ligne par Code_IPCAL) ----
    ws = wb.active
    ws.title = 'Historique'
    fixed_cols = ['Code_IPCAL', 'Prefixe', 'Conjoint', 'Nature_variable',
                  'Premiere_annee_connue', 'Derniere_annee_connue', 'Nb_annees_presentes',
                  'Continuite_pct', 'Rupture_disponibilite', 'Nb_ruptures',
                  'Rupture_semantique_suspectee', 'Similarite_min_avant_apres_rupture']
    year_cols = [f'Libelle_{y}' for y in years]
    cols = fixed_cols + year_cols

    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    variables_sorted = sorted(variables, key=lambda v: v['Code_IPCAL'])
    for i, v in enumerate(variables_sorted):
        row_vals = [v.get(k, '') for k in fixed_cols]
        libelles = v['Libelle_par_annee']
        row_vals += [libelles.get(str(y), libelles.get(y, '')) for y in years]
        ws.append([san(x) for x in row_vals])
        rr = i + 2
        present_years = set(v['Annees_presentes'])
        span = range(v['Premiere_annee_connue'], v['Derniere_annee_connue'] + 1)
        for c, k in enumerate(cols, start=1):
            cell = ws.cell(row=rr, column=c); cell.font = CELLF; cell.border = BORDER
            if k in ('Rupture_disponibilite', 'Rupture_semantique_suspectee'):
                cell.alignment = Alignment(horizontal='center')
            elif k == 'Continuite_pct':
                cell.alignment = Alignment(horizontal='right')
            if k.startswith('Libelle_'):
                y = int(k.split('_')[1])
                if y in span and y not in present_years:
                    cell.fill = gap_fill
                elif i % 2 == 1:
                    cell.fill = alt
            elif v['Rupture_semantique_suspectee']:
                cell.fill = alert_fill
            elif i % 2 == 1:
                cell.fill = alt
    ws.freeze_panes = 'E2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(variables_sorted) + 1}"
    W = {'Code_IPCAL': 11, 'Nature_variable': 32, 'Rupture_semantique_suspectee': 14,
         'Similarite_min_avant_apres_rupture': 12, 'Rupture_disponibilite': 12}
    for i, k in enumerate(cols, start=1):
        width = W.get(k, 30 if k.startswith('Libelle_') else 13)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 40

    # ---- Feuille 2 : Alertes de rupture (une ligne par épisode de coupure) ----
    al = wb.create_sheet('Alertes_rupture')
    al_cols = ['Code_IPCAL', 'Nature_variable', 'Annees_absentes', 'Avant_annee', 'Avant_libelle',
               'Avant_chemin', 'Apres_annee', 'Apres_libelle', 'Apres_chemin', 'Similarite', 'Suspect']
    al.append(al_cols)
    for c in range(1, len(al_cols) + 1):
        cell = al.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    alert_rows = []
    for v in variables_sorted:
        for a in v['Alertes_rupture']:
            alert_rows.append([
                v['Code_IPCAL'], v['Nature_variable'], ','.join(str(y) for y in a['annees_absentes']),
                a['avant_annee'], a['avant_libelle'], a['avant_chemin'],
                a['apres_annee'], a['apres_libelle'], a['apres_chemin'],
                a['similarite'], a['suspect'],
            ])
    alert_rows.sort(key=lambda r: r[9])  # tri par similarité croissante = cas les plus suspects d'abord
    for i, row in enumerate(alert_rows):
        al.append([san(x) for x in row])
        rr = i + 2
        for c in range(1, len(al_cols) + 1):
            cell = al.cell(row=rr, column=c); cell.font = CELLF; cell.border = BORDER
            if row[10]:
                cell.fill = alert_fill
            elif i % 2 == 1:
                cell.fill = alt
    if alert_rows:
        al.freeze_panes = 'E2'
        al.auto_filter.ref = f"A1:{get_column_letter(len(al_cols))}{len(alert_rows) + 1}"
    ALW = {'Avant_libelle': 36, 'Apres_libelle': 36, 'Avant_chemin': 50, 'Apres_chemin': 50,
           'Nature_variable': 32, 'Annees_absentes': 16}
    for i, k in enumerate(al_cols, start=1):
        al.column_dimensions[get_column_letter(i)].width = ALW.get(k, 13)
    al.row_dimensions[1].height = 30

    # ---- Feuille 3 : Légende / méthode ----
    lg = wb.create_sheet('Legende')
    lg.column_dimensions['A'].width = 34; lg.column_dimensions['B'].width = 112
    tf = Font(name='Arial', bold=True, size=13, color='1F4E78')
    sf = Font(name='Arial', bold=True, size=11, color='1F4E78')
    bf = Font(name='Arial', bold=True, size=10); nf = Font(name='Arial', size=10)
    def put(row, a, b):
        ca = lg.cell(row=row, column=1, value=a); ca.font = bf; ca.alignment = Alignment(wrap_text=True, vertical='top')
        cb = lg.cell(row=row, column=2, value=b); cb.font = nf; cb.alignment = Alignment(wrap_text=True, vertical='top')
    r = 1
    lg.cell(row=r, column=1, value='Historique multi-année par variable sémantique — IPCAL').font = tf; r += 2
    lg.cell(row=r, column=1, value='PRINCIPE').font = sf; r += 1
    put(r, 'Grain', "Une ligne par Code_IPCAL. Vérifié sur 2017-2023 : les codes ne sont pas renumérotés d'une année à l'autre (Code_IPCAL_precedent == Code_IPCAL dans 100% des cas) — le code est donc traité comme l'identifiant stable de la variable. À revalider si une année future introduit une renumérotation."); r += 1
    put(r, 'Rupture_disponibilite', "Le code est absent au moins une année entre sa première et sa dernière apparition connue, puis réapparaît. Fait constaté, pas une heuristique."); r += 1
    put(r, 'Rupture_semantique_suspectee', "En plus de la rupture de disponibilité, le libellé + la hiérarchie d'avant/après la coupure sont jugés peu similaires (similarité difflib < 0.6, après normalisation : minuscules, ponctuation et années à 4 chiffres retirées — cf. scripts/05_build_historique.py, SEUIL_SUSPECT). Seuil calibré sur 2017-2023 : sur les 76 ruptures observées, la similarité max était de 0.559 et correspondait, à relecture manuelle, à un vrai changement de sens (ex. A7990 'Option sur action' -> 'Periode 2 : juin - sept') — dans ce jeu de données, TOUTE rupture de disponibilité s'accompagne d'un changement de libellé substantiel. Ce constat pourrait ne pas se généraliser à des données futures : le score sert de priorisation, pas de filtre définitif — toujours vérifier Avant_libelle/Apres_libelle dans l'onglet Alertes_rupture avant de conclure."); r += 1
    put(r, 'Continuite_pct', "Part des années du span [Premiere_annee_connue, Derniere_annee_connue] où le code est effectivement présent (100% = aucune coupure)."); r += 1
    put(r, 'Cellules oranges (Historique)', "Année dans le span mais code absent (trou de disponibilité)."); r += 1
    put(r, 'Lignes/cellules rouges', "Rupture sémantique suspectée — à vérifier en priorité dans Alertes_rupture."); r += 1
    lg.cell(row=r, column=1, value='LIMITES CONNUES').font = sf; r += 1
    put(r, 'Fenêtre observée', "Les 'première/dernière année connue' sont relatives aux données chargées (2017-2023 actuellement), pas à l'existence réelle du code avant/après cette fenêtre."); r += 1
    put(r, 'Faux positifs possibles', "Un changement de source (Excel -> PDF ou l'inverse) autour d'une coupure peut à lui seul faire baisser la similarité même si le sens n'a pas changé (styles de libellé différents)."); r += 1
    put(r, 'Faux négatifs possibles', "Un changement de sens sans rupture de disponibilité (le code reste présent en continu mais change de signification) n'est pas détecté par cette méthode — seules les coupures sont examinées, conformément au signal empirique documenté dans CLAUDE.md."); r += 1
    put(r, 'Voir aussi', "CLAUDE.md à la racine du dépôt, section 'Dérive sémantique multi-année', et scripts/05_build_historique.py pour le détail de l'algorithme.")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    wb.save(args.out)
    n_susp = sum(1 for v in variables if v['Rupture_semantique_suspectee'])
    print(f'Écrit: {args.out} | {len(variables)} variables, {len(alert_rows)} épisodes de rupture ({n_susp} suspects)')


if __name__ == '__main__':
    main()

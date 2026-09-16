#!/usr/bin/env python3
"""
Construit les enregistrements du data dictionary IPCAL pour UNE année, à partir :
  - de l'Excel maître de l'année (feuille contenant les codes IPCAL — voir --sheet)
  - de la structure PDF produite par 02_parse_pdf_structure.py (--struct)

Granularité : UNE LIGNE PAR CODE IPCAL (conjoint 1 et conjoint 2 sur deux lignes
distinctes, reliées par Code_IPCAL_conjoint). La hiérarchie et le libellé PDF sont
prioritaires sur l'Excel quand le code y figure ; sinon repli sur l'Excel.

Usage :
    python3 03_build_dictionary.py \
        --excel data/2024/IPCAL_20232024.xlsx \
        --sheet IPCAL_Codes \
        --struct data/2024/pdf_struct.json \
        --annee-revenus 2024 --exercice 2025 \
        -o data/2024/records.json

    # équivalent, en s'appuyant sur config.json + data/2024/manifest.json :
    python3 03_build_dictionary.py --annee 2024

Hypothèses par défaut, reprises du test 2024 (À VÉRIFIER/AJUSTER pour chaque
nouvelle année, voir CLAUDE.md à la racine du dépôt pour le détail) :
  - La feuille maître a un en-tête sur 3 lignes, données à partir de la ligne 4.
  - Colonnes (index 0-based) : 2=IPCAL_A(année N-1), 4=IPCAL_B(année N-1),
    5=Décl_A(année N), 6=IPCAL_A(année N), 7=Décl_B(année N), 8=IPCAL_B(année N),
    9=IPType, 14-17=hiérarchie NL (4 niveaux), 18-21=hiérarchie FR (4 niveaux).
    ADAPTER ces index si la structure du fichier Excel change d'une année à l'autre
    (imprimer les 5 premières lignes de la feuille pour vérifier avant de lancer).

Pour ne pas éditer ce fichier quand une année s'écarte de ces hypothèses, tout ce
qui précède peut être surchargé par année dans manifest.json, sous une clé
"excel_columns" optionnelle :
    {
      "annee_revenus": 2021, "exercice_imposition": 2022,
      "documents": {"P1_BXL": "raw/P1_BXL.txt", ...},
      "excel_columns": {
        "sheet": "Feuil1", "header_rows": 3,
        "ipcal_a_prev": 2, "ipcal_b_prev": 4,
        "decl_a": 5, "ipcal_a": 6, "decl_b": 7, "ipcal_b": 8, "iptype": 9,
        "hier_nl": [14, 15, 16, 17], "hier_fr": [18, 19, 20, 21],
        "src_label": {"P1_BXL": "PDF P1 Bruxelles"}
      }
    }
Seules les clés qui s'écartent des valeurs par défaut ont besoin d'être présentes.
La liste et l'ordre des clés de documents (utilisés pour la priorité PDF et les
colonnes Present_<clé>) sont également lus depuis manifest.json ("documents"),
donc un manifest pré-2014 avec une seule clé "P1" fonctionne sans autre changement.
"""
import openpyxl, json, re, argparse, os
from _layout import paths_for_year, add_annee_arg

TITULAIRE = set('ACEGIK')
CONJOINT = set('BDFHJL')
REGIO_RE = re.compile(r'gewest|régional|regional|regio|vlaams|wallon|bruxell|brussel', re.I)

DEFAULT_COLUMNS = {
    'ipcal_a_prev': 2, 'ipcal_b_prev': 4,
    'decl_a': 5, 'ipcal_a': 6, 'decl_b': 7, 'ipcal_b': 8,
    'iptype': 9,
    'hier_nl': [14, 15, 16, 17],
    'hier_fr': [18, 19, 20, 21],
}

IPTYPE_LABELS = {
    90: 'Revenus immobiliers – régime 90 (à préciser)',
    91: 'Revenus immobiliers – régime 91 (à préciser)',
    92: 'Revenus étrangers exonérés – Pays-Bas (réserve de progression)',
    93: 'Revenus étrangers exonérés – Allemagne (réserve de progression)',
    94: 'Revenus étrangers exonérés – Luxembourg (réserve de progression)',
    95: 'CSSS exonéré (cotisation spéciale sécurité sociale)',
    96: "Revenus mobiliers/épargne d'origine française",
    97: 'Exonéré IPP – autres pays (réserve de progression, hors 92/93/94/96)',
}

SRC_LABEL_DEFAULT = {
    'P1_BXL': 'PDF P1 Bruxelles', 'P1_RF': 'PDF P1 Flandre', 'P1_RW': 'PDF P1 Wallonie',
    'P2': 'PDF P2 (indépendants)', 'INR_P1': 'PDF INR P1', 'INR_P2': 'PDF INR P2',
}
# Convention de classement résident/INR par préfixe de clé ("INR_" -> non-résident) ;
# la liste et l'ordre des clés eux-mêmes viennent de manifest.json (voir build()).


def cl(v):
    return str(v).strip() if v is not None else ''


def dnum(v):
    m = re.match(r'^(\d{4})', cl(v))
    return m.group(1) if m else ''


def infer_type_unit(fr, nl, ctx):
    txt = f"{fr} {nl}".lower()
    c = ctx.lower()
    if re.search(r'\bnombre\b|\baantal\b', txt):
        return 'Entier', 'nombre'
    if re.search(r'%|pourcent|percent|quotit', txt):
        return 'Décimal', 'pourcentage'
    if re.search(r'\bdate\b|datum', txt):
        return 'Date', 'date'
    if re.search(r'r[ée]gime|taxatie|aanslagvoet|aanspreektitel|civilit|titre|code ', txt):
        return 'Catégoriel', 'code'
    ind = ['célibataire', 'marié', 'veuf', 'veuve', 'cohabitant', 'handicap', 'décès', 'overlijden',
           'gehuwd', 'ongehuwd', 'weduw', 'feitelijk', 'internationa', 'gescheiden', 'exonér', 'vrijgesteld']
    if 'personalia' in c and any(k in txt for k in ind):
        return 'Indicateur (0/1)', 'booléen'
    return 'Décimal', 'EUR (présumé)'


def nature_of(prefix, declared):
    role = 'titulaire' if prefix in TITULAIRE else 'conjoint'
    if prefix in ('A', 'B'):
        fam = 'Fédérale – déclarée' if declared else 'Fédérale – calculée / administrative'
    elif prefix in ('C', 'D'):
        fam = 'Régionale – déclarée' if declared else 'Régionale – calculée / administrative'
    else:
        fam = 'Calculée / administrative'
    return f'{fam} ({role})'


def build(excel_path, sheet_name, struct_path, annee_revenus, exercice, header_rows=3,
          columns=None, doc_keys=None, src_label_overrides=None):
    columns = {**DEFAULT_COLUMNS, **(columns or {})}

    pdf_struct = json.load(open(struct_path, encoding='utf-8'))
    # Ordre/liste des clés de documents : depuis manifest.json (doc_keys) si fourni,
    # sinon depuis les clés effectivement présentes dans pdf_struct.json (compatible
    # avec un appel sans manifest, ex. usage direct de la fonction build()).
    pdf_order = list(doc_keys) if doc_keys is not None else list(pdf_struct.keys())
    for k in pdf_order:
        pdf_struct.setdefault(k, {})
    resident_keys = [k for k in pdf_order if not k.startswith('INR_')]
    inr_keys = [k for k in pdf_order if k.startswith('INR_')]
    src_label = {k: SRC_LABEL_DEFAULT.get(k, k.replace('_', ' ')) for k in pdf_order}
    if src_label_overrides:
        src_label.update(src_label_overrides)

    def pdf_lookup(d4):
        if not d4:
            return None, None
        for k in pdf_order:
            if d4 in pdf_struct.get(k, {}):
                return pdf_struct[k][d4], k
        return None, None

    def present_in(d4, keys):
        return any(d4 in pdf_struct.get(k, {}) for k in keys) if d4 else False

    def regions_of(d4):
        return [k for k in resident_keys if d4 and d4 in pdf_struct.get(k, {}) and k.startswith('P1_')]

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=header_rows + 1, values_only=True))

    records = []
    for idx, r in enumerate(rows):
        ipcalA = cl(r[columns['ipcal_a']]); ipcalB = cl(r[columns['ipcal_b']])
        declA = dnum(r[columns['decl_a']]); declB = dnum(r[columns['decl_b']])
        oldA = cl(r[columns['ipcal_a_prev']]); oldB = cl(r[columns['ipcal_b_prev']])
        iptype = r[columns['iptype']]
        nl = [cl(r[i]) for i in columns['hier_nl']]
        fr = [cl(r[i]) for i in columns['hier_fr']]
        leaf_fr = next((x for x in reversed(fr) if x), '')
        leaf_nl = next((x for x in reversed(nl) if x), '')
        ipt_raw = iptype if isinstance(iptype, int) else ''
        ipt_lbl = IPTYPE_LABELS.get(iptype, f'Régime {iptype}') if isinstance(iptype, int) else ''
        nouveau = not oldA

        sides = []
        if ipcalA:
            sides.append(('1', ipcalA, declA, ipcalB, declB))
        if ipcalB:
            sides.append(('2', ipcalB, declB, ipcalA, declA))
        has_pair = bool(ipcalA and ipcalB)

        for conj, code, decl, code_pair, decl_pair in sides:
            prefix = code[0]
            st, src = pdf_lookup(decl)
            declared = present_in(decl, resident_keys) or present_in(decl, inr_keys)
            if st:
                cadre_pdf, sec_pdf, rub_pdf, lib_pdf = st['cadre'], st['section'], st['rubrique'], st['label']
                chk = st['check']
            else:
                cadre_pdf = sec_pdf = rub_pdf = lib_pdf = ''
                chk = None
            decl_full = f"{decl}-{chk}" if (decl and chk) else decl
            st2, _ = pdf_lookup(decl_pair)
            chk2 = st2['check'] if st2 else None
            decl_pair_full = f"{decl_pair}-{chk2}" if (decl_pair and chk2) else decl_pair

            presence = {k: (decl in pdf_struct.get(k, {})) for k in pdf_order}
            in_resident = present_in(decl, resident_keys)
            in_inr = present_in(decl, inr_keys)
            dispo_pdf = in_resident or in_inr

            hier_txt = ' '.join(fr + nl)
            region_dep = prefix in ('C', 'D') or (decl[:1] in ('3', '4') if decl else False) \
                or (not decl and bool(REGIO_RE.search(hier_txt)))
            regs = regions_of(decl)
            if not regs and presence.get('P2'):
                regions_disp = 'Fédéral (P2, sans distinction régionale)'
            elif regs and len(regs) == len(resident_keys) and not region_dep:
                regions_disp = ','.join(r.replace('P1_', '') for r in regs) + ' (identique – fédéral)'
            else:
                regions_disp = ','.join(r.replace('P1_', '') for r in regs)

            parties = []
            if any(presence.get(k) for k in resident_keys if k.startswith('P1_')):
                parties.append('1')
            if presence.get('P2'):
                parties.append('2')
            partie = ' et '.join(parties)
            if not partie and in_inr:
                pi = [n for n, k in [('1', 'INR_P1'), ('2', 'INR_P2')] if presence.get(k)]
                partie = (' et '.join(pi) + ' (INR)') if pi else ''

            disponibilite = 'PDF + Excel' if in_resident else ('PDF (INR) + Excel' if in_inr else 'Excel uniquement')
            if in_inr:
                present_inr, champ = True, 'IPP + INR'
            elif decl:
                present_inr, champ = False, 'IPP uniquement'
            else:
                present_inr, champ = False, 'IPP (INR indéterminé – code calculé)'

            if st:
                source_hier = 'PDF'
                cadre, categorie, sous_cat = cadre_pdf, sec_pdf, rub_pdf
                lib_fr = lib_pdf or leaf_fr
                source_lib = src_label.get(src, 'PDF')
                lib_pdf_complet = (rub_pdf + ' > ' + lib_pdf) if (rub_pdf and lib_pdf and rub_pdf != lib_pdf) else (lib_pdf or rub_pdf)
            else:
                source_hier = 'Excel'
                cadre, categorie, sous_cat = fr[0], fr[1], fr[2]
                lib_fr, source_lib, lib_pdf_complet = leaf_fr, 'Excel', ''
            chemin_resolu = ' > '.join([x for x in [cadre, categorie, sous_cat,
                                                       (lib_fr if lib_fr != sous_cat else '')] if x])

            typ, unit = infer_type_unit(lib_fr or leaf_fr, leaf_nl, ' '.join(fr[:2] + nl[:2]))

            rec = {
                'Annee_revenus': annee_revenus, 'Exercice_imposition': exercice,
                'Code_IPCAL': code, 'Prefixe': prefix, 'Conjoint': conj,
                'Portee': 'Par conjoint' if has_pair else 'Individuelle / commune (pas de code conjoint)',
                'A_pendant_conjoint': has_pair, 'Code_IPCAL_conjoint': code_pair,
                'Code_declaration': decl, 'Code_declaration_complet': decl_full,
                'Code_declaration_conjoint': decl_pair, 'Code_declaration_conjoint_complet': decl_pair_full,
                'Source_hierarchie': source_hier, 'Cadre': cadre, 'Categorie': categorie,
                'Sous_categorie': sous_cat, 'Libelle_FR': lib_fr, 'Source_libelle': source_lib,
                'Chemin_hierarchique': chemin_resolu,
                'Cadre_PDF': cadre_pdf, 'Section_PDF': sec_pdf, 'Rubrique_PDF': rub_pdf,
                'Libelle_PDF': lib_pdf, 'Libelle_PDF_complet': lib_pdf_complet,
                'Cadre_Excel_FR': fr[0], 'Categorie_Excel_FR': fr[1], 'SousCat_Excel_FR': fr[2], 'Detail_Excel_FR': fr[3],
                'Libelle_Excel_FR': leaf_fr, 'Libelle_Excel_NL': leaf_nl,
                'Cadre_Excel_NL': nl[0], 'Categorie_Excel_NL': nl[1], 'SousCat_Excel_NL': nl[2], 'Detail_Excel_NL': nl[3],
                'Type_infere': typ, 'Unite_inferee': unit,
                'Nature_variable': nature_of(prefix, declared),
                'Region_dependante': region_dep, 'Regions_disponibilite': regions_disp,
                'Dispo_PDF': dispo_pdf, 'Dispo_PDF_resident': in_resident, 'Dispo_PDF_INR': in_inr,
                'Dispo_Excel': True, 'Disponibilite_source': disponibilite, 'Partie_declaration': partie,
                'Present_IPP': True, 'Present_INR': present_inr, 'Champ_IPP_INR': champ,
                'IPType': ipt_raw, 'IPType_libelle': ipt_lbl,
                'Code_IPCAL_precedent': (oldA if conj == '1' else oldB),
                'Nouveau_cette_annee': nouveau,
                'Date_debut_disponibilite': annee_revenus if nouveau else '',
                'Rupture_semantique': '',
                'Ligne_source_Excel': idx + header_rows + 1,
            }
            for k in pdf_order:
                rec[f'Present_{k}'] = presence.get(k, False)
            records.append(rec)
    return records


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--excel', help="Excel maître (optionnel si --annee : résolu via config.json)")
    ap.add_argument('--sheet', help="feuille Excel (défaut IPCAL_Codes, ou manifest.excel_columns.sheet)")
    ap.add_argument('--struct', help="JSON produit par 02_parse_pdf_structure.py (optionnel si --annee)")
    ap.add_argument('--manifest', help="manifest.json de l'année : fournit la liste des clés de documents "
                                        "et, en option, la structure Excel (voir docstring). Résolu via --annee si omis.")
    ap.add_argument('--annee-revenus', type=int)
    ap.add_argument('--exercice', type=int)
    ap.add_argument('--header-rows', type=int, help="nb de lignes d'en-tête (défaut 3, ou manifest.excel_columns.header_rows)")
    ap.add_argument('-o', '--out', help="optionnel si --annee")
    add_annee_arg(ap)
    args = ap.parse_args()

    paths = paths_for_year(args.annee, args.config) if args.annee is not None else {}
    args.excel = args.excel or paths.get('excel_master')
    args.struct = args.struct or paths.get('pdf_struct')
    args.manifest = args.manifest or paths.get('manifest')
    args.out = args.out or paths.get('records')
    args.annee_revenus = args.annee_revenus or args.annee

    manifest = None
    if args.manifest and os.path.exists(args.manifest):
        manifest = json.load(open(args.manifest, encoding='utf-8'))

    if args.exercice is None:
        if manifest and manifest.get('exercice_imposition'):
            args.exercice = manifest['exercice_imposition']
        elif args.annee_revenus is not None:
            args.exercice = args.annee_revenus + 1

    excel_cfg = (manifest or {}).get('excel_columns', {})
    sheet = args.sheet or excel_cfg.get('sheet') or 'IPCAL_Codes'
    header_rows = args.header_rows if args.header_rows is not None else excel_cfg.get('header_rows', 3)
    columns = {k: v for k, v in excel_cfg.items() if k not in ('sheet', 'header_rows', 'src_label')}
    doc_keys = list(manifest['documents'].keys()) if manifest else None
    src_label_overrides = excel_cfg.get('src_label')

    missing = [n for n, v in [('--excel', args.excel), ('--struct', args.struct), ('-o/--out', args.out),
                               ('--annee-revenus', args.annee_revenus), ('--exercice', args.exercice)] if not v]
    if missing:
        ap.error(f"paramètres manquants ({', '.join(missing)}) — fournir --annee, ou tout spécifier explicitement.")

    records = build(args.excel, sheet, args.struct, args.annee_revenus, args.exercice, header_rows,
                     columns=columns, doc_keys=doc_keys, src_label_overrides=src_label_overrides)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    json.dump(records, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)
    from collections import Counter
    print(f'{len(records)} lignes-codes écrites dans {args.out}')
    print('  Source hiérarchie:', dict(Counter(r['Source_hierarchie'] for r in records)))
    print('  Dispo PDF:', dict(Counter(r['Dispo_PDF'] for r in records)))


if __name__ == '__main__':
    main()

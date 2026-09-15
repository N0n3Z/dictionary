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

Hypothèses reprises du test 2024 (À VÉRIFIER/AJUSTER pour chaque nouvelle année,
voir CLAUDE.md à la racine du dépôt pour le détail) :
  - La feuille maître a un en-tête sur 3 lignes, données à partir de la ligne 4.
  - Colonnes (index 0-based) : 2=IPCAL_A(année N-1), 4=IPCAL_B(année N-1),
    5=Décl_A(année N), 6=IPCAL_A(année N), 7=Décl_B(année N), 8=IPCAL_B(année N),
    9=IPType, 14-17=hiérarchie NL (4 niveaux), 18-21=hiérarchie FR (4 niveaux).
    ADAPTER ces index si la structure du fichier Excel change d'une année à l'autre
    (imprimer les 5 premières lignes de la feuille pour vérifier avant de lancer).
"""
import openpyxl, json, re, argparse

TITULAIRE = set('ACEGIK')
CONJOINT = set('BDFHJL')
REGIO_RE = re.compile(r'gewest|régional|regional|regio|vlaams|wallon|bruxell|brussel', re.I)

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
RESIDENT_KEYS_DEFAULT = ['P1_BXL', 'P1_RF', 'P1_RW', 'P2']
INR_KEYS_DEFAULT = ['INR_P1', 'INR_P2']


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
          resident_keys=None, inr_keys=None, src_label=None):
    resident_keys = resident_keys or RESIDENT_KEYS_DEFAULT
    inr_keys = inr_keys or INR_KEYS_DEFAULT
    src_label = src_label or SRC_LABEL_DEFAULT
    pdf_order = list(src_label.keys())

    pdf_struct = json.load(open(struct_path, encoding='utf-8'))
    for k in pdf_order:
        pdf_struct.setdefault(k, {})

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
        ipcalA = cl(r[6]); ipcalB = cl(r[8])
        declA = dnum(r[5]); declB = dnum(r[7])
        oldA = cl(r[2]); oldB = cl(r[4])
        iptype = r[9]
        nl = [cl(r[14]), cl(r[15]), cl(r[16]), cl(r[17])]
        fr = [cl(r[18]), cl(r[19]), cl(r[20]), cl(r[21])]
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
    ap.add_argument('--excel', required=True)
    ap.add_argument('--sheet', default='IPCAL_Codes')
    ap.add_argument('--struct', required=True, help='JSON produit par 02_parse_pdf_structure.py')
    ap.add_argument('--annee-revenus', type=int, required=True)
    ap.add_argument('--exercice', type=int, required=True)
    ap.add_argument('--header-rows', type=int, default=3, help='nb de lignes d\'en-tête avant les données (def=3)')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    records = build(args.excel, args.sheet, args.struct, args.annee_revenus, args.exercice, args.header_rows)
    json.dump(records, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)
    from collections import Counter
    print(f'{len(records)} lignes-codes écrites dans {args.out}')
    print('  Source hiérarchie:', dict(Counter(r['Source_hierarchie'] for r in records)))
    print('  Dispo PDF:', dict(Counter(r['Dispo_PDF'] for r in records)))


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""
Builds the IPCAL data dictionary records for ONE year, from:
  - that year's master Excel file (the sheet holding the IPCAL codes -- see --sheet)
  - the PDF structure produced by 02_parse_pdf_structure.py (--struct)

Grain: ONE ROW PER IPCAL CODE (spouse 1 and spouse 2 on two separate rows, linked by
IPCAL_code_spouse). PDF hierarchy and labels take precedence over the Excel file
whenever the code appears in a PDF; otherwise the Excel file is used as fallback.

Usage:
    python3 03_build_dictionary.py \
        --excel data/2024/IPCAL_2024.xlsx \
        --sheet IPCAL_Codes \
        --struct data/2024/pdf_struct.json \
        --income-year 2024 --assessment-year 2025 \
        -o data/2024/records.json

    # equivalent, relying on config.json + data/2024/manifest.json:
    python3 03_build_dictionary.py --year 2024

Default assumptions, carried over from the 2024 test year (VERIFY/ADJUST for every
new year -- see CLAUDE.md at the repository root for details):
  - The master sheet has a 3-row header, data starts on row 4.
  - Columns (0-based): 2=IPCAL_A(year N-1), 4=IPCAL_B(year N-1), 5=Decl_A(year N),
    6=IPCAL_A(year N), 7=Decl_B(year N), 8=IPCAL_B(year N), 9=IPType,
    14-17=Dutch hierarchy (4 levels), 18-21=French hierarchy (4 levels).
    ADJUST these indices if the Excel structure shifts between years (print the
    sheet's first 5 rows to check before running).

So this file never needs editing when a year deviates, all of the above can be
overridden per year in manifest.json, under an optional "excel_columns" key:
    {
      "income_year": 2021, "assessment_year": 2022,
      "documents": {"P1_BXL": "raw/P1_BXL.txt", ...},
      "excel_columns": {
        "sheet": "Feuil1", "header_rows": 3,
        "ipcal_a_prev": 2, "ipcal_b_prev": 4,
        "decl_a": 5, "ipcal_a": 6, "decl_b": 7, "ipcal_b": 8, "iptype": 9,
        "hier_nl": [14, 15, 16, 17], "hier_fr": [18, 19, 20, 21],
        "src_label": {"P1_BXL": "PDF P1 Brussels"}
      }
    }
Only the keys that differ from the defaults need to be present. The list and order
of document keys (used for PDF precedence and for the Present_<key> columns) is also
read from manifest.json ("documents"), so a pre-2014 manifest with a single "P1" key
works with no further change.
"""
import openpyxl, json, re, argparse, os
from _layout import paths_for_year, add_year_arg

FIRST_SPOUSE = set('ACEGIK')
SECOND_SPOUSE = set('BDFHJL')
REGIO_RE = re.compile(r'gewest|régional|regional|regio|vlaams|wallon|bruxell|brussel', re.I)

DEFAULT_COLUMNS = {
    'ipcal_a_prev': 2, 'ipcal_b_prev': 4,
    'decl_a': 5, 'ipcal_a': 6, 'decl_b': 7, 'ipcal_b': 8,
    'iptype': 9,
    'hier_nl': [14, 15, 16, 17],
    'hier_fr': [18, 19, 20, 21],
}

IPTYPE_LABELS = {
    90: 'Real-estate income - regime 90 (to be clarified)',
    91: 'Real-estate income - regime 91 (to be clarified)',
    92: 'Exempt foreign income - Netherlands (exemption with progression)',
    93: 'Exempt foreign income - Germany (exemption with progression)',
    94: 'Exempt foreign income - Luxembourg (exemption with progression)',
    95: 'Exempt CSSS (special social security contribution)',
    96: 'Investment income / savings of French origin',
    97: 'PIT-exempt - other countries (exemption with progression, excl. 92/93/94/96)',
}

SRC_LABEL_DEFAULT = {
    'P1_BXL': 'PDF P1 Brussels', 'P1_RF': 'PDF P1 Flanders', 'P1_RW': 'PDF P1 Wallonia',
    'P2': 'PDF P2 (self-employed)', 'INR_P1': 'PDF INR P1', 'INR_P2': 'PDF INR P2',
}
# Resident/non-resident classification follows the key prefix ("INR_" -> non-resident);
# the list and order of keys themselves come from manifest.json (see build()).


def cl(v):
    return str(v).strip() if v is not None else ''


def dnum(v):
    m = re.match(r'^(\d{4})', cl(v))
    return m.group(1) if m else ''


def infer_type_unit(fr, nl, ctx):
    txt = f"{fr} {nl}".lower()
    c = ctx.lower()
    if re.search(r'\bnombre\b|\baantal\b', txt):
        return 'Integer', 'count'
    if re.search(r'%|pourcent|percent|quotit', txt):
        return 'Decimal', 'percentage'
    if re.search(r'\bdate\b|datum', txt):
        return 'Date', 'date'
    if re.search(r'r[ée]gime|taxatie|aanslagvoet|aanspreektitel|civilit|titre|code ', txt):
        return 'Categorical', 'code'
    ind = ['célibataire', 'marié', 'veuf', 'veuve', 'cohabitant', 'handicap', 'décès', 'overlijden',
           'gehuwd', 'ongehuwd', 'weduw', 'feitelijk', 'internationa', 'gescheiden', 'exonér', 'vrijgesteld']
    if 'personalia' in c and any(k in txt for k in ind):
        return 'Indicator (0/1)', 'boolean'
    return 'Decimal', 'EUR (assumed)'


def nature_of(prefix, declared):
    role = 'first spouse' if prefix in FIRST_SPOUSE else 'second spouse'
    if prefix in ('A', 'B'):
        fam = 'Federal - declared' if declared else 'Federal - computed / administrative'
    elif prefix in ('C', 'D'):
        fam = 'Regional - declared' if declared else 'Regional - computed / administrative'
    else:
        fam = 'Computed / administrative'
    return f'{fam} ({role})'


def build(excel_path, sheet_name, struct_path, income_year, assessment_year, header_rows=3,
          columns=None, doc_keys=None, src_label_overrides=None):
    columns = {**DEFAULT_COLUMNS, **(columns or {})}

    pdf_struct = json.load(open(struct_path, encoding='utf-8'))
    # Document key list/order: from manifest.json (doc_keys) when provided, otherwise
    # from the keys actually present in pdf_struct.json (so build() can be called
    # directly without a manifest).
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
        ipt_lbl = IPTYPE_LABELS.get(iptype, f'Regime {iptype}') if isinstance(iptype, int) else ''
        is_new = not oldA

        sides = []
        if ipcalA:
            sides.append(('1', ipcalA, declA, ipcalB, declB))
        if ipcalB:
            sides.append(('2', ipcalB, declB, ipcalA, declA))
        has_pair = bool(ipcalA and ipcalB)

        for spouse, code, decl, code_pair, decl_pair in sides:
            prefix = code[0]
            st, src = pdf_lookup(decl)
            declared = present_in(decl, resident_keys) or present_in(decl, inr_keys)
            if st:
                frame_pdf, sec_pdf, item_pdf, label_pdf = st['frame'], st['section'], st['item'], st['label']
                chk = st['check']
            else:
                frame_pdf = sec_pdf = item_pdf = label_pdf = ''
                chk = None
            decl_full = f"{decl}-{chk}" if (decl and chk) else decl
            st2, _ = pdf_lookup(decl_pair)
            chk2 = st2['check'] if st2 else None
            decl_pair_full = f"{decl_pair}-{chk2}" if (decl_pair and chk2) else decl_pair

            presence = {k: (decl in pdf_struct.get(k, {})) for k in pdf_order}
            in_resident = present_in(decl, resident_keys)
            in_inr = present_in(decl, inr_keys)
            available_pdf = in_resident or in_inr

            hier_txt = ' '.join(fr + nl)
            region_dep = prefix in ('C', 'D') or (decl[:1] in ('3', '4') if decl else False) \
                or (not decl and bool(REGIO_RE.search(hier_txt)))
            regs = regions_of(decl)
            if not regs and presence.get('P2'):
                regions_disp = 'Federal (P2, no regional split)'
            elif regs and len(regs) == len(resident_keys) and not region_dep:
                regions_disp = ','.join(r.replace('P1_', '') for r in regs) + ' (identical - federal)'
            else:
                regions_disp = ','.join(r.replace('P1_', '') for r in regs)

            parts = []
            if any(presence.get(k) for k in resident_keys if k.startswith('P1_')):
                parts.append('1')
            if presence.get('P2'):
                parts.append('2')
            part = ' and '.join(parts)
            if not part and in_inr:
                pi = [n for n, k in [('1', 'INR_P1'), ('2', 'INR_P2')] if presence.get(k)]
                part = (' and '.join(pi) + ' (INR)') if pi else ''

            availability = 'PDF + Excel' if in_resident else ('PDF (INR) + Excel' if in_inr else 'Excel only')
            if in_inr:
                present_inr, scope = True, 'IPP + INR'
            elif decl:
                present_inr, scope = False, 'IPP only'
            else:
                present_inr, scope = False, 'IPP (INR undetermined - computed code)'

            if st:
                hierarchy_source = 'PDF'
                frame, category, subcategory = frame_pdf, sec_pdf, item_pdf
                label_fr = label_pdf or leaf_fr
                label_source = src_label.get(src, 'PDF')
                label_pdf_full = (item_pdf + ' > ' + label_pdf) if (item_pdf and label_pdf and item_pdf != label_pdf) else (label_pdf or item_pdf)
            else:
                hierarchy_source = 'Excel'
                frame, category, subcategory = fr[0], fr[1], fr[2]
                label_fr, label_source, label_pdf_full = leaf_fr, 'Excel', ''
            resolved_path = ' > '.join([x for x in [frame, category, subcategory,
                                                      (label_fr if label_fr != subcategory else '')] if x])

            typ, unit = infer_type_unit(label_fr or leaf_fr, leaf_nl, ' '.join(fr[:2] + nl[:2]))

            rec = {
                'Income_year': income_year, 'Assessment_year': assessment_year,
                'IPCAL_code': code, 'Prefix': prefix, 'Spouse': spouse,
                'Scope': 'Per spouse' if has_pair else 'Individual / shared (no spouse code)',
                'Has_spouse_counterpart': has_pair, 'IPCAL_code_spouse': code_pair,
                'Declaration_code': decl, 'Declaration_code_full': decl_full,
                'Declaration_code_spouse': decl_pair, 'Declaration_code_spouse_full': decl_pair_full,
                'Hierarchy_source': hierarchy_source, 'Frame': frame, 'Category': category,
                'Subcategory': subcategory, 'Label_FR': label_fr, 'Label_source': label_source,
                'Hierarchy_path': resolved_path,
                'Frame_PDF': frame_pdf, 'Section_PDF': sec_pdf, 'Item_PDF': item_pdf,
                'Label_PDF': label_pdf, 'Label_PDF_full': label_pdf_full,
                'Frame_Excel_FR': fr[0], 'Category_Excel_FR': fr[1], 'Subcategory_Excel_FR': fr[2], 'Detail_Excel_FR': fr[3],
                'Label_Excel_FR': leaf_fr, 'Label_Excel_NL': leaf_nl,
                'Frame_Excel_NL': nl[0], 'Category_Excel_NL': nl[1], 'Subcategory_Excel_NL': nl[2], 'Detail_Excel_NL': nl[3],
                'Inferred_type': typ, 'Inferred_unit': unit,
                'Variable_nature': nature_of(prefix, declared),
                'Region_dependent': region_dep, 'Regions_available': regions_disp,
                'Available_PDF': available_pdf, 'Available_PDF_resident': in_resident, 'Available_PDF_INR': in_inr,
                'Available_Excel': True, 'Availability_source': availability, 'Declaration_part': part,
                'Present_IPP': True, 'Present_INR': present_inr, 'Scope_IPP_INR': scope,
                'IPType': ipt_raw, 'IPType_label': ipt_lbl,
                'IPCAL_code_previous': (oldA if spouse == '1' else oldB),
                'New_this_year': is_new,
                'Availability_start_year': income_year if is_new else '',
                'Semantic_break': '',
                'Excel_source_row': idx + header_rows + 1,
            }
            for k in pdf_order:
                rec[f'Present_{k}'] = presence.get(k, False)
            records.append(rec)
    return records


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--excel', help="master Excel file (optional with --year: resolved via config.json)")
    ap.add_argument('--sheet', help="Excel sheet (default IPCAL_Codes, or manifest.excel_columns.sheet)")
    ap.add_argument('--struct', help="JSON produced by 02_parse_pdf_structure.py (optional with --year)")
    ap.add_argument('--manifest', help="the year's manifest.json: provides the document key list and, "
                                        "optionally, the Excel structure (see docstring). Resolved via --year if omitted.")
    ap.add_argument('--income-year', type=int)
    ap.add_argument('--assessment-year', type=int)
    ap.add_argument('--header-rows', type=int, help="header rows before the data (default 3, or manifest.excel_columns.header_rows)")
    ap.add_argument('-o', '--out', help="optional with --year")
    add_year_arg(ap)
    args = ap.parse_args()

    paths = paths_for_year(args.year, args.config) if args.year is not None else {}
    args.excel = args.excel or paths.get('excel_master')
    args.struct = args.struct or paths.get('pdf_struct')
    args.manifest = args.manifest or paths.get('manifest')
    args.out = args.out or paths.get('records')
    args.income_year = args.income_year or args.year

    manifest = None
    if args.manifest and os.path.exists(args.manifest):
        manifest = json.load(open(args.manifest, encoding='utf-8'))

    if args.assessment_year is None:
        if manifest and manifest.get('assessment_year'):
            args.assessment_year = manifest['assessment_year']
        elif args.income_year is not None:
            args.assessment_year = args.income_year + 1

    excel_cfg = (manifest or {}).get('excel_columns', {})
    sheet = args.sheet or excel_cfg.get('sheet') or 'IPCAL_Codes'
    header_rows = args.header_rows if args.header_rows is not None else excel_cfg.get('header_rows', 3)
    columns = {k: v for k, v in excel_cfg.items() if k not in ('sheet', 'header_rows', 'src_label')}
    doc_keys = list(manifest['documents'].keys()) if manifest else None
    src_label_overrides = excel_cfg.get('src_label')

    missing = [n for n, v in [('--excel', args.excel), ('--struct', args.struct), ('-o/--out', args.out),
                               ('--income-year', args.income_year), ('--assessment-year', args.assessment_year)] if not v]
    if missing:
        ap.error(f"missing parameters ({', '.join(missing)}) - provide --year, or specify everything explicitly.")

    records = build(args.excel, sheet, args.struct, args.income_year, args.assessment_year, header_rows,
                     columns=columns, doc_keys=doc_keys, src_label_overrides=src_label_overrides)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    json.dump(records, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)
    from collections import Counter
    print(f'{len(records)} code-rows written to {args.out}')
    print('  Hierarchy source:', dict(Counter(r['Hierarchy_source'] for r in records)))
    print('  Available in PDF:', dict(Counter(r['Available_PDF'] for r in records)))


if __name__ == '__main__':
    main()

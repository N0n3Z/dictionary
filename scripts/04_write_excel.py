#!/usr/bin/env python3
"""
Writes the final Excel workbook (Dictionary / Diagnostics / Legend) from one or more
records.json files produced by 03_build_dictionary.py.

Usage (single year):
    python3 04_write_excel.py data/2024/records.json -o IPCAL_data_dictionary_2024.xlsx

Usage (multi-year series, one records.json per year):
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
    'Income_year', 'Assessment_year', 'IPCAL_code', 'Prefix', 'Spouse', 'Scope',
    'Has_spouse_counterpart', 'IPCAL_code_spouse',
    'Declaration_code', 'Declaration_code_full', 'Declaration_code_spouse', 'Declaration_code_spouse_full',
    'Hierarchy_source', 'Frame', 'Category', 'Subcategory', 'Label_FR', 'Label_source', 'Hierarchy_path',
    'Frame_PDF', 'Section_PDF', 'Item_PDF', 'Label_PDF', 'Label_PDF_full',
    'Frame_Excel_FR', 'Category_Excel_FR', 'Subcategory_Excel_FR', 'Detail_Excel_FR', 'Label_Excel_FR', 'Label_Excel_NL',
    'Frame_Excel_NL', 'Category_Excel_NL', 'Subcategory_Excel_NL', 'Detail_Excel_NL',
    'Inferred_type', 'Inferred_unit', 'Variable_nature', 'Region_dependent', 'Regions_available',
    'Available_PDF', 'Available_PDF_resident', 'Available_PDF_INR', 'Available_Excel', 'Availability_source', 'Declaration_part',
    'Present_IPP', 'Present_INR', 'Scope_IPP_INR',
    # Present_P1_BXL, Present_P1_RF, ... are inserted dynamically (see below)
    'IPType', 'IPType_label',
    'IPCAL_code_previous', 'New_this_year', 'Availability_start_year', 'Semantic_break',
    'Excel_source_row',
]
BOOL_COLS = {'Has_spouse_counterpart', 'Region_dependent', 'Available_PDF', 'Available_PDF_resident',
             'Available_PDF_INR', 'Available_Excel', 'Present_IPP', 'Present_INR', 'New_this_year'}


def load_records(patterns):
    records = []
    for pat in patterns:
        for path in sorted(glob.glob(pat)) if any(c in pat for c in '*?[') else [pat]:
            recs = json.load(open(path, encoding='utf-8'))
            records.extend(recs)
            print(f'  {path}: {len(recs)} rows')
    return records


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('records', nargs='+', help='one or more records.json (globs accepted)')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    print('Loading records:')
    records = load_records(args.records)
    print(f'Total: {len(records)} rows')
    if not records:
        raise SystemExit('No records loaded - check the paths.')

    # Dynamic Present_<source> columns (they vary with the sources available per year)
    present_keys = sorted({k for r in records for k in r if k.startswith('Present_')
                            and k not in ('Present_IPP', 'Present_INR')})
    cols = COLS[:47] + present_keys + COLS[47:]  # inserted after Scope_IPP_INR
    bool_cols = BOOL_COLS | set(present_keys)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Dictionary'
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

    # Sort by year, then IPCAL code (readability of the multi-year series)
    records_sorted = sorted(records, key=lambda r: (r.get('Income_year', 0), r.get('IPCAL_code', '')))

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

    W = {'IPCAL_code': 11, 'IPCAL_code_spouse': 13, 'Spouse': 9, 'Prefix': 8, 'Scope': 20,
         'Has_spouse_counterpart': 10, 'Frame': 40, 'Category': 30, 'Subcategory': 38, 'Label_FR': 46,
         'Hierarchy_path': 60, 'Frame_PDF': 40, 'Section_PDF': 30, 'Item_PDF': 40,
         'Label_PDF': 40, 'Label_PDF_full': 50, 'Variable_nature': 32, 'IPType_label': 44,
         'Regions_available': 24, 'Scope_IPP_INR': 22}
    for i, k in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = W.get(k, 13)
    ws.row_dimensions[1].height = 40

    # ---- Diagnostics (lightweight, computed from the records only) ----
    dg = wb.create_sheet('Diagnostics')
    from collections import Counter
    dg.column_dimensions['A'].width = 46; dg.column_dimensions['B'].width = 70
    bf = Font(name='Arial', bold=True, size=10); nf = Font(name='Arial', size=10)
    years = sorted({r['Income_year'] for r in records})
    stats = [
        ('STATISTICS', ''),
        ('Years covered', ', '.join(str(y) for y in years)),
        ('Code-rows (total)', len(records)),
        ('Hierarchy/label taken from the PDF', Counter(r['Hierarchy_source'] for r in records)['PDF']),
        ('Hierarchy/label taken from the Excel file', Counter(r['Hierarchy_source'] for r in records)['Excel']),
        ('Region-dependent', Counter(r['Region_dependent'] for r in records)[True]),
        ('Also present in the INR scope', Counter(r['Scope_IPP_INR'] for r in records)['IPP + INR']),
        ('', ''),
    ]
    r = 1
    for a, b in stats:
        dg.cell(row=r, column=1, value=a).font = Font(name='Arial', bold=(b == ''), size=10)
        dg.cell(row=r, column=2, value=b).font = nf; r += 1
    dg.cell(row=r, column=1, value='Per year:').font = bf; r += 1
    dg.cell(row=r, column=1, value='Year').font = bf; dg.cell(row=r, column=2, value='Rows').font = bf; r += 1
    for y in years:
        n = sum(1 for rec in records if rec['Income_year'] == y)
        dg.cell(row=r, column=1, value=y).font = nf; dg.cell(row=r, column=2, value=n).font = nf; r += 1

    # ---- Legend (static, adjust if the schema changes) ----
    lg = wb.create_sheet('Legend')
    lg.column_dimensions['A'].width = 32; lg.column_dimensions['B'].width = 112
    tf = Font(name='Arial', bold=True, size=13, color='1F4E78')
    sf = Font(name='Arial', bold=True, size=11, color='1F4E78')
    def put(row, a, b):
        ca = lg.cell(row=row, column=1, value=a); ca.font = bf; ca.alignment = Alignment(wrap_text=True, vertical='top')
        cb = lg.cell(row=row, column=2, value=b); cb.font = nf; cb.alignment = Alignment(wrap_text=True, vertical='top')
    r = 1
    lg.cell(row=r, column=1, value='IPCAL data dictionary - personal income tax').font = tf; r += 2
    lg.cell(row=r, column=1, value='STRUCTURE').font = sf; r += 1
    put(r, 'Grain', "One row per IPCAL code. Spouse 1 (prefixes A/C/E/G/I/K) and spouse 2 (B/D/F/H/J/L) sit on two rows linked by IPCAL_code_spouse."); r += 1
    put(r, 'Hierarchy_source / Label_source', "The PDF wins whenever the code appears in one (labels are up to date); otherwise the Excel file is used as fallback (computed/administrative codes)."); r += 1
    put(r, 'Frame / Section_PDF / Item_PDF', "Structural levels of the paper form: Frame is the form's 'Cadre' (Cadre I, II, III...), Item the numbered 'Rubrique' inside a section."); r += 1
    put(r, 'IPType', "90/91 = real-estate income (to be clarified); 92=Netherlands, 93=Germany, 94=Luxembourg (exempt); 95=exempt CSSS; 96=French savings; 97=other PIT-exempt countries."); r += 1
    put(r, 'Labels stay in FR/NL', "Label_FR / Label_Excel_NL and the hierarchy columns reproduce the administration's wording verbatim, in the source language - they are source data, not translated text."); r += 1
    put(r, 'See CLAUDE.md', "For the full construction rules, see CLAUDE.md at the repository root.")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    wb.save(args.out)
    print(f'Written: {args.out} | {len(cols)} columns, {len(records_sorted)} rows')


if __name__ == '__main__':
    main()

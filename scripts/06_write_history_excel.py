#!/usr/bin/env python3
"""
Writes the "per-variable history" workbook -- the AGGREGATED view derived from the
annual dictionary (primary table: one row per code per year) -- from the JSON
produced by 05_build_history.py: one row per generation of an IPCAL code (a single
one for the vast majority of codes, several for codes with a suspected semantic
break), with one label column per year for quick visual scanning, plus a dedicated
sheet listing the split codes unambiguously (the one to check before building a
mapping to national aggregates), and the detail of every semantic break.

A generation's key is (IPCAL_code, Valid_from), in two separate columns: IPCAL_code
is never altered nor suffixed.

Usage:
    python3 06_write_history_excel.py data/history.json -o IPCAL_history_variables.xlsx
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
    ap.add_argument('history', help='JSON produced by 05_build_history.py')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    data = json.load(open(args.history, encoding='utf-8'))
    years = data['years']
    variables = data['variables']
    breaks = data['semantic_breaks']
    split_codes = sorted({v['IPCAL_code'] for v in variables if v['Semantic_break']})
    print(f'{len(variables)} rows (generations), {len(split_codes)} split codes, {len(years)} years ({years[0]}-{years[-1]})')

    wb = openpyxl.Workbook()
    HEAD = PatternFill('solid', fgColor='1F4E78')
    HEADF = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    CELLF = Font(name='Arial', size=9)
    thin = Side(style='thin', color='D9D9D9')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt = PatternFill('solid', fgColor='F2F6FB')
    gap_fill = PatternFill('solid', fgColor='FDE9D9')        # year inside the range but code absent
    flag_fill = PatternFill('solid', fgColor='F8696B')       # semantic break -- highly visible
    flag_font = Font(name='Arial', bold=True, color='FFFFFF', size=9)

    # ---- Sheet 0: Codes_with_break (THE sheet to consult before any mapping) ----
    cs = wb.active
    cs.title = 'Codes_with_break'
    cs_cols = ['IPCAL_code', 'Generations', 'Valid_from', 'Valid_to', 'Label']
    cs.append(['⚠ ' + c if c == 'IPCAL_code' else c for c in cs_cols])
    for c in range(1, len(cs_cols) + 1):
        cell = cs.cell(row=1, column=c)
        cell.fill = flag_fill; cell.font = flag_font
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER
    by_code_gens = {}
    for v in variables:
        if v['Semantic_break']:
            by_code_gens.setdefault(v['IPCAL_code'], []).append(v)
    rr = 2
    for code in split_codes:
        gens = sorted(by_code_gens[code], key=lambda v: v['Generation'])
        for v in gens:
            label = v['Label_by_year'].get(str(v['Valid_to']), v['Label_by_year'].get(v['Valid_to'], ''))
            row = [code, v['Generations_total'], v['Valid_from'], v['Valid_to'], label]
            cs.append([san(x) for x in row])
            for c in range(1, len(cs_cols) + 1):
                cell = cs.cell(row=rr, column=c); cell.border = BORDER
                cell.font = Font(name='Arial', bold=(c == 1), size=9)
                if c == 1:
                    cell.fill = PatternFill('solid', fgColor='FDE9D9')
            rr += 1
    cs.freeze_panes = 'A2'
    cs.auto_filter.ref = f"A1:{get_column_letter(len(cs_cols))}{rr - 1}"
    CSW = {'IPCAL_code': 12, 'Generations': 13, 'Valid_from': 12, 'Valid_to': 11, 'Label': 46}
    for i, k in enumerate(cs_cols, start=1):
        cs.column_dimensions[get_column_letter(i)].width = CSW.get(k, 16)
    cs.row_dimensions[1].height = 30
    if not split_codes:
        cs.cell(row=2, column=1, value="No split code in this dataset.").font = Font(name='Arial', italic=True, size=10)

    # ---- Sheet 1: History (one row per generation) ----
    ws = wb.create_sheet('History')
    fixed_cols = ['IPCAL_code', 'Valid_from', 'Valid_to', 'Semantic_break',
                  'Generation', 'Generations_total',
                  'Prefix', 'Spouse', 'Variable_nature',
                  'Years_present_count', 'Continuity_pct', 'Minor_availability_gap']
    year_cols = [f'Label_{y}' for y in years]
    cols = fixed_cols + year_cols

    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    variables_sorted = sorted(variables, key=lambda v: (v['IPCAL_code'], v['Generation']))
    for i, v in enumerate(variables_sorted):
        row_vals = [v.get(k, '') for k in fixed_cols]
        labels = v['Label_by_year']
        row_vals += [labels.get(str(y), labels.get(y, '')) for y in years]
        ws.append([san(x) for x in row_vals])
        rr2 = i + 2
        gen_years = set(v['Years_present'])
        span = range(v['Valid_from'], v['Valid_to'] + 1)  # validity range of THIS generation
        for c, k in enumerate(cols, start=1):
            cell = ws.cell(row=rr2, column=c); cell.font = CELLF; cell.border = BORDER
            if k == 'Semantic_break':
                cell.alignment = Alignment(horizontal='center')
                if v['Semantic_break']:
                    cell.fill = flag_fill; cell.font = flag_font
                    cell.value = '⚠ YES'
                else:
                    cell.value = 'No'
            elif k == 'Continuity_pct':
                cell.alignment = Alignment(horizontal='right')
            elif k.startswith('Label_'):
                y = int(k.split('_')[1])
                if y not in span:
                    cell.fill = PatternFill('solid', fgColor='E7E6E6')  # outside this generation
                elif y not in gen_years:
                    cell.fill = gap_fill  # minor gap inside the generation
                elif i % 2 == 1:
                    cell.fill = alt
            elif i % 2 == 1 and k != 'Semantic_break':
                cell.fill = alt
    ws.freeze_panes = 'E2'  # keeps the key (IPCAL_code, Valid_from), Valid_to and the flag visible
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(variables_sorted) + 1}"
    W = {'Variable_nature': 32, 'Semantic_break': 13, 'Continuity_pct': 11,
         'Valid_from': 12, 'Valid_to': 11,
         'Minor_availability_gap': 14, 'Generations_total': 12, 'Years_present_count': 12}
    for i, k in enumerate(cols, start=1):
        width = W.get(k, 30 if k.startswith('Label_') else 13)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 40

    # ---- Sheet 2: Semantic_breaks (one row per boundary between generations) ----
    al = wb.create_sheet('Semantic_breaks')
    al_cols = ['IPCAL_code', 'Variable_nature', 'Valid_from_before', 'Valid_from_after', 'missing_years',
               'before_year', 'before_label', 'before_path', 'after_year', 'after_label', 'after_path', 'similarity']
    al_headers = ['IPCAL_code', 'Variable_nature', 'Generation before: Valid_from', 'Generation after: Valid_from',
                  'Missing years',
                  'Before: year', 'Before: label', 'Before: hierarchy', 'After: year', 'After: label',
                  'After: hierarchy', 'Similarity']
    al.append(al_headers)
    for c in range(1, len(al_cols) + 1):
        cell = al.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    breaks_sorted = sorted(breaks, key=lambda g: g['similarity'])  # most suspect first
    for i, g in enumerate(breaks_sorted):
        row = [g['IPCAL_code'], g['Variable_nature'], g['Valid_from_before'], g['Valid_from_after'],
               ','.join(str(y) for y in g['missing_years']), g['before_year'], g['before_label'],
               g['before_path'], g['after_year'], g['after_label'], g['after_path'], g['similarity']]
        al.append([san(x) for x in row])
        rr3 = i + 2
        for c in range(1, len(al_cols) + 1):
            cell = al.cell(row=rr3, column=c); cell.font = CELLF; cell.border = BORDER
            if i % 2 == 1:
                cell.fill = alt
    if breaks_sorted:
        al.freeze_panes = 'E2'
        al.auto_filter.ref = f"A1:{get_column_letter(len(al_cols))}{len(breaks_sorted) + 1}"
    ALW = {'before_label': 36, 'after_label': 36, 'before_path': 50, 'after_path': 50,
           'Variable_nature': 32, 'missing_years': 16,
           'Valid_from_before': 17, 'Valid_from_after': 17}
    for i, k in enumerate(al_cols, start=1):
        al.column_dimensions[get_column_letter(i)].width = ALW.get(k, 13)
    al.row_dimensions[1].height = 30

    # ---- Sheet 3: Legend / method ----
    lg = wb.create_sheet('Legend')
    lg.column_dimensions['A'].width = 34; lg.column_dimensions['B'].width = 112
    tf = Font(name='Arial', bold=True, size=13, color='1F4E78')
    sf = Font(name='Arial', bold=True, size=11, color='1F4E78')
    bf = Font(name='Arial', bold=True, size=10); nf = Font(name='Arial', size=10)
    def put(row, a, b):
        ca = lg.cell(row=row, column=1, value=a); ca.font = bf; ca.alignment = Alignment(wrap_text=True, vertical='top')
        cb = lg.cell(row=row, column=2, value=b); cb.font = nf; cb.alignment = Alignment(wrap_text=True, vertical='top')
    r = 1
    lg.cell(row=r, column=1, value='Multi-year history per semantic variable - IPCAL').font = tf; r += 2
    lg.cell(row=r, column=1, value='PRINCIPLE').font = sf; r += 1
    put(r, 'Two levels', "Primary table = the annual dictionary: one row per code PER YEAR, where (IPCAL_code, Income_year) already disambiguates everything. This workbook is the derived aggregated view: one row per semantic code -- a single one for the ~99% of codes whose meaning never changes, several ('generations') for those with a suspected semantic break."); r += 1
    put(r, 'No synthetic identifier', "IPCAL_code is NEVER altered nor suffixed, in any table. A generation is designated by the two-column composite key (IPCAL_code, Valid_from) -- deliberately not a concatenated string such as \"A0270-2021\", which would look like an IPCAL code without being one."); r += 1
    put(r, 'Semantic_break (⚠)', "True on ALL generations of a split IPCAL_code. Before mapping such a code to a national aggregate, consult the Codes_with_break sheet: as many rows as generations, each with its validity range -- never map IPCAL_code alone in that case."); r += 1
    put(r, 'Joining to the annual data', "IPCAL_code equal AND Income_year between Valid_from and Valid_to. For codes without a break (Semantic_break = No), Valid_from is purely descriptive and the key effectively reduces to IPCAL_code alone."); r += 1
    put(r, 'Valid_from / Valid_to', "First/last known year of THIS generation (not of the whole code when it is split). Valid_from is the 2nd component of the key."); r += 1
    put(r, 'Continuity_pct', "Share of the years in [Valid_from, Valid_to] where the code is actually present (100% = no minor gap)."); r += 1
    put(r, 'Minor_availability_gap', "Availability gap inside this generation, but with a label judged stable (similarity >= 0.6) -- not treated as a change of meaning, hence no split."); r += 1
    put(r, 'Grey cells (History)', "Year outside this generation's validity range (belongs to another generation of the same code)."); r += 1
    put(r, 'Orange cells (History)', "Year inside the validity range but code absent (minor gap)."); r += 1
    put(r, 'Labels stay in FR', "Label_<year> reproduces the administration's wording verbatim, in the source language - source data, not translated text."); r += 1
    lg.cell(row=r, column=1, value='DETECTION THRESHOLD').font = sf; r += 1
    put(r, 'Semantic_break (per gap)', "difflib similarity of label+hierarchy before/after < 0.6, after normalisation (lowercase, punctuation and 4-digit years removed -- see scripts/05_build_history.py, SUSPECT_THRESHOLD). Threshold calibrated on 2017-2023: among the 76 gaps observed, the maximum similarity was 0.559 and corresponded, on manual review, to a genuine change of meaning every time (e.g. A7990 'Option sur action' -> 'Periode 2 : juin - sept'). That finding may not generalise to future data: this is a prioritisation signal, not absolute certainty -- check Before/After in Semantic_breaks before concluding."); r += 1
    lg.cell(row=r, column=1, value='KNOWN LIMITS').font = sf; r += 1
    put(r, 'Observed window', "'Valid_from/to' are relative to the loaded data, not to the code's real existence before/after that window."); r += 1
    put(r, 'Possible false negatives', "A change of meaning WITHOUT an availability gap (code continuously present but whose meaning evolves) is not detected -- not observed over 2017-2023, but not proven impossible. A mapping relying on Semantic_break alone would then be wrong with no warning signal."); r += 1
    put(r, 'See also', "CLAUDE.md at the repository root, section 'Semantic drift across years', and scripts/05_build_history.py for the algorithm details.")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    wb.save(args.out)
    print(f'Written: {args.out} | {len(variables)} rows, {len(split_codes)} split codes, {len(breaks)} semantic breaks')


if __name__ == '__main__':
    main()

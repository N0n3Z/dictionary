#!/usr/bin/env python3
"""
Source coverage check, per year: for every expected document (part 1 PDF per region,
part 2, non-residents, master Excel), reports whether it is present on disk AND
declared in manifest.json, missing, or not applicable for that year (e.g. no regional
split before 2014).

Purpose: spot a collection gap at a glance (e.g. "the Brussels part is missing for
2016") BEFORE running the pipeline on it -- unlike the dictionary's Available_PDF_*
columns, which say nothing about a document that was never received (they answer "is
this code in the PDFs we have", not "do we have all the PDFs we should").

Depends only on config.json + manifest.json + the data/<year>/raw/ tree -- runs
independently of the rest of the pipeline (no need to have run 01-04 first).

Per-cell states:
  OK             file present on disk AND declared in manifest.json
  MISSING        expected this year, neither file nor declaration -> collection gap
  FILE_ABSENT    declared in manifest.json but the referenced file does not exist
  UNDECLARED     file present in raw/ but absent from manifest.json (oversight)
  N/A            not applicable this year (e.g. regional keys before 2014)
  YEAR_ABSENT    data/<year>/ does not exist at all

Usage:
    python3 00_check_sources.py                      # 2014..last year found
    python3 00_check_sources.py --start 2014 --end 2024
    python3 00_check_sources.py -o IPCAL_source_coverage.xlsx
"""
import json, os, re, argparse, glob
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from _layout import load_config, paths_for_year

PRE2014_KEYS = ['P1']
POST2014_CORE = ['P1_BXL', 'P1_RF', 'P1_RW', 'P2']
POST2014_SUPP = ['INR_P1', 'INR_P2']
ALL_DOC_KEYS = PRE2014_KEYS + POST2014_CORE + POST2014_SUPP  # display order
PSEUDO_EXCEL_KEY = 'Master_Excel'

LABELS = {
    'P1': 'Part 1 (single, before regional split)',
    'P1_BXL': 'Part 1 - Brussels', 'P1_RF': 'Part 1 - Flanders', 'P1_RW': 'Part 1 - Wallonia',
    'P2': 'Part 2 (self-employed)',
    'INR_P1': 'Non-residents, part 1', 'INR_P2': 'Non-residents, part 2',
    PSEUDO_EXCEL_KEY: 'Master Excel',
}


def applicable_keys(year):
    return set(PRE2014_KEYS) if year < 2014 else set(POST2014_CORE) | set(POST2014_SUPP)


def is_core(key):
    return key in POST2014_CORE or key in PRE2014_KEYS


def check_year(year, config_path=None):
    paths = paths_for_year(year, config_path)
    year_dir = paths['year_dir']
    result = {'year': year, 'year_absent': not os.path.isdir(year_dir), 'cells': {}}
    if result['year_absent']:
        for k in ALL_DOC_KEYS + [PSEUDO_EXCEL_KEY]:
            result['cells'][k] = 'YEAR_ABSENT'
        return result

    manifest = {}
    if os.path.exists(paths['manifest']):
        manifest = json.load(open(paths['manifest'], encoding='utf-8'))
    declared = manifest.get('documents', {})
    raw_dir = paths['raw_dir']
    applicable = applicable_keys(year)

    for key in ALL_DOC_KEYS:
        if key not in applicable:
            result['cells'][key] = 'N/A'
            continue
        on_disk = any(os.path.exists(os.path.join(raw_dir, key + ext)) for ext in ('.pdf', '.txt'))
        is_declared = key in declared
        decl_file_exists = is_declared and os.path.exists(os.path.join(year_dir, declared[key]))
        if is_declared and (decl_file_exists or on_disk):
            result['cells'][key] = 'OK'
        elif is_declared and not decl_file_exists:
            result['cells'][key] = 'FILE_ABSENT'
        elif on_disk and not is_declared:
            result['cells'][key] = 'UNDECLARED'
        else:
            result['cells'][key] = 'MISSING'

    excel_path = paths.get('excel_master')
    excel_on_disk = bool(excel_path and os.path.exists(excel_path))
    if not excel_on_disk and os.path.isdir(raw_dir):
        excel_on_disk = bool(glob.glob(os.path.join(raw_dir, '*.xlsx')))
    result['cells'][PSEUDO_EXCEL_KEY] = 'OK' if excel_on_disk else 'MISSING'
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--start', type=int, default=2014, help="first income year to check (default 2014, target documented in CLAUDE.md)")
    ap.add_argument('--end', type=int, help="last income year to check (default: last year found in data/, or --start if none)")
    ap.add_argument('--config', help="path to config.json (default: config.json at the repository root)")
    ap.add_argument('-o', '--out', default='IPCAL_source_coverage.xlsx')
    args = ap.parse_args()

    cfg = load_config(args.config)
    data_dir = cfg['data_dir']
    found_years = sorted(int(os.path.basename(p)) for p in glob.glob(os.path.join(data_dir, '*'))
                          if os.path.isdir(p) and re.fullmatch(r'\d{4}', os.path.basename(p)))
    end = args.end if args.end is not None else (max(found_years) if found_years else args.start)

    rows = [check_year(y, args.config) for y in range(args.start, end + 1)]

    # ---- Console report ----
    problems = []
    for r in rows:
        if r['year_absent']:
            problems.append(f"  {r['year']}: folder {data_dir}/{r['year']}/ absent")
            continue
        missing = [LABELS[k] for k, v in r['cells'].items()
                   if v == 'MISSING' and (is_core(k) or k == PSEUDO_EXCEL_KEY)]
        file_absent = [LABELS[k] for k, v in r['cells'].items() if v == 'FILE_ABSENT']
        undeclared = [LABELS[k] for k, v in r['cells'].items() if v == 'UNDECLARED']
        supp_missing = [LABELS[k] for k, v in r['cells'].items() if v == 'MISSING' and k in POST2014_SUPP]
        if missing:
            problems.append(f"  {r['year']}: MISSING (core) -> {', '.join(missing)}")
        if file_absent:
            problems.append(f"  {r['year']}: declared but file absent -> {', '.join(file_absent)}")
        if undeclared:
            problems.append(f"  {r['year']}: file present but undeclared in manifest.json -> {', '.join(undeclared)}")
        if supp_missing:
            problems.append(f"  {r['year']}: absent (supplementary, non-residents) -> {', '.join(supp_missing)}")

    print(f'Source coverage check, {args.start}-{end} ({len(rows)} years):')
    if problems:
        print(f'{len(problems)} issue(s):')
        for p in problems:
            print(p)
    else:
        print('  No issue: every expected core source is present and declared.')

    write_excel(rows, args.out)
    print(f'Written: {args.out}')


def write_excel(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Source_coverage'

    HEAD = PatternFill('solid', fgColor='1F4E78')
    HEADF = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    CELLF = Font(name='Arial', size=9, bold=True)
    thin = Side(style='thin', color='D9D9D9')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    FILL = {
        'OK': PatternFill('solid', fgColor='C6E0B4'),
        'MISSING': PatternFill('solid', fgColor='F8696B'),
        'FILE_ABSENT': PatternFill('solid', fgColor='F8696B'),
        'UNDECLARED': PatternFill('solid', fgColor='FFD966'),
        'N/A': PatternFill('solid', fgColor='E7E6E6'),
        'YEAR_ABSENT': PatternFill('solid', fgColor='808080'),
    }
    TEXT = {'OK': 'OK', 'MISSING': 'MISSING', 'FILE_ABSENT': 'FILE ABSENT',
            'UNDECLARED': 'UNDECLARED', 'N/A': '—', 'YEAR_ABSENT': 'YEAR ABSENT'}

    cols = ['Income_year'] + ALL_DOC_KEYS + [PSEUDO_EXCEL_KEY]
    header_labels = ['Income year'] + [LABELS[k] for k in ALL_DOC_KEYS] + [LABELS[PSEUDO_EXCEL_KEY]]
    ws.append(header_labels)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD; cell.font = HEADF
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER

    for i, r in enumerate(rows):
        rr = i + 2
        cell = ws.cell(row=rr, column=1, value=r['year'])
        cell.font = Font(name='Arial', bold=True, size=9); cell.border = BORDER
        cell.alignment = Alignment(horizontal='center')
        for c, key in enumerate(ALL_DOC_KEYS + [PSEUDO_EXCEL_KEY], start=2):
            state = r['cells'][key]
            cell = ws.cell(row=rr, column=c, value=TEXT[state])
            cell.font = CELLF; cell.border = BORDER; cell.fill = FILL[state]
            cell.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(rows) + 1}"
    ws.column_dimensions['A'].width = 13
    for i in range(2, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
    ws.row_dimensions[1].height = 46

    lg = wb.create_sheet('Legend')
    lg.column_dimensions['A'].width = 20; lg.column_dimensions['B'].width = 100
    bf = Font(name='Arial', bold=True, size=10); nf = Font(name='Arial', size=10)
    tf = Font(name='Arial', bold=True, size=13, color='1F4E78')
    r = 1
    lg.cell(row=r, column=1, value='Raw source coverage - IPCAL').font = tf; r += 2
    entries = [
        ('OK', "File present on disk and declared in manifest.json."),
        ('MISSING', "Expected this year (core: regions / part 2 / Excel), neither file nor declaration -- a collection gap to fill."),
        ('FILE ABSENT', "Declared in manifest.json but the referenced file does not exist -- broken path or file never delivered."),
        ('UNDECLARED', "The file exists in raw/ but manifest.json does not mention it -- the pipeline will ignore it until the manifest is updated."),
        ('—  (N/A)', "Not applicable this year (e.g. regional keys before 2014, when a single PDF existed -- see CLAUDE.md)."),
        ('YEAR ABSENT', "data/<year>/ does not exist at all in the repository."),
    ]
    for label, desc in entries:
        c1 = lg.cell(row=r, column=1, value=label); c1.font = bf; c1.alignment = Alignment(wrap_text=True, vertical='top')
        c2 = lg.cell(row=r, column=2, value=desc); c2.font = nf; c2.alignment = Alignment(wrap_text=True, vertical='top')
        r += 1
    r += 1
    lg.cell(row=r, column=1, value='Non-residents (INR)').font = bf; r += 1
    lg.cell(row=r, column=2, value="Classified as \"supplementary\" rather than \"core\": it is not known with certainty whether they exist for every year (only 2023 has been provided so far). A MISSING on those columns is therefore informative, not necessarily a collection issue.").font = nf
    lg.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical='top')

    wb.save(out_path)


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""
Builds a per-variable history view (by IPCAL_code) from the records.json files of
several years: availability per year, label evolution, and semantic-drift detection.

Assumption validated on 2017-2023 (see CLAUDE.md, "Semantic drift across years"):
IPCAL codes are not renumbered between years in this dataset (IPCAL_code_previous ==
IPCAL_code in 100% of non-"new" cases) -> an IPCAL_code is treated as the stable
identifier of a "semantic variable". If a future year introduces renumbering
(IPCAL_code_previous differing from IPCAL_code), this script must be extended to
chain codes through that field rather than by strict equality.

Two-level architecture, product decision (do not revisit without explicit instruction):
 1. primary table = the annual dictionary (03/04): ONE ROW PER CODE PER YEAR, where
    (IPCAL_code, Income_year) already disambiguates fully (a 2017 row for A0270
    unambiguously denotes its 2017 meaning);
 2. aggregated view = this script: one row per "semantic code" -- a single one for
    the ~99% of codes whose meaning never changes, several ("generations") for those
    with a suspected semantic break.

NO synthetic identifier: IPCAL_code is never altered nor suffixed, in any table.
Disambiguating a split code is done through a COMPOSITE KEY of two columns --
(IPCAL_code, Valid_from) -- rather than a concatenated string such as "A0270-2021",
which would look like an IPCAL code without being one and would invite confusion.
To join this view to the annual dictionary: IPCAL_code equal AND Income_year between
Valid_from and Valid_to. For codes without a break, Valid_from is purely descriptive
and the key effectively reduces to IPCAL_code alone.

Semantic_break (boolean, on EVERY generation of a split code) is THE signal to check
before using IPCAL_code alone in a mapping: True => never map that code without
specifying a year or validity range.

Two distinct signals:
- availability gap: the code is missing for >=1 year between its first and last
  observed appearance in the loaded dataset, then reappears. This is a fact, not a
  heuristic.
- Semantic_break (at the level of one gap): on top of the availability gap, the
  label/hierarchy before and after the gap are dissimilar (difflib comparison after
  normalisation: lowercase, punctuation and 4-digit years removed -- a phrase such as
  "married in 2024" must not read as a change of meaning merely because the year
  cited changes). This is a prioritisation heuristic for manual review, not a
  certainty (observed: pure rewordings of the Excel hierarchy from one edition to the
  next, without any real change of meaning, can also produce a low similarity -- see
  the output workbook's Legend). A gap judged suspect splits the code into two
  generations; a non-suspect gap (label stayed similar) remains a plain temporary
  absence, reported within the generation concerned.

Usage:
    python3 05_build_history.py data/*/records.json -o data/history.json
"""
import json, argparse, glob, re
from difflib import SequenceMatcher
from collections import defaultdict

YEAR_RE = re.compile(r'\b(19|20)\d{2}\b')
PUNCT_RE = re.compile(r'[^\w\s]', re.UNICODE)
SPACE_RE = re.compile(r'\s+')

SUSPECT_THRESHOLD = 0.6  # similarity below which a gap is judged suspect (=> splits into 2 generations)
# Calibrated empirically on 2017-2023: among the 76 gaps observed, the label/hierarchy
# similarity before->after peaks at 0.559 (max) -- even the "least suspect" pairs by that
# score correspond, on manual review, to a real change of meaning (e.g. A7990 "Option sur
# action" -> "Periode 2 : juin - sept"). A 0.45 threshold let 4 clearly drifted cases
# through as "not suspect" (false negatives); 0.6 catches them all while leaving room for
# a future case where the label stays nearly identical after a gap (which would then,
# rightly, not be a drift).


def normalize(txt):
    txt = (txt or '').lower()
    txt = YEAR_RE.sub('', txt)
    txt = PUNCT_RE.sub(' ', txt)
    txt = SPACE_RE.sub(' ', txt).strip()
    return txt


def similarity(a, b):
    a, b = normalize(a), normalize(b)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def load_records(patterns):
    by_year = {}
    for pat in patterns:
        paths = sorted(glob.glob(pat)) if any(c in pat for c in '*?[') else [pat]
        for path in paths:
            recs = json.load(open(path, encoding='utf-8'))
            if not recs:
                continue
            y = recs[0]['Income_year']
            by_year[y] = recs
    return by_year


def gaps_in_span(years_present):
    """Runs of missing years strictly between the first and last present year of a
    given range."""
    if len(years_present) < 2:
        return []
    lo, hi = years_present[0], years_present[-1]
    missing = [y for y in range(lo, hi + 1) if y not in years_present]
    runs, cur = [], []
    for y in missing:
        if cur and y == cur[-1] + 1:
            cur.append(y)
        else:
            if cur:
                runs.append(cur)
            cur = [y]
    if cur:
        runs.append(cur)
    return runs


def build(by_year):
    all_years = sorted(by_year)
    by_code = defaultdict(dict)
    for y in all_years:
        for r in by_year[y]:
            by_code[r['IPCAL_code']][y] = r

    variables = []
    breaks = []
    for code, per_year in sorted(by_code.items()):
        years_present = sorted(per_year)

        # 1) every gap over the code's full span, suspect or not
        gap_runs = gaps_in_span(years_present)
        gaps = []
        for run in gap_runs:
            before_year, after_year = run[0] - 1, run[-1] + 1
            before, after = per_year.get(before_year), per_year.get(after_year)
            if not (before and after):
                continue
            sim = similarity(
                before['Label_FR'] + ' ' + before['Hierarchy_path'],
                after['Label_FR'] + ' ' + after['Hierarchy_path'],
            )
            gaps.append({
                'missing_years': run, 'before_year': before_year, 'after_year': after_year,
                'before_label': before['Label_FR'], 'after_label': after['Label_FR'],
                'before_path': before['Hierarchy_path'], 'after_path': after['Hierarchy_path'],
                'similarity': round(sim, 3),
                'suspect': sim < SUSPECT_THRESHOLD,
            })

        # 2) split into generations at suspect gaps only
        split_after = {g['before_year'] for g in gaps if g['suspect']}
        generations, current = [], []
        for y in years_present:
            current.append(y)
            if y in split_after:
                generations.append(current)
                current = []
        if current:
            generations.append(current)
        nb_gen = len(generations)
        semantic_break = nb_gen > 1

        for gi, gen_years in enumerate(generations, start=1):
            gfirst, glast = gen_years[0], gen_years[-1]
            gen_gap_runs = gaps_in_span(gen_years)  # minor (non-suspect) gaps inside this generation
            gen_gaps = [g for g in gaps if g['before_year'] >= gfirst and g['after_year'] <= glast]
            min_sim = min((g['similarity'] for g in gen_gaps), default=None)
            latest = per_year[glast]

            variables.append({
                # Key: (IPCAL_code, Valid_from). IPCAL_code is never altered.
                'IPCAL_code': code,
                'Generation': gi,
                'Generations_total': nb_gen,
                'Semantic_break': semantic_break,
                'Prefix': latest['Prefix'],
                'Spouse': latest['Spouse'],
                'Variable_nature': latest['Variable_nature'],
                'Valid_from': gfirst,
                'Valid_to': glast,
                'Years_present': gen_years,
                'Years_present_count': len(gen_years),
                'Continuity_pct': round(100 * len(gen_years) / (glast - gfirst + 1), 1),
                'Minor_availability_gap': bool(gen_gap_runs),
                'Minor_gaps_count': len(gen_gap_runs),
                'Min_similarity_minor_gaps': round(min_sim, 3) if min_sim is not None else None,
                'Minor_gap_details': gen_gaps,
                'Label_by_year': {y: per_year[y]['Label_FR'] for y in gen_years},
                'Path_by_year': {y: per_year[y]['Hierarchy_path'] for y in gen_years},
                'Source_by_year': {y: per_year[y]['Hierarchy_source'] for y in gen_years},
            })

        # 3) semantic-break events (= boundaries between generations), for the workbook
        if semantic_break:
            gen_of_year = {y: gi for gi, gy in enumerate(generations, start=1) for y in gy}
            for g in gaps:
                if not g['suspect']:
                    continue
                gi_before, gi_after = gen_of_year[g['before_year']], gen_of_year[g['after_year']]
                breaks.append({
                    'IPCAL_code': code, 'Variable_nature': per_year[years_present[-1]]['Variable_nature'],
                    # The two generations on either side of the gap, designated by their
                    # Valid_from (2nd component of the composite key), not by a
                    # concatenated string.
                    'Valid_from_before': generations[gi_before - 1][0],
                    'Valid_from_after': generations[gi_after - 1][0],
                    **g,
                })
    return all_years, variables, breaks


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('records', nargs='+', help='one or more records.json (globs accepted)')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    by_year = load_records(args.records)
    if not by_year:
        raise SystemExit('No records loaded - check the paths.')
    all_years, variables, breaks = build(by_year)

    json.dump({'years': all_years, 'variables': variables, 'semantic_breaks': breaks},
               open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)

    split_codes = {v['IPCAL_code'] for v in variables if v['Semantic_break']}
    print(f'{len({v["IPCAL_code"] for v in variables})} codes over {len(all_years)} years ({all_years[0]}-{all_years[-1]})')
    print(f'  {len(variables)} rows (generations) in total')
    print(f'  codes with a semantic break (split): {len(split_codes)} -> {len(breaks)} gap(s) in total')
    print(f'Written: {args.out}')


if __name__ == '__main__':
    main()

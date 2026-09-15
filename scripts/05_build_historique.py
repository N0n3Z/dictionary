#!/usr/bin/env python3
"""
Construit une vue historique par variable sémantique (Code_IPCAL) à partir des
records.json de plusieurs années : disponibilité par année, évolution du libellé,
et détection de dérive sémantique.

Hypothèse validée sur 2017-2023 (voir CLAUDE.md, "Dérive sémantique multi-année") :
les codes IPCAL ne sont pas renumérotés d'une année à l'autre dans ce jeu de
données (Code_IPCAL_precedent == Code_IPCAL dans 100% des cas non-"nouveaux") ->
un Code_IPCAL est traité comme l'identifiant stable d'une "variable sémantique".
Si une future année introduit une renumérotation (Code_IPCAL_precedent différent
de Code_IPCAL), ce script devra être étendu pour chaîner les codes via ce champ
plutôt que par égalité stricte.

Deux signaux distincts :
- Rupture_disponibilite : le code est absent >=1 an entre sa première et sa
  dernière apparition observée dans le jeu de données chargé, puis réapparaît.
  C'est un fait, pas une heuristique.
- Rupture_semantique_suspectee : en plus de la rupture de disponibilité, le
  libellé/la hiérarchie d'avant et d'après la coupure sont peu similaires
  (comparaison difflib après normalisation : minuscules, ponctuation et années
  à 4 chiffres retirées -- une phrase du type "mariés en 2024" ne doit pas être
  vue comme un changement de sens du seul fait que l'année citée change).
  C'est une heuristique de priorisation pour revue manuelle, pas une certitude
  (constaté : des reformulations pures de la hiérarchie Excel d'une édition à
  l'autre, sans rupture de sens réelle, peuvent aussi produire une similarité
  basse -- cf. Legende du classeur de sortie).

Usage :
    python3 05_build_historique.py data/*/records.json -o data/historique.json
"""
import json, argparse, glob, re
from difflib import SequenceMatcher
from collections import defaultdict

YEAR_RE = re.compile(r'\b(19|20)\d{2}\b')
PUNCT_RE = re.compile(r'[^\w\s]', re.UNICODE)
SPACE_RE = re.compile(r'\s+')

SEUIL_SUSPECT = 0.6  # similarité en-dessous de laquelle une rupture est jugée suspecte
# Calibré empiriquement sur 2017-2023 : parmi les 76 ruptures observées, la similarité
# libellé/hiérarchie avant->après culmine à 0.559 (max) -- même les paires les "moins
# suspectes" au sens de ce score correspondent, à relecture manuelle, à un changement de
# sens réel (ex. A7990 "Option sur action" -> "Periode 2 : juin - sept"). Un seuil de 0.45
# laissait passer 4 cas clairement dérivés comme "non suspects" (faux négatifs) ; 0.6 les
# capture tous tout en laissant la place à un futur cas où le libellé resterait quasi
# identique après une coupure (auquel cas ce ne serait, à raison, pas une dérive).


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
            y = recs[0]['Annee_revenus']
            by_year[y] = recs
    return by_year


def gaps_in_span(years_present):
    """Runs d'années manquantes strictement entre la première et la dernière
    année présente (une absence après la dernière année n'est pas une "rupture"
    au sens de ce script -- juste une possible fin de vie de la variable)."""
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
            by_code[r['Code_IPCAL']][y] = r

    out = []
    for code, per_year in sorted(by_code.items()):
        years_present = sorted(per_year)
        first, last = years_present[0], years_present[-1]
        gap_runs = gaps_in_span(years_present)

        alerts = []
        min_sim = None
        for run in gap_runs:
            before_year, after_year = run[0] - 1, run[-1] + 1
            before, after = per_year.get(before_year), per_year.get(after_year)
            if not (before and after):
                continue
            sim = similarity(
                before['Libelle_FR'] + ' ' + before['Chemin_hierarchique'],
                after['Libelle_FR'] + ' ' + after['Chemin_hierarchique'],
            )
            min_sim = sim if min_sim is None else min(min_sim, sim)
            alerts.append({
                'annees_absentes': run, 'avant_annee': before_year, 'apres_annee': after_year,
                'avant_libelle': before['Libelle_FR'], 'apres_libelle': after['Libelle_FR'],
                'avant_chemin': before['Chemin_hierarchique'], 'apres_chemin': after['Chemin_hierarchique'],
                'similarite': round(sim, 3),
                'suspect': sim < SEUIL_SUSPECT,
            })

        latest = per_year[last]
        out.append({
            'Code_IPCAL': code,
            'Prefixe': latest['Prefixe'],
            'Conjoint': latest['Conjoint'],
            'Nature_variable': latest['Nature_variable'],
            'Premiere_annee_connue': first,
            'Derniere_annee_connue': last,
            'Annees_presentes': years_present,
            'Nb_annees_presentes': len(years_present),
            'Continuite_pct': round(100 * len(years_present) / (last - first + 1), 1),
            'Rupture_disponibilite': bool(gap_runs),
            'Nb_ruptures': len(gap_runs),
            'Rupture_semantique_suspectee': any(a['suspect'] for a in alerts),
            'Similarite_min_avant_apres_rupture': round(min_sim, 3) if min_sim is not None else None,
            'Alertes_rupture': alerts,
            'Libelle_par_annee': {y: per_year[y]['Libelle_FR'] for y in years_present},
            'Chemin_par_annee': {y: per_year[y]['Chemin_hierarchique'] for y in years_present},
            'Source_par_annee': {y: per_year[y]['Source_hierarchie'] for y in years_present},
        })
    return all_years, out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('records', nargs='+', help='un ou plusieurs records.json (globs acceptés)')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    by_year = load_records(args.records)
    if not by_year:
        raise SystemExit('Aucun enregistrement chargé — vérifier les chemins.')
    all_years, out = build(by_year)

    json.dump({'annees': all_years, 'variables': out}, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)

    n_gap = sum(1 for v in out if v['Rupture_disponibilite'])
    n_susp = sum(1 for v in out if v['Rupture_semantique_suspectee'])
    print(f'{len(out)} variables (codes) sur {len(all_years)} années ({all_years[0]}-{all_years[-1]})')
    print(f'  avec au moins une rupture de disponibilité : {n_gap}')
    print(f'  dont dérive sémantique suspectée (libellé très différent après coupure) : {n_susp}')
    print(f'Écrit : {args.out}')


if __name__ == '__main__':
    main()

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

Architecture en deux niveaux, décision utilisateur (ne pas y revenir sans consigne
explicite) :
 1. table primaire = le Dictionnaire annuel (03/04) : UNE LIGNE PAR CODE PAR ANNÉE,
    où (Code_IPCAL, Annee_revenus) désambiguïse déjà totalement (une ligne 2017 pour
    A0270 désigne sans ambiguïté sa signification 2017) ;
 2. vue agrégée = ce script : une ligne par "code sémantique" -- une seule pour les
    ~99% de codes dont le sens ne change pas, plusieurs ("générations") pour ceux
    dont une rupture sémantique est suspectée.

AUCUN identifiant synthétique : Code_IPCAL n'est jamais altéré ni suffixé, dans
aucune table. La désambiguïsation d'un code scindé se fait par CLÉ COMPOSITE en deux
colonnes -- (Code_IPCAL, Validite_debut) -- et non par une chaîne concaténée du type
"A0270-2021", qui ressemblerait à un code IPCAL sans en être un et inviterait à la
confusion. Pour joindre cette vue au Dictionnaire annuel : Code_IPCAL égal ET
Annee_revenus compris entre Validite_debut et Validite_fin. Pour les codes sans
rupture, Validite_debut est purement descriptif et la clé se réduit de fait à
Code_IPCAL seul.

Rupture_semantique (booléen, sur CHAQUE génération d'un code scindé) est LE signal
à regarder avant d'utiliser Code_IPCAL seul dans un mapping : True => ne jamais
mapper ce code sans préciser une année/plage de validité.

Deux signaux distincts :
- Rupture_disponibilite : le code est absent >=1 an entre sa première et sa
  dernière apparition observée dans le jeu de données chargé, puis réapparaît.
  C'est un fait, pas une heuristique.
- Rupture_semantique (au niveau d'une coupure) : en plus de la rupture de
  disponibilité, le libellé/la hiérarchie d'avant et d'après la coupure sont peu
  similaires (comparaison difflib après normalisation : minuscules, ponctuation et
  années à 4 chiffres retirées -- une phrase du type "mariés en 2024" ne doit pas
  être vue comme un changement de sens du seul fait que l'année citée change).
  C'est une heuristique de priorisation pour revue manuelle, pas une certitude
  (constaté : des reformulations pures de la hiérarchie Excel d'une édition à
  l'autre, sans rupture de sens réelle, peuvent aussi produire une similarité
  basse -- cf. Legende du classeur de sortie). Une coupure jugée suspecte scinde le
  code en deux générations ; une coupure non suspecte (libellé resté similaire)
  reste une simple absence temporaire signalée dans la génération concernée.

Usage :
    python3 05_build_historique.py data/*/records.json -o data/historique.json
"""
import json, argparse, glob, re
from difflib import SequenceMatcher
from collections import defaultdict

YEAR_RE = re.compile(r'\b(19|20)\d{2}\b')
PUNCT_RE = re.compile(r'[^\w\s]', re.UNICODE)
SPACE_RE = re.compile(r'\s+')

SEUIL_SUSPECT = 0.6  # similarité en-dessous de laquelle une rupture est jugée suspecte (=> scinde en 2 générations)
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
    année présente d'une plage donnée."""
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

    variables = []
    ruptures = []
    for code, per_year in sorted(by_code.items()):
        years_present = sorted(per_year)

        # 1) toutes les coupures sur le span complet du code, suspectes ou non
        gap_runs = gaps_in_span(years_present)
        alerts = []
        for run in gap_runs:
            before_year, after_year = run[0] - 1, run[-1] + 1
            before, after = per_year.get(before_year), per_year.get(after_year)
            if not (before and after):
                continue
            sim = similarity(
                before['Libelle_FR'] + ' ' + before['Chemin_hierarchique'],
                after['Libelle_FR'] + ' ' + after['Chemin_hierarchique'],
            )
            alerts.append({
                'annees_absentes': run, 'avant_annee': before_year, 'apres_annee': after_year,
                'avant_libelle': before['Libelle_FR'], 'apres_libelle': after['Libelle_FR'],
                'avant_chemin': before['Chemin_hierarchique'], 'apres_chemin': after['Chemin_hierarchique'],
                'similarite': round(sim, 3),
                'suspect': sim < SEUIL_SUSPECT,
            })

        # 2) scission en générations uniquement aux coupures suspectes
        split_after = {a['avant_annee'] for a in alerts if a['suspect']}
        generations, current = [], []
        for y in years_present:
            current.append(y)
            if y in split_after:
                generations.append(current)
                current = []
        if current:
            generations.append(current)
        nb_gen = len(generations)
        rupture_semantique = nb_gen > 1

        for gi, gen_years in enumerate(generations, start=1):
            gfirst, glast = gen_years[0], gen_years[-1]
            gen_gap_runs = gaps_in_span(gen_years)  # coupures mineures (non suspectes) internes à cette génération
            gen_alerts = [a for a in alerts if a['avant_annee'] >= gfirst and a['apres_annee'] <= glast]
            min_sim = min((a['similarite'] for a in gen_alerts), default=None)
            latest = per_year[glast]

            variables.append({
                # Clé : (Code_IPCAL, Validite_debut). Code_IPCAL n'est jamais altéré.
                'Code_IPCAL': code,
                'Generation': gi,
                'Nb_generations_total': nb_gen,
                'Rupture_semantique': rupture_semantique,
                'Prefixe': latest['Prefixe'],
                'Conjoint': latest['Conjoint'],
                'Nature_variable': latest['Nature_variable'],
                'Validite_debut': gfirst,
                'Validite_fin': glast,
                'Annees_presentes': gen_years,
                'Nb_annees_presentes': len(gen_years),
                'Continuite_pct': round(100 * len(gen_years) / (glast - gfirst + 1), 1),
                'Rupture_disponibilite_mineure': bool(gen_gap_runs),
                'Nb_ruptures_mineures': len(gen_gap_runs),
                'Similarite_min_coupures_mineures': round(min_sim, 3) if min_sim is not None else None,
                'Alertes_mineures': gen_alerts,
                'Libelle_par_annee': {y: per_year[y]['Libelle_FR'] for y in gen_years},
                'Chemin_par_annee': {y: per_year[y]['Chemin_hierarchique'] for y in gen_years},
                'Source_par_annee': {y: per_year[y]['Source_hierarchie'] for y in gen_years},
            })

        # 3) événements de rupture sémantique (= frontières entre générations), pour le classeur
        if rupture_semantique:
            gen_of_year = {y: gi for gi, gy in enumerate(generations, start=1) for y in gy}
            for a in alerts:
                if not a['suspect']:
                    continue
                gi_avant, gi_apres = gen_of_year[a['avant_annee']], gen_of_year[a['apres_annee']]
                ruptures.append({
                    'Code_IPCAL': code, 'Nature_variable': per_year[years_present[-1]]['Nature_variable'],
                    # Les deux générations de part et d'autre de la coupure, désignées par
                    # leur Validite_debut (2e composant de la clé composite), pas par une
                    # chaîne concaténée.
                    'Validite_debut_avant': generations[gi_avant - 1][0],
                    'Validite_debut_apres': generations[gi_apres - 1][0],
                    **a,
                })
    return all_years, variables, ruptures


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('records', nargs='+', help='un ou plusieurs records.json (globs acceptés)')
    ap.add_argument('-o', '--out', required=True)
    args = ap.parse_args()

    by_year = load_records(args.records)
    if not by_year:
        raise SystemExit('Aucun enregistrement chargé — vérifier les chemins.')
    all_years, variables, ruptures = build(by_year)

    json.dump({'annees': all_years, 'variables': variables, 'ruptures_semantiques': ruptures},
               open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)

    codes_scindes = {v['Code_IPCAL'] for v in variables if v['Rupture_semantique']}
    print(f'{len({v["Code_IPCAL"] for v in variables})} codes sur {len(all_years)} années ({all_years[0]}-{all_years[-1]})')
    print(f'  {len(variables)} lignes (générations) au total')
    print(f'  codes avec rupture sémantique (scindés) : {len(codes_scindes)} -> {len(ruptures)} coupure(s) au total')
    print(f'Écrit : {args.out}')


if __name__ == '__main__':
    main()

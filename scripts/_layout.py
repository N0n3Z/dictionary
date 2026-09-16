"""
Résolution des chemins du projet à partir de config.json (racine du dépôt), pour
permettre de faire tourner le pipeline en local avec une disposition de dossiers
différente (ou une autre localisation des données) sans toucher au code des
scripts 01-06.

Chaque script accepte --annee (préremplit ses chemins par défaut depuis
config.json) + --config (chemin explicite vers config.json). Les options de
chemin existantes (--excel, --struct, -o, le manifest positionnel, ...) restent
prioritaires quand elles sont fournies : --annee ne fait que préremplir des
valeurs par défaut, il ne les impose jamais.

Priorité de résolution de config.json : --config explicite > variable d'env.
IPCAL_CONFIG > config.json à la racine du dépôt (à côté de scripts/) > gabarit
intégré ci-dessous (défaut historique : data/<année>/raw, manifest.json, etc.
— identique à la disposition utilisée jusqu'ici, donc aucun changement de
comportement pour ce dépôt tant que config.json n'est pas modifié).

Gabarits : chaînes avec {annee}, {data_dir}, ou le nom de toute autre entrée de
"layout" (résolution multi-passe, donc {raw_dir} peut lui-même référencer
{year_dir} qui référence {data_dir}).
"""
import json, os

DEFAULT_LAYOUT = {
    "data_dir": "data",
    "layout": {
        "year_dir": "{data_dir}/{annee}",
        "raw_dir": "{year_dir}/raw",
        "manifest": "{year_dir}/manifest.json",
        "pdf_struct": "{year_dir}/pdf_struct.json",
        "records": "{year_dir}/records.json",
        "excel_master": "{raw_dir}/IPCAL_{annee}.xlsx",
        "workbook_annee": "{year_dir}/IPCAL_data_dictionary_{annee}.xlsx",
    },
}


def _repo_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _default_config_path():
    cand = os.path.join(_repo_root(), 'config.json')
    return cand if os.path.exists(cand) else None


def load_config(config_path=None):
    path = config_path or os.environ.get('IPCAL_CONFIG') or _default_config_path()
    merged = json.loads(json.dumps(DEFAULT_LAYOUT))  # copie profonde du gabarit intégré
    if not path or not os.path.exists(path):
        return merged
    cfg = json.load(open(path, encoding='utf-8'))
    merged.update({k: v for k, v in cfg.items() if k != 'layout'})
    merged['layout'].update(cfg.get('layout', {}))
    return merged


def paths_for_year(annee, config_path=None):
    """Retourne un dict {nom_entree_layout: chemin_resolu} pour une année donnée,
    relatif à la racine du dépôt (chdir non requis : les scripts appelants
    utilisent ces chemins tels quels, relatifs au répertoire courant d'exécution,
    comme le reste du pipeline aujourd'hui)."""
    cfg = load_config(config_path)
    layout = cfg['layout']
    resolved = {}
    remaining = dict(layout)
    for _ in range(len(layout) + 1):
        progressed = False
        for k, tmpl in list(remaining.items()):
            try:
                resolved[k] = tmpl.format(annee=annee, data_dir=cfg['data_dir'], **resolved)
                del remaining[k]
                progressed = True
            except KeyError:
                continue
        if not remaining or not progressed:
            break
    if remaining:
        raise ValueError(f"config.json: entrées de layout non résolues (référence circulaire ou clé inconnue ?) : {list(remaining)}")
    return resolved


def add_annee_arg(ap):
    ap.add_argument('--annee', type=int,
                     help="Année revenus : préremplit les chemins par défaut depuis config.json "
                          "(chaque option de chemin explicite garde priorité sur cette valeur)")
    ap.add_argument('--config', help="Chemin vers config.json (défaut : config.json à la racine du "
                                      "dépôt si présent, sinon la disposition data/<année>/... historique)")

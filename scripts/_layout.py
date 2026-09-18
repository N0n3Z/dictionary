"""
Resolves project paths from config.json (repository root), so the pipeline can run
locally against a different folder layout (or a different data location) without
touching the code of scripts 01-06.

Every script accepts --year (pre-fills its default paths from config.json) and
--config (explicit path to config.json). Existing path options (--excel, --struct,
-o, the positional manifest, ...) always take precedence when supplied: --year only
pre-fills defaults, it never overrides them.

config.json resolution order: explicit --config > IPCAL_CONFIG environment variable
> config.json at the repository root (next to scripts/) > the built-in template
below (historical default: data/<year>/raw, manifest.json, ... -- identical to the
layout used so far, so behaviour is unchanged for this repository until config.json
is edited).

Templates are strings containing {year}, {data_dir}, or the name of any other
"layout" entry (multi-pass resolution, so {raw_dir} may itself reference {year_dir},
which references {data_dir}).
"""
import json, os

DEFAULT_LAYOUT = {
    "data_dir": "data",
    "layout": {
        "year_dir": "{data_dir}/{year}",
        "raw_dir": "{year_dir}/raw",
        "manifest": "{year_dir}/manifest.json",
        "pdf_struct": "{year_dir}/pdf_struct.json",
        "records": "{year_dir}/records.json",
        "excel_master": "{raw_dir}/IPCAL_{year}.xlsx",
        "workbook_year": "{year_dir}/IPCAL_data_dictionary_{year}.xlsx",
    },
}


def _repo_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _default_config_path():
    cand = os.path.join(_repo_root(), 'config.json')
    return cand if os.path.exists(cand) else None


def load_config(config_path=None):
    path = config_path or os.environ.get('IPCAL_CONFIG') or _default_config_path()
    merged = json.loads(json.dumps(DEFAULT_LAYOUT))  # deep copy of the built-in template
    if not path or not os.path.exists(path):
        return merged
    cfg = json.load(open(path, encoding='utf-8'))
    merged.update({k: v for k, v in cfg.items() if k != 'layout'})
    merged['layout'].update(cfg.get('layout', {}))
    return merged


def paths_for_year(year, config_path=None):
    """Return {layout_entry_name: resolved_path} for a given year, relative to the
    repository root (no chdir required: callers use these paths as-is, relative to
    the current working directory, like the rest of the pipeline)."""
    cfg = load_config(config_path)
    layout = cfg['layout']
    resolved = {}
    remaining = dict(layout)
    for _ in range(len(layout) + 1):
        progressed = False
        for k, tmpl in list(remaining.items()):
            try:
                resolved[k] = tmpl.format(year=year, data_dir=cfg['data_dir'], **resolved)
                del remaining[k]
                progressed = True
            except KeyError:
                continue
        if not remaining or not progressed:
            break
    if remaining:
        raise ValueError(f"config.json: unresolved layout entries (circular reference or unknown key?): {list(remaining)}")
    return resolved


def add_year_arg(ap):
    ap.add_argument('--year', type=int,
                     help="Income year: pre-fills default paths from config.json "
                          "(any explicit path option takes precedence over this)")
    ap.add_argument('--config', help="Path to config.json (default: config.json at the repository "
                                      "root if present, otherwise the historical data/<year>/... layout)")

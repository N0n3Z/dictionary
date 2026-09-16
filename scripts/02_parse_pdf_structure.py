#!/usr/bin/env python3
"""
Reconstruit la structure Cadre -> Section -> Rubrique -> Libellé à partir du texte
OCR des documents préparatoires IPCAL, et associe cette structure à chaque code
déclaration (format NNNN-CC) trouvé dans le texte.

Entrée : un manifest JSON décrivant, pour une année donnée, les documents à traiter.
Chaque document peut être soit :
  - un .txt produit par 01_extract_pdf_text.py (avec marqueurs "=== PAGE N ===")
  - une archive "pdf" zip d'origine (texte OCR extrait à la volée)
  - un vrai PDF (fallback pdftotext, comme 01_extract_pdf_text.py)

Exemple de manifest (data/2024/manifest.json) :
{
  "annee_revenus": 2024,
  "exercice_imposition": 2025,
  "documents": {
    "P1_BXL":  "P1_bxl.txt",
    "P1_RF":   "P1_rf.txt",
    "P1_RW":   "P1_rw.txt",
    "P2":      "P2.txt",
    "INR_P1":  "INR_P1.txt",
    "INR_P2":  "INR_P2.txt"
  }
}
Les clés de "documents" sont libres (P1_BXL, P1_RF, ... sont juste la convention
utilisée jusqu'ici : P1_<région> pour la partie 1, P2 pour la partie 2, INR_P1/INR_P2
pour les non-résidents). Les chemins sont relatifs au dossier du manifest.

Usage :
    python3 02_parse_pdf_structure.py data/2024/manifest.json -o data/2024/pdf_struct.json
    python3 02_parse_pdf_structure.py --annee 2024   # manifest/-o résolus via config.json

Sortie JSON : { "<clé_document>": { "<code4chiffres>": {
    "check": "<chiffre de contrôle>", "page": <int>,
    "cadre": "...", "section": "...", "rubrique": "...", "label": "..."
} } }
"""
import zipfile, re, json, argparse, os, subprocess, shutil
from _layout import paths_for_year, add_annee_arg

CODE = re.compile(r'(?<!\d)(\d{4})-(\d{2})(?!\d)')
CADRE = re.compile(r'^\s*(cadre|kader)\s+[IVXLC0-9]+\s*[-–]', re.I)
SECTION = re.compile(r'^\s*([A-Z])\.\s+[A-ZÉÈÀÔÎ]')
RUBRIC = re.compile(r'^\s*(\d{1,2})\.\s+\S')
SUBRUBRIC = re.compile(r'^\s*([a-z])\)\s+\S')
PAGE_MARK = re.compile(r'^=== PAGE (\d+) ===\s*$')

def clean_label(txt):
    txt = txt.replace('□', '').replace('☐', '')
    txt = re.sub(r'\.{3,}', ' ', txt)
    txt = re.sub(r'\([0-9]{3}\)', '', txt)
    txt = re.sub(r'\s+', ' ', txt).strip(' .:-–…')
    return txt.strip()

def strip_leading_marker(txt):
    return re.sub(r'^\s*(\d{1,2}\.|\d{1,2}\)|[a-z]\)|-|•)\s*', '', txt)

def load_pages(path):
    """Retourne une liste [(page_num, texte), ...] quel que soit le format d'entrée."""
    if path.lower().endswith('.txt'):
        raw = open(path, encoding='utf-8', errors='replace').read()
        pages, cur_num, buf = [], None, []
        for line in raw.splitlines():
            m = PAGE_MARK.match(line)
            if m:
                if cur_num is not None:
                    pages.append((cur_num, '\n'.join(buf)))
                cur_num, buf = int(m.group(1)), []
            else:
                buf.append(line)
        if cur_num is not None:
            pages.append((cur_num, '\n'.join(buf)))
        if not pages:  # pas de marqueurs -> tout sur une "page" 1
            pages = [(1, raw)]
        return pages
    # zip (archive "pdf" d'origine) ou vrai pdf
    try:
        z = zipfile.ZipFile(path)
        txts = sorted(
            [n for n in z.namelist() if n.endswith('.txt') and n != 'manifest.json'],
            key=lambda n: int(''.join(filter(str.isdigit, n.split('/')[-1])) or 0)
        )
        return [(int(''.join(filter(str.isdigit, t.split('/')[-1]))),
                  z.read(t).decode('utf-8', errors='replace')) for t in txts]
    except zipfile.BadZipFile:
        if shutil.which('pdftotext') is None:
            raise RuntimeError(f"{path}: vrai PDF mais pdftotext introuvable (apt/brew install poppler-utils)")
        r = subprocess.run(['pdftotext', '-layout', path, '-'], capture_output=True, text=True, check=True)
        chunks = r.stdout.split('\f')
        if chunks and chunks[-1] == '':
            chunks = chunks[:-1]
        return [(p, txt) for p, txt in enumerate(chunks, start=1)]

def parse_document(path):
    out = {}
    cadre = section = rubric = ''
    rubric_open = False
    pending = []
    for page, text in load_pages(path):
        for line in text.splitlines():
            L = line.strip()
            if not L or L.lower().startswith('page '):
                continue
            has_code = CODE.search(L)
            if not has_code:
                if CADRE.match(L):
                    cadre = clean_label(L); section = ''; rubric = ''; rubric_open = False; pending = []
                    continue
                if SECTION.match(L) and len(L) < 90:
                    section = clean_label(L); rubric = ''; rubric_open = False; pending = []
                    continue
                if RUBRIC.match(L):
                    rubric = clean_label(L); pending = [clean_label(L)]
                    rubric_open = not rubric.rstrip().endswith(':')
                    continue
                if SUBRUBRIC.match(L):
                    rubric_open = False
                if cadre and L.isupper() and not section and not rubric and len(L) < 90:
                    cadre = clean_label(cadre + ' ' + L)
                    continue
                if rubric_open and not SUBRUBRIC.match(L) and L[:1].islower():
                    rubric = clean_label(rubric + ' ' + L)
                    if rubric.rstrip().endswith(':'):
                        rubric_open = False
                    continue
                rubric_open = False
                pending.append(clean_label(strip_leading_marker(L)))
                pending = pending[-4:]
                continue
            rubric_open = False
            if RUBRIC.match(L):
                rubric = clean_label(CODE.sub('', L))
            codes = list(CODE.finditer(L))
            first = codes[0]
            pre = clean_label(strip_leading_marker(L[:first.start()]))
            if pre and re.search(r'[A-Za-zÀ-ÿ]', pre):
                label = pre
            else:
                tail = [p for p in pending if p and re.search(r'[A-Za-zÀ-ÿ]', p)]
                label = tail[-1] if tail else rubric
            for m in codes:
                d4, chk = m.group(1), m.group(2)
                if d4 not in out:
                    out[d4] = {'check': chk, 'page': page, 'cadre': cadre,
                               'section': section, 'rubrique': rubric, 'label': label}
            pending = []
    return out

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('manifest', nargs='?', help='chemin vers manifest.json de l\'année à traiter (optionnel si --annee est fourni)')
    ap.add_argument('-o', '--out', help='chemin du fichier JSON de sortie (optionnel si --annee est fourni)')
    add_annee_arg(ap)
    args = ap.parse_args()

    if not args.manifest or not args.out:
        if args.annee is None:
            ap.error("manifest et -o/--out sont requis quand --annee n'est pas fourni")
        paths = paths_for_year(args.annee, args.config)
        args.manifest = args.manifest or paths['manifest']
        args.out = args.out or paths['pdf_struct']

    manifest = json.load(open(args.manifest, encoding='utf-8'))
    base = os.path.dirname(os.path.abspath(args.manifest))
    result = {}
    for key, relpath in manifest['documents'].items():
        path = os.path.join(base, relpath)
        if not os.path.exists(path):
            print(f'  [absent] {key}: {path} introuvable — ignoré')
            continue
        result[key] = parse_document(path)
        print(f'  {key}: {len(result[key])} codes structurés ({path})')

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    json.dump(result, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f'Écrit: {args.out}')

if __name__ == '__main__':
    main()

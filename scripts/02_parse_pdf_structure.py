#!/usr/bin/env python3
"""
Rebuilds the Frame -> Section -> Item -> Label structure from the OCR text of the
IPCAL preparatory documents, and attaches that structure to every declaration code
(format NNNN-CC) found in the text.

Note on vocabulary: the source documents are in French/Dutch, so the structural
levels keep their domain meaning -- "Frame" is the form's *Cadre* (Cadre I, II,
III...), "Item" is the numbered *Rubrique* inside a section.

Input: a JSON manifest describing, for one year, the documents to process. Each
document may be either:
  - a .txt produced by 01_extract_pdf_text.py (with "=== PAGE N ===" markers)
  - the original zip "pdf" archive (OCR text extracted on the fly)
  - a real PDF (pdftotext fallback, same as 01_extract_pdf_text.py)

Example manifest (data/2024/manifest.json):
{
  "income_year": 2024,
  "assessment_year": 2025,
  "documents": {
    "P1_BXL":  "raw/P1_BXL.txt",
    "P1_RF":   "raw/P1_RF.txt",
    "P1_RW":   "raw/P1_RW.txt",
    "P2":      "raw/P2.txt",
    "INR_P1":  "raw/INR_P1.txt",
    "INR_P2":  "raw/INR_P2.txt"
  }
}
The "documents" keys are free-form (P1_BXL, P1_RF, ... is simply the convention used
so far: P1_<region> for part 1, P2 for part 2, INR_P1/INR_P2 for non-residents).
Paths are relative to the manifest's folder.

Usage:
    python3 02_parse_pdf_structure.py data/2024/manifest.json -o data/2024/pdf_struct.json
    python3 02_parse_pdf_structure.py --year 2024   # manifest/-o resolved via config.json

JSON output: { "<document_key>": { "<4digit_code>": {
    "check": "<check digits>", "page": <int>,
    "frame": "...", "section": "...", "item": "...", "label": "..."
} } }
"""
import zipfile, re, json, argparse, os, subprocess, shutil
from _layout import paths_for_year, add_year_arg

CODE = re.compile(r'(?<!\d)(\d{4})-(\d{2})(?!\d)')
FRAME = re.compile(r'^\s*(cadre|kader)\s+[IVXLC0-9]+\s*[-–]', re.I)
SECTION = re.compile(r'^\s*([A-Z])\.\s+[A-ZÉÈÀÔÎ]')
ITEM = re.compile(r'^\s*(\d{1,2})\.\s+\S')
SUBITEM = re.compile(r'^\s*([a-z])\)\s+\S')
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
    """Return [(page_number, text), ...] whatever the input format."""
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
        if not pages:  # no markers -> everything on a single "page" 1
            pages = [(1, raw)]
        return pages
    # zip (original "pdf" archive) or real pdf
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
            raise RuntimeError(f"{path}: real PDF but pdftotext not found (apt/brew install poppler-utils)")
        r = subprocess.run(['pdftotext', '-layout', path, '-'], capture_output=True, text=True, check=True)
        chunks = r.stdout.split('\f')
        if chunks and chunks[-1] == '':
            chunks = chunks[:-1]
        return [(p, txt) for p, txt in enumerate(chunks, start=1)]


def parse_document(path):
    out = {}
    frame = section = item = ''
    item_open = False
    pending = []
    for page, text in load_pages(path):
        for line in text.splitlines():
            L = line.strip()
            if not L or L.lower().startswith('page '):
                continue
            has_code = CODE.search(L)
            if not has_code:
                if FRAME.match(L):
                    frame = clean_label(L); section = ''; item = ''; item_open = False; pending = []
                    continue
                if SECTION.match(L) and len(L) < 90:
                    section = clean_label(L); item = ''; item_open = False; pending = []
                    continue
                if ITEM.match(L):
                    item = clean_label(L); pending = [clean_label(L)]
                    item_open = not item.rstrip().endswith(':')
                    continue
                if SUBITEM.match(L):
                    item_open = False
                if frame and L.isupper() and not section and not item and len(L) < 90:
                    frame = clean_label(frame + ' ' + L)
                    continue
                if item_open and not SUBITEM.match(L) and L[:1].islower():
                    item = clean_label(item + ' ' + L)
                    if item.rstrip().endswith(':'):
                        item_open = False
                    continue
                item_open = False
                pending.append(clean_label(strip_leading_marker(L)))
                pending = pending[-4:]
                continue
            item_open = False
            if ITEM.match(L):
                item = clean_label(CODE.sub('', L))
            codes = list(CODE.finditer(L))
            first = codes[0]
            pre = clean_label(strip_leading_marker(L[:first.start()]))
            if pre and re.search(r'[A-Za-zÀ-ÿ]', pre):
                label = pre
            else:
                tail = [p for p in pending if p and re.search(r'[A-Za-zÀ-ÿ]', p)]
                label = tail[-1] if tail else item
            for m in codes:
                d4, chk = m.group(1), m.group(2)
                if d4 not in out:
                    out[d4] = {'check': chk, 'page': page, 'frame': frame,
                               'section': section, 'item': item, 'label': label}
            pending = []
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('manifest', nargs='?', help="path to the year's manifest.json (optional if --year is given)")
    ap.add_argument('-o', '--out', help='path of the output JSON file (optional if --year is given)')
    add_year_arg(ap)
    args = ap.parse_args()

    if not args.manifest or not args.out:
        if args.year is None:
            ap.error("manifest and -o/--out are required when --year is not given")
        paths = paths_for_year(args.year, args.config)
        args.manifest = args.manifest or paths['manifest']
        args.out = args.out or paths['pdf_struct']

    manifest = json.load(open(args.manifest, encoding='utf-8'))
    base = os.path.dirname(os.path.abspath(args.manifest))
    result = {}
    for key, relpath in manifest['documents'].items():
        path = os.path.join(base, relpath)
        if not os.path.exists(path):
            print(f'  [missing] {key}: {path} not found - skipped')
            continue
        result[key] = parse_document(path)
        print(f'  {key}: {len(result[key])} codes structured ({path})')

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    json.dump(result, open(args.out, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f'Written: {args.out}')


if __name__ == '__main__':
    main()

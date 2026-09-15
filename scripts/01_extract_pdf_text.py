#!/usr/bin/env python3
"""
Allège les archives "PDF" IPCAL (en réalité des ZIP page1.jpeg+page1.txt+...+manifest.json)
en ne conservant que le texte OCR, page par page.

Usage :
    python3 extract_pdf_text.py fichier1.pdf [fichier2.pdf ...]
    python3 extract_pdf_text.py --dir /chemin/vers/dossier   # tous les .pdf du dossier

Produit un .txt à côté de chaque fichier source (même nom, extension .txt),
~40x plus léger, avec un marqueur "=== PAGE N ===" avant chaque page
(nécessaire pour que le parseur retrouve la structure Cadre/Section/Rubrique).
"""
import zipfile, sys, os, glob, subprocess, shutil

def extract_from_zip(path, out_path):
    z = zipfile.ZipFile(path)
    txts = sorted(
        [n for n in z.namelist() if n.endswith('.txt') and n != 'manifest.json'],
        key=lambda n: int(''.join(filter(str.isdigit, n.split('/')[-1])) or 0)
    )
    if not txts:
        print(f'  [ignoré] {path} : aucun fichier .txt trouvé dans l\'archive')
        return False
    with open(out_path, 'w', encoding='utf-8') as out:
        for t in txts:
            page_num = ''.join(filter(str.isdigit, t.split('/')[-1]))
            content = z.read(t).decode('utf-8', errors='replace')
            out.write(f'=== PAGE {page_num} ===\n')
            out.write(content)
            out.write('\n')
    return True

def extract_from_real_pdf(path, out_path):
    """Vrai PDF (pas d'archive) : pdftotext -layout, pages séparées par \\f
    (saut de page inséré par pdftotext par défaut) — pas besoin de pdfinfo."""
    if shutil.which('pdftotext') is None:
        print(f'  [ignoré] {path} : pdftotext introuvable (installer poppler-utils)')
        return False
    r = subprocess.run(['pdftotext', '-layout', path, '-'], capture_output=True, text=True)
    if r.returncode != 0:
        print(f'  [erreur] {path} : pdftotext a échoué : {r.stderr.strip()[:200]}')
        return False
    pages = r.stdout.split('\f')
    if pages and pages[-1] == '':
        pages = pages[:-1]
    with open(out_path, 'w', encoding='utf-8') as out:
        for i, content in enumerate(pages, start=1):
            out.write(f'=== PAGE {i} ===\n')
            out.write(content)
            out.write('\n')
    return True

def extract_one(path):
    out_path = os.path.splitext(path)[0] + '.txt'
    try:
        ok = extract_from_zip(path, out_path)
    except zipfile.BadZipFile:
        ok = extract_from_real_pdf(path, out_path)
    if not ok:
        return
    before = os.path.getsize(path)
    after = os.path.getsize(out_path)
    print(f'  {os.path.basename(path)}: {before/1e6:.2f} MB -> {os.path.basename(out_path)}: {after/1e3:.1f} KB  (x{before/max(after,1):.0f} plus léger)')

def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(1)
    if args[0] == '--dir':
        files = sorted(glob.glob(os.path.join(args[1], '*.pdf')))
    else:
        files = args
    print(f'{len(files)} fichier(s) à traiter :')
    for f in files:
        extract_one(f)

if __name__ == '__main__':
    main()

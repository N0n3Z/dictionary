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
import zipfile, sys, os, glob

def extract_one(path):
    out_path = os.path.splitext(path)[0] + '.txt'
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        print(f'  [ignoré] {path} : pas une archive ZIP (peut-être déjà un vrai PDF)')
        return
    txts = sorted(
        [n for n in z.namelist() if n.endswith('.txt') and n != 'manifest.json'],
        key=lambda n: int(''.join(filter(str.isdigit, n.split('/')[-1])) or 0)
    )
    if not txts:
        print(f'  [ignoré] {path} : aucun fichier .txt trouvé dans l\'archive')
        return
    with open(out_path, 'w', encoding='utf-8') as out:
        for t in txts:
            page_num = ''.join(filter(str.isdigit, t.split('/')[-1]))
            content = z.read(t).decode('utf-8', errors='replace')
            out.write(f'=== PAGE {page_num} ===\n')
            out.write(content)
            out.write('\n')
    before = os.path.getsize(path)
    after = os.path.getsize(out_path)
    print(f'  {os.path.basename(path)}: {before/1e6:.2f} MB -> {os.path.basename(out_path)}: {after/1e3:.1f} KB  (x{before/after:.0f} plus léger)')

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

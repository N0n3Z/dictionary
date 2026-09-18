#!/usr/bin/env python3
"""
Strips IPCAL "PDF" files down to their OCR text, page by page.

Two input formats are handled, detected by content rather than extension:
  - real PDFs (extracted with pdftotext -layout)
  - ZIP archives disguised as .pdf (page1.jpeg + page1.txt + ... + manifest.json),
    which is how the administration shipped some years

Usage:
    python3 01_extract_pdf_text.py file1.pdf [file2.pdf ...]
    python3 01_extract_pdf_text.py --dir /path/to/folder   # every .pdf in the folder
    python3 01_extract_pdf_text.py --year 2023             # resolves raw_dir via config.json

Writes a .txt next to each source file (same name, .txt extension), roughly 5-40x
smaller, with a "=== PAGE N ===" marker before each page (required for the parser to
recover the Frame/Section/Item structure).
"""
import zipfile, os, glob, argparse, subprocess, shutil
from _layout import paths_for_year, add_year_arg


def extract_from_zip(path, out_path):
    z = zipfile.ZipFile(path)
    txts = sorted(
        [n for n in z.namelist() if n.endswith('.txt') and n != 'manifest.json'],
        key=lambda n: int(''.join(filter(str.isdigit, n.split('/')[-1])) or 0)
    )
    if not txts:
        print(f'  [skipped] {path}: no .txt file inside the archive')
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
    """Real PDF (not an archive): pdftotext -layout, pages separated by \\f
    (form feed inserted by pdftotext by default) -- no need for pdfinfo."""
    if shutil.which('pdftotext') is None:
        print(f'  [skipped] {path}: pdftotext not found (install poppler-utils)')
        return False
    r = subprocess.run(['pdftotext', '-layout', path, '-'], capture_output=True, text=True)
    if r.returncode != 0:
        print(f'  [error] {path}: pdftotext failed: {r.stderr.strip()[:200]}')
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
    print(f'  {os.path.basename(path)}: {before/1e6:.2f} MB -> {os.path.basename(out_path)}: {after/1e3:.1f} KB  ({before/max(after,1):.0f}x smaller)')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('files', nargs='*', help='PDF files to process (alternative to --dir / --year)')
    ap.add_argument('--dir', help='process every .pdf in this folder')
    add_year_arg(ap)
    args = ap.parse_args()

    if args.dir:
        files = sorted(glob.glob(os.path.join(args.dir, '*.pdf')))
    elif args.year is not None:
        raw_dir = paths_for_year(args.year, args.config)['raw_dir']
        files = sorted(glob.glob(os.path.join(raw_dir, '*.pdf')))
    elif args.files:
        files = args.files
    else:
        ap.error('provide files, --dir <folder>, or --year <year>')

    if not files:
        print('No .pdf file found.')
        return
    print(f'{len(files)} file(s) to process:')
    for f in files:
        extract_one(f)


if __name__ == '__main__':
    main()

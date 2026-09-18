# IPCAL data dictionary

Builds a multi-year data dictionary of the Belgian personal income tax declaration
(IPCAL micro data) from the administration's preparatory PDFs and annual master Excel
files.

**[→ Full usage guide: `GUIDE.md`](GUIDE.md)** — prerequisites, the seven steps,
handling years that deviate, and how to read the outputs.
`CLAUDE.md` holds the business rules and project context.

> Code, CLI options, config keys and output column names are in English; data values
> (labels, hierarchy levels) keep the administration's original French/Dutch wording.

**Current coverage:** income years 2017–2023 · 43 703 dictionary rows · 6 982 codes

## Quick start (a new year)

```bash
# 1. Drop the raw files (PDF + Excel) into data/<year>/raw/, renamed to match the
#    manifest keys: IPCAL_2024.xlsx, P1_BXL.pdf, P1_RF.pdf, P1_RW.pdf, P2.pdf, ...
mkdir -p data/2024/raw

# 2. Check what is missing before going further (also useful routinely, across the
#    whole series, to spot a collection gap such as "Brussels 2016")
python3 scripts/00_check_sources.py

# 3. Write data/2024/manifest.json (see GUIDE.md), then run the pipeline
python3 scripts/01_extract_pdf_text.py --year 2024
python3 scripts/02_parse_pdf_structure.py --year 2024
python3 scripts/03_build_dictionary.py --year 2024

# 4. Write the workbooks (one year, or the whole series via glob)
python3 scripts/04_write_excel.py data/2024/records.json -o data/2024/IPCAL_data_dictionary_2024.xlsx
python3 scripts/04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_full.xlsx

# 5. (multi-year) Per-variable history + semantic drift detection
python3 scripts/05_build_history.py data/*/records.json -o data/history.json
python3 scripts/06_write_history_excel.py data/history.json -o IPCAL_history_variables.xlsx
```

`--year <year>` (steps 1–3) resolves paths automatically via `config.json` at the
repository root — handy for a local run, with no code change. Any explicit path option
(`--excel`, `--struct`, `-o`, ...) always takes precedence. If your local folder layout
differs (no `raw/` subfolder, data stored outside `data/`, ...), edit `config.json`
rather than the scripts — see the header comment in `scripts/_layout.py`.

Requires Python 3.9+, `openpyxl`, and `pdftotext` (poppler-utils).

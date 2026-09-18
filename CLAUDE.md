# IPCAL — data dictionary of the Belgian personal income tax declaration

Language convention (user decision, 2026-09-17): **code, CLI options, config keys and
output column names are in English**; **data values** (labels, hierarchy levels) keep
the administration's original French/Dutch wording, because they are source data, not
translated text. Do not "translate" label content when touching this project.

## Project goal
Produce a structured, multi-year data dictionary for IPCAL micro data (Belgian
personal income tax declaration). Income years 2017-2023 processed and validated
(7 years, see `data/`); the goal now is to extend to the full historical series
(2014 → today, with a special case before 2014 where a single PDF existed rather than
one per region) as the missing raw files (2014-2016, and the income-2024 /
assessment-2025 test year validated manually but not yet re-integrated here) are
dropped into `data/<year>/raw/`.

## Sources (two per year)
1. **Preparatory PDF**: one per region (BXL/RF/RW) × part (1=standard,
   2=self-employed), plus two "non-resident" versions (INR, simplified, parts 1 and 2).
   Before 2014: a single PDF (no regional split).
   - **Trap**: some of these "PDFs" are actually **ZIP archives in disguise** (check
     the magic bytes `PK\x03\x04`, not the extension). They contain `N.jpeg` (page N
     image) + `N.txt` (page N OCR text) + `manifest.json`.
   - OCR text is only ~2.5% of the total weight (the rest is page images).
     `scripts/01_extract_pdf_text.py` extracts the text only (5-40x smaller), with a
     `pdftotext -layout` path for real PDFs (which is what 2017-2023 turned out to be).
2. **Annual master Excel** (`IPCAL_YYYY.xlsx`): the exhaustive list of IPCAL codes,
   including the **computed/administrative** variables absent from the PDFs.
   - Useful sheet: the one holding the codes (`IPCAL_Codes` usually, `Feuil1` in 2021).
     There is often a second, near-empty sheet — check before ignoring it outright, but
     do not assume it is always empty.
   - 3-row header, data from row 4 (**recheck every year**: print the first 5 rows
     before running the pipeline).
   - Key columns (0-based indices observed 2017-2023, **may shift between years — always
     verify**): 2=IPCAL_A(N-1), 4=IPCAL_B(N-1), 5=Decl_A(N), 6=IPCAL_A(N), 7=Decl_B(N),
     8=IPCAL_B(N), 9=IPType, 14-17=Dutch hierarchy (4 levels), 18-21=French hierarchy.

## Coding rules (stable over time, a priori)
- Declaration code (4 digits) → IPCAL code: first digit replaced by a letter.
  `1→A, 2→B, 3→C, 4→D`. The target is always the codes starting with a letter.
- **A↔B** = federal declared (first/second spouse); **C↔D** = regional declared;
  **E/F, G/H, I/J, K/L** = computed/administrative variables (no declaration code),
  always in first-spouse/second-spouse pairs.
- Spouse convention: prefixes `ACEGIK` = spouse 1, `BDFHJL` = spouse 2.
- The check digits (last 2 digits of the declaration code, e.g. `1250-11`) are only
  readable from the PDF — the Excel file does not contain them.

## Core business rule: PDF wins over Excel
**Excel labels can be stale** (observed: "Marié en 2007" in the Excel while the 2025
PDF reads "mariés en 2024..."). Rule: for any code present in a PDF, the hierarchy
(Frame/Section/Item) and the label from the PDF take precedence over the Excel. The
Excel is only authoritative for codes appearing in no PDF (computed/administrative
variables, roughly half the codes).

PDF structure is rebuilt by regex over the OCR text (`02_parse_pdf_structure.py`):
`Cadre <roman numeral> - ...` → `<Letter>. SECTION IN CAPITALS` → `<n>. Item...` →
label (usually just before the code on the same line for amounts; before the `□`
checkboxes for indicators). Watch out for items spanning 2 lines (e.g. "...au lieu /
de travail") — handled by a continuation heuristic (next line starting lowercase,
without an a)/b)/• marker).

Vocabulary note: the output columns use `Frame` for the form's *Cadre* and `Item_PDF`
for the numbered *Rubrique*, to keep a 1:1 mapping with the paper form's structure.

## Output grain: one row per code, with a spouse link
User decision (important, do not revisit without explicit instruction):
**one row per IPCAL code** (required for downstream processing of the raw data), with:
- `Spouse` (1 or 2) to filter spouse 1 directly (main use case after the data has been
  individualised).
- `IPCAL_code_spouse` (+ `Declaration_code_spouse*`) to find the paired code without
  duplicating the row.
- `Has_spouse_counterpart` (boolean): does the variable genuinely exist for both
  spouses, or is it individual/shared (e.g. household totals)?

## Availability: boolean columns (no text to parse)
`Available_PDF`, `Available_PDF_resident`, `Available_PDF_INR`, `Available_Excel`,
`Present_IPP`, `Present_INR`, `Region_dependent`, `New_this_year`, plus one
`Present_<document_key>` column per source document (e.g. `Present_P1_BXL`,
`Present_P2`, `Present_INR_P1`...) generated dynamically from the manifest keys.

## IPType (Excel column, tax regimes)
90/91 = real-estate income (to be clarified — still open with the user); 92=Netherlands,
93=Germany, 94=Luxembourg (exempt foreign income, exemption with progression);
95=exempt CSSS; 96=French savings; 97=other PIT-exempt countries.

## Semantic drift across years (instrumented — `scripts/05_build_history.py` + `06_write_history_excel.py`)
A variable can change meaning over time under the same code. Empirical signal observed
by the user: an **availability gap** of ≥1 year, followed by a reappearance, often
indicates a change of semantics. Confirmed on 2017-2023: of 6982 multi-year codes, 76
have an availability gap, and all 76 correspond on review to a substantial label change
(max difflib similarity = 0.559 after normalisation — see `SUSPECT_THRESHOLD`).
- `05_build_history.py`: reads every `records.json`, groups by `IPCAL_code` (verified:
  no renumbering between years in this dataset — `IPCAL_code_previous == IPCAL_code` in
  100% of non-new cases — so the code serves as the stable variable identifier; revisit
  if a future year renumbers), detects runs of missing years within the known span,
  computes a label+hierarchy similarity before/after each gap, and **splits into
  "generations" the codes whose gap is judged suspect**.
  `python3 05_build_history.py data/*/records.json -o data/history.json`
- `06_write_history_excel.py`: 4 sheets — `Codes_with_break` (THE sheet to consult
  before any mapping: only the ~76 split codes, one row per generation with its
  validity range), `History` (one row per generation, label per year in columns, grey
  cells outside the generation's validity range, orange for a minor internal gap),
  `Semantic_breaks` (one row per boundary between generations, sorted by ascending
  similarity = most suspect first), `Legend`.
  `python3 06_write_history_excel.py data/history.json -o IPCAL_history_variables.xlsx`
- **Two-level architecture + no synthetic identifier** (user decision, 2026-09-17, do
  not revisit without explicit instruction):
  1. **primary table** = the annual `Dictionary` (03/04), one row per code **per year**,
     where `(IPCAL_code, Income_year)` already disambiguates fully;
  2. **derived aggregated view** = 05/06, one row per semantic code (a single one for
     ~99% of codes, several generations for codes with a semantic break).

  `IPCAL_code` is **never altered nor suffixed**, in any table. A generation is
  designated by a **two-column composite key** — `(IPCAL_code, Valid_from)` — and not
  by a concatenated string such as `A0270-2021`, which would look like an IPCAL code
  without being one (an early version had a `Cle_mapping` column of that kind; it was
  removed for this reason). Join to the annual data: `IPCAL_code` equal AND
  `Income_year` between `Valid_from` and `Valid_to`. The `Semantic_break` flag (True on
  every generation of a split code) must be checked before any mapping by `IPCAL_code`
  alone.
- Known limit: only drift accompanied by an availability gap is detected (consistent
  with the documented empirical signal); a change of meaning without a gap (code
  continuously present but whose meaning shifts) is not covered — such a case would
  remain a single generation, with `Semantic_break=False`, with no warning signal for a
  future mapping.

## Technical gotchas
- OCR control characters corrupt openpyxl writes: sanitize every string before writing
  with `re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', v)` (already done in
  `04_write_excel.py` and `06_write_history_excel.py`).
- Real Python booleans written straight into an openpyxl cell are preserved as Excel
  booleans — no workaround needed.
- `zipfile.ZipFile(path)` raises `BadZipFile` when the file is a real PDF (not an
  archive) → that is the signal used to fall back to `pdftotext`.
- `pdfinfo` is not available in this environment: page splitting relies on the form
  feed (`\f`) that `pdftotext` inserts by default, not on a page count.
- Windows console is cp1252: printing the `⚠` character from a script crashes unless
  `PYTHONIOENCODING=utf-8` is set. Affects ad-hoc inspection scripts, not the pipeline.

## Pipeline (scripts/, in order)
0. **`00_check_sources.py`** — raw source coverage check per year (PDF per
   region/part, non-residents, master Excel): present+declared, missing,
   declared-but-file-absent, present-but-undeclared, or not applicable (e.g. no
   regional split before 2014). Depends on no other step — run it as soon as a year is
   dropped in, before 01. `python3 00_check_sources.py [--start 2014] [--end <last
   year found>] -o IPCAL_source_coverage.xlsx`
1. **`01_extract_pdf_text.py`** — slims the "PDFs" down (zip→text, or real
   PDF→pdftotext). `python3 01_extract_pdf_text.py --year <year>`
2. **`02_parse_pdf_structure.py`** — rebuilds Frame/Section/Item/Label per code, from
   a per-year `manifest.json`.
   `python3 02_parse_pdf_structure.py --year <year>`
3. **`03_build_dictionary.py`** — merges master Excel + PDF structure, produces the
   records (one row per code) for one year.
   `python3 03_build_dictionary.py --year <year>`
4. **`04_write_excel.py`** — writes the final workbook (one year, or several via glob).
   `python3 04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_2017_2024.xlsx`
5. **`05_build_history.py`** — (multi-year only) per-variable history + semantic drift
   detection. `python3 05_build_history.py data/*/records.json -o data/history.json`
6. **`06_write_history_excel.py`** — Codes_with_break/History/Semantic_breaks/Legend
   workbook. `python3 06_write_history_excel.py data/history.json -o IPCAL_history_variables.xlsx`

A full usage guide for end users lives in `GUIDE.md` at the repository root.

## Folder convention for new years
```
data/
  2024/
    raw/                              <- raw PDF + Excel, as received, renamed:
      IPCAL_2024.xlsx                    master Excel
      P1_BXL.pdf / P1_RF.pdf / P1_RW.pdf part 1, residents, per region
      P2.pdf                              part 2, self-employed
      INR_P1.pdf / INR_P2.pdf            non-residents, parts 1/2
      P1.pdf                              before 2014: single PDF (no regional split)
    manifest.json                     <- key -> file mapping, see example below
    pdf_struct.json                   <- output of script 02
    records.json                      <- output of script 03
  2023/
    ...
```
Naming rules inside `raw/`: always `<key>.pdf` where `<key>` is exactly the key used in
`manifest.json` (no stray prefix such as `111-`, a single `.pdf` extension); master
Excel always `IPCAL_<income_year>.xlsx` (never the assessment year, to avoid the
"2023-2024" ambiguity — the assessment year is derived as `income_year + 1`).

Example `manifest.json` (the `excel_columns` block is optional — see next section):
```json
{
  "income_year": 2024,
  "assessment_year": 2025,
  "documents": {
    "P1_BXL": "raw/P1_BXL.txt",
    "P1_RF": "raw/P1_RF.txt",
    "P1_RW": "raw/P1_RW.txt",
    "P2": "raw/P2.txt",
    "INR_P1": "raw/INR_P1.txt",
    "INR_P2": "raw/INR_P2.txt"
  }
}
```
For years before 2014 (a single PDF, no regional split), adapt the manifest with a
single key (e.g. `"P1": "raw/P1.txt"`) — `03_build_dictionary.py` now reads the list
and order of keys from `manifest.json` itself (instead of a hard-coded list), so no
code change is needed for that case.

## Configuration for a local run (`config.json` + `manifest.json`)
Two levels of configuration, both optional (defaults = the layout above):
- **`config.json`** (repository root): location and layout of folders/files (templates
  with `{year}`/`{data_dir}`, resolved by `scripts/_layout.py`). Edit it if the local
  layout differs (no `raw/` subfolder, data outside the repository, another master
  Excel filename, etc.) — see the module's header comment for details. Scripts 01-03
  accept `--year <year>` (+ optional `--config <path>`) to resolve their paths from
  this config; any explicit path option (`--excel`, `--struct`, `-o`, the positional
  manifest...) takes precedence over what `--year` would derive.
- **per-year `manifest.json`**, optional `excel_columns` block: sheet, number of header
  rows, column indices (see `03_build_dictionary.py` docstring for the full schema),
  and PDF source labels. Use it when a year deviates from the defaults (e.g. 2021:
  sheet named `Feuil1` instead of `IPCAL_Codes` — see `data/2021/manifest.json`) rather
  than passing `--sheet`/`--header-rows` on the command line, so the adjustment stays
  versioned and reproducible across sessions.

## Open points / to confirm with the user before generalising broadly
- Confirm whether earlier years' archives share the same internal structure
  (page-by-page zip) or are real scanned PDFs without a text layer (in which case
  `pdftotext` will not suffice — OCR would be needed, not covered by the current script).
- Exact meaning of IPType regimes 90/91 (unclear).
- Region-dependency rule for computed variables: currently based on a keyword
  (gewest/régional/...) in the Excel hierarchy — imprecise, to be refined if a more
  reliable source exists.
- `Inferred_type`/`Inferred_unit` are heuristics over the labels, not truth from a
  structured source — to be validated or replaced if a better typing source exists.
- The PDF structure regexes have never been independently audited at scale against the
  source PDFs; items spanning two lines are the most fragile case.

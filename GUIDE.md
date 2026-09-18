# IPCAL pipeline — usage guide

Seven scripts that turn the administration's preparatory documents (PDF) and master
Excel file into a multi-year data dictionary of the Belgian personal income tax
declaration.

> Language note: code, CLI options, config keys and output column names are in
> English. The **data values** — labels, hierarchy levels — reproduce the
> administration's wording verbatim in French/Dutch, because they are source data,
> not translated text.

**Current coverage:** 2017–2023 · 43 703 dictionary rows · 6 982 distinct codes

---

## 1. What it produces

Three Excel deliverables, each answering a different question. All of them
regenerate entirely from the raw files — nothing is hand-edited.

| Deliverable | Grain | Purpose |
|---|---|---|
| **Annual dictionary** | one row per code per year | The primary table. Label, Frame/Section/Item hierarchy, availability per source document, paired spouse code. |
| **Variable history** | one row per semantic code | Derived aggregated view. Tracks each code over time and splits those whose meaning changed. Basis for a future mapping to national aggregates. |
| **Source coverage** | one row per year | Which documents were received, which are missing, which are present but forgotten in the manifest. |

`(IPCAL_code, Income_year)` is enough to disambiguate anything in the first one;
the second exists because a code whose meaning changed has to be disambiguated
*outside* any annual context. See [§6](#6-reading-the-outputs).

---

## 2. Prerequisites

| Dependency | Role | Install |
|---|---|---|
| Python 3.9+ | Running the scripts | python.org |
| `openpyxl` | Reading the master Excel, writing the workbooks | `pip install openpyxl` |
| `pdftotext` | Extracting text from the preparatory PDFs | `poppler-utils` (ships with Git for Windows) |

```bash
python -c "import openpyxl; print(openpyxl.__version__)"
pdftotext -v
```

If `pdftotext` is missing, step 1 reports it file by file instead of failing hard —
but no text gets extracted.

---

## 3. Adding a new year — the normal path

Example with income year 2024 (assessment year 2025):

```bash
# 1. Drop and rename the raw files
mkdir -p data/2024/raw
# -> data/2024/raw/IPCAL_2024.xlsx, P1_BXL.pdf, P1_RF.pdf, P1_RW.pdf, P2.pdf, ...

# 2. Write data/2024/manifest.json (template below)

# 3. Run the pipeline
python scripts/00_check_sources.py
python scripts/01_extract_pdf_text.py --year 2024
python scripts/02_parse_pdf_structure.py --year 2024
python scripts/03_build_dictionary.py --year 2024
python scripts/04_write_excel.py data/2024/records.json -o data/2024/IPCAL_data_dictionary_2024.xlsx

# 4. Regenerate the multi-year views
python scripts/04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_2017_2024.xlsx
python scripts/05_build_history.py data/*/records.json -o data/history.json
python scripts/06_write_history_excel.py data/history.json -o IPCAL_history_variables.xlsx
```

### File naming inside `raw/`

Each PDF is named exactly after its manifest key — no stray prefix, a single
extension. The master Excel is named after the **income year**, never the assessment
year, to avoid the "2023-2024" ambiguity.

| File | Content |
|---|---|
| `IPCAL_<year>.xlsx` | Master Excel from the administration |
| `P1_BXL.pdf` · `P1_RF.pdf` · `P1_RW.pdf` | Part 1, residents, per region |
| `P2.pdf` | Part 2 (self-employed) |
| `INR_P1.pdf` · `INR_P2.pdf` | Non-residents, parts 1 and 2 |
| `P1.pdf` | Before 2014: single PDF, no regional split |

### The year's manifest

`data/<year>/manifest.json` declares which documents exist. The keys listed here
drive everything else: PDF source precedence and the `Present_<key>` columns of the
dictionary.

```json
{
  "income_year": 2024,
  "assessment_year": 2025,
  "documents": {
    "P1_BXL": "raw/P1_BXL.txt",
    "P1_RF":  "raw/P1_RF.txt",
    "P1_RW":  "raw/P1_RW.txt",
    "P2":     "raw/P2.txt",
    "INR_P1": "raw/INR_P1.txt",
    "INR_P2": "raw/INR_P2.txt"
  }
}
```

**Declare only what exists.** A key absent from the manifest is treated as "document
not received", not as an error — a year without part 2 simply declares the three
`P1_*` keys. The paths point at the `.txt` files produced by step 1, not the original
PDFs.

---

## 4. The seven steps

Script numbers are the execution order. Steps 0–4 work one year at a time; steps 5
and 6 only make sense across several years.

### `00_check_sources.py` — coverage check

Checks, year by year, which documents are present on disk *and* declared in the
manifest. Depends on nothing else: run it as soon as a year is dropped in, and
routinely across the whole series.

```bash
python scripts/00_check_sources.py [--start 2014] [--end 2024]
```

States: `OK` · `MISSING` (never received) · `FILE_ABSENT` (declared, broken path) ·
`UNDECLARED` (received but forgotten in the manifest, therefore invisible to the
pipeline) · `N/A`.

### `01_extract_pdf_text.py` — slim the PDFs down

Writes a `.txt` next to each PDF, with a `=== PAGE N ===` marker per page. Handles
two formats: real PDFs (via `pdftotext`) and ZIP archives disguised as `.pdf` that
the administration shipped some years — detected by their bytes, not their
extension.

```bash
python scripts/01_extract_pdf_text.py --year 2024
python scripts/01_extract_pdf_text.py --dir data/2024/raw   # explicit equivalent
```

Typical reduction: 5× to 40× depending on the source format.

### `02_parse_pdf_structure.py` — rebuild the hierarchy

Reconstructs the `Frame → Section → Item → Label` structure by regex and attaches it
to every declaration code found in the text. Output: `pdf_struct.json`.

```bash
python scripts/02_parse_pdf_structure.py --year 2024
```

**Sanity check:** the script prints the number of codes structured per document.
Expect ~500 for a regional part 1, ~300 for a part 2. A much lower figure signals
poor OCR or a changed layout.

### `03_build_dictionary.py` — merge Excel and PDF

The business core. Cross-references the master Excel (exhaustive code list,
including computed variables absent from the PDFs) with the extracted structure, and
produces one row per IPCAL code.

```bash
python scripts/03_build_dictionary.py --year 2024
```

**Precedence rule:** for any code present in a PDF, **the PDF's label and hierarchy
win over the Excel file**, whose labels can be stale. The Excel is authoritative only
for computed or administrative codes absent from every PDF — roughly half the total.

### `04_write_excel.py` — write the dictionary

Assembles one or more `records.json` into a workbook (*Dictionary*, *Diagnostics*,
*Legend* sheets). Accepts a glob to produce the full series.

```bash
# one year
python scripts/04_write_excel.py data/2024/records.json -o data/2024/IPCAL_data_dictionary_2024.xlsx
# whole series
python scripts/04_write_excel.py data/*/records.json -o IPCAL_data_dictionary_2017_2024.xlsx
```

Writing the full series takes a few minutes (~44 000 styled rows).

### `05_build_history.py` — track variables over time

Groups codes across all loaded years, detects availability gaps and compares labels
on either side. Splits into "generations" the codes whose meaning changed.

```bash
python scripts/05_build_history.py data/*/records.json -o data/history.json
```

### `06_write_history_excel.py` — write the history view

Four sheets: *Codes_with_break* (check this one first), *History*,
*Semantic_breaks*, *Legend*.

```bash
python scripts/06_write_history_excel.py data/history.json -o IPCAL_history_variables.xlsx
```

---

## 5. When a year deviates

The administration's files are not perfectly stable year to year. Every deviation
encountered so far is fixed **in the manifest**, without touching the Python code.

### The `excel_columns` block

Optional, and only the keys that differ from the defaults need to appear. Real
example — in 2021 the useful sheet is called `Feuil1`, not `IPCAL_Codes`:

```json
{
  "income_year": 2021,
  "assessment_year": 2022,
  "documents": { "...": "..." },
  "excel_columns": {
    "sheet": "Feuil1"
  }
}
```

Full schema, with defaults (0-based column indices):

```jsonc
"excel_columns": {
  "sheet": "IPCAL_Codes",   // sheet holding the codes
  "header_rows": 3,         // header rows before the data
  "ipcal_a_prev": 2,        // IPCAL code, first spouse,  year N-1
  "ipcal_b_prev": 4,        // IPCAL code, second spouse, year N-1
  "decl_a": 5, "ipcal_a": 6,
  "decl_b": 7, "ipcal_b": 8,
  "iptype": 9,
  "hier_nl": [14, 15, 16, 17],   // Dutch hierarchy, 4 levels
  "hier_fr": [18, 19, 20, 21]    // French hierarchy, 4 levels
}
```

> **Reflex before every new year.** Open the master Excel and check the sheet name
> and the first five rows. Column indices held from 2017 to 2023, but nothing
> guarantees it — a silent shift would produce an entirely wrong dictionary without
> raising any error.
>
> ```bash
> python -c "import openpyxl;wb=openpyxl.load_workbook('data/2024/raw/IPCAL_2024.xlsx',read_only=True);print(wb.sheetnames)"
> ```

### Deviations already encountered

| Year | Deviation | Handling |
|---|---|---|
| 2017 | No PDF at all, Excel only | Manifest with `"documents": {}`; hierarchy 100 % from Excel |
| 2018 | No part 2 | Manifest limited to the three `P1_*` keys |
| 2021 | Sheet named `Feuil1` | `excel_columns.sheet` |
| 2023 | Only year with the non-resident PDFs | `INR_P1` / `INR_P2` keys added |

---

## 6. Reading the outputs

### Two levels, one join rule

The annual dictionary is the primary table; the history is its aggregated view.
**`IPCAL_code` is never modified nor suffixed** in either — a generation is
designated by a two-column composite key.

```
  PRIMARY TABLE                                   AGGREGATED VIEW
  (IPCAL_code, Income_year)   ──────────────▶     (IPCAL_code, Valid_from)
  Dictionary                   IPCAL_code equal    History
  one row per code per year    AND year within     one row per semantic code
                               [Valid_from, Valid_to]
```

Verified across the whole dataset: all 43 703 annual rows join to exactly one
generation — no orphans, no ambiguity.

### Columns worth knowing — annual dictionary

| Column | What it says |
|---|---|
| `IPCAL_code` | Variable identifier. Prefixes `A/C/E/G/I/K` = first spouse, `B/D/F/H/J/L` = second |
| `Spouse` | `1` or `2` — direct filter once the data is individualised |
| `IPCAL_code_spouse` | The paired code, without duplicating the row |
| `Has_spouse_counterpart` | Whether the variable really exists for both spouses, or is shared at household level |
| `Hierarchy_source` | `PDF` or `Excel` — where the retained label and hierarchy come from |
| `Frame` / `Section_PDF` / `Item_PDF` | Structural levels of the paper form (`Frame` = the form's *Cadre*, `Item` = the numbered *Rubrique*) |
| `Available_PDF_resident` · `Available_PDF_INR` | Availability booleans per scope |
| `Present_<key>` | One boolean per source document, generated from the manifest |
| `Region_dependent` | Regionalised variable (see limits) |

### Columns worth knowing — history

| Column | What it says |
|---|---|
| `Semantic_break` | **The decisive flag.** True on every generation of a split code |
| `Valid_from` · `Valid_to` | Year range of this generation — not of the whole code |
| `Generation` | Rank of the generation for that code |
| `Continuity_pct` | Share of the range's years where the code is actually present |
| `Minor_availability_gap` | Internal gap, but label stayed stable — not a change of meaning |

> ### ⚠ Before any mapping to a national aggregate
>
> Out of 6 982 codes, **76 changed meaning under the same number**. Example:
> `A0270` denotes "n'était ni marié, ni cohabitant légal" in 2017, but a
> child-care-expense condition from 2021 onwards. Mapping `A0270` to a single
> aggregate would be wrong over part of the period, *with no visible error*.
> Filter on `Semantic_break` in the *Codes_with_break* sheet before building the
> correspondence table.

---

## 7. Known limits

Four places where the dictionary relies on heuristics rather than an authoritative
source. None is blocking; all are worth knowing before using the outputs as an
official reference.

| Point | Nature of the limit |
|---|---|
| **Drift without a gap** | Only semantic drift *accompanied by an availability gap* is detected. A code continuously present that quietly changed meaning would stay a single generation, with no warning signal. |
| **PDF structure** | The hierarchy is rebuilt by regex over OCR text. Items spanning two lines are handled heuristically — the most fragile case, never independently audited at scale. |
| **`Region_dependent`** | For computed variables, inferred from a keyword (*gewest*, *régional*…) in the Excel hierarchy. Imprecise by construction. |
| **`Inferred_type` · `Inferred_unit`** | Guessed from the label, not taken from a structured source. |

### Source coverage

2014 to 2016 have not been collected yet. The non-resident PDFs are only available
for 2023 — for other years, `Available_PDF_INR` means "not verified", not "absent
from the INR scope".

---

## 8. Configurable folder layout

`config.json`, at the repository root, describes where files live. The defaults
reproduce the repository's layout; for a local run with a different organisation,
edit this file rather than the scripts.

```json
{
  "data_dir": "data",
  "layout": {
    "year_dir":       "{data_dir}/{year}",
    "raw_dir":        "{year_dir}/raw",
    "manifest":       "{year_dir}/manifest.json",
    "pdf_struct":     "{year_dir}/pdf_struct.json",
    "records":        "{year_dir}/records.json",
    "excel_master":   "{raw_dir}/IPCAL_{year}.xlsx",
    "workbook_year":  "{year_dir}/IPCAL_data_dictionary_{year}.xlsx"
  }
}
```

Templates accept `{year}`, `{data_dir}`, and any other `layout` entry — resolution is
multi-pass, so `{raw_dir}` may itself reference `{year_dir}`. This is what `--year`
uses to derive paths in steps 1 to 3.

**Path precedence.** An explicit option (`--excel`, `--struct`, `-o`, the positional
manifest) always wins over what `--year` would derive. An alternative `config.json`
is passed via `--config <path>` or the `IPCAL_CONFIG` environment variable.

---

Regenerating everything from scratch: `00` → `06` for each year, then the multi-year
views. All outputs are reproducible from the files in `data/<year>/raw/` — the
intermediates (`pdf_struct.json`, `records.json`, `history.json`) are not versioned.
Business rules in detail: `CLAUDE.md` at the repository root.

# PDF Workbench

[![CI](https://github.com/Iamakashgaur/pdf-workbench/actions/workflows/ci.yml/badge.svg)](https://github.com/Iamakashgaur/pdf-workbench/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.11%20%7C%203.13-blue)](https://www.python.org/)
[![Tests](https://img.shields.io/badge/tests-99%20passing-brightgreen)](test_pdf_to_excel.py)
[![License](https://img.shields.io/badge/license-MIT-lightgrey)](#license)

Turns supplier order reports from PDF into reconciled Excel workbooks — and tells you,
every single time, whether anything was left behind.

**Verified against six live supplier invoices: 100 of 100 line items extracted,
reconciling to the cent ($91,948.90).**

![The converter reconciling an extracted report against the invoice's own stated total](docs/demo.png)

<sub>The whole flow, top to bottom. It reads the invoice's own `8 Items - $19,824.50`
header line, sums what it actually extracted, and leads with the difference — `$0.00`
here — so you can see the file is complete before you download it. Rows 4 and 7 are
settings with **no certificate at all**, which the column parser handles as an empty
column rather than a failed match.</sub>

<sub>*Real screenshot of the running app. The report is synthetic — invented customers,
order numbers and certificate ids — because real supplier data cannot go in a public
repository. The arithmetic is genuine: the stated total is correct for those rows, and
the reconciliation was computed, not staged.*</sub>

---

## The problem I solved

A back-office team received supplier order reports as PDFs and nothing else. Every line
item — customer, stone, certificate, order number, amount, ship date — was **re-keyed into
Excel by hand**. Slow, and quietly error-prone in the way manual transcription always is:
you don't find the mistake, the mistake finds you, three weeks later, in a number someone
already invoiced against.

The obvious fix is a PDF-to-Excel converter. The obvious fix does not work here, for two
reasons.

**1. These reports have no table to find.** They carry no ruling lines, so grid-detection
engines return nothing. What is left is flattened text, and pattern-matching that text
breaks on every real-world case in the document:

| In the document | What breaks |
|---|---|
| A stone description wraps onto the next line | The continuation looks like a new row |
| A setting or mounting has **no certificate** | The field the pattern expects is simply absent |
| A product type that isn't `Diamond` | The pattern has never heard of it |
| A totals line — `95 Items - $94222.52` | Starts with digits, reads as a data row |

**2. The dangerous failure is the quiet one.** A converter that drops four rows out of
ninety-five produces a file that looks *exactly* like a correct one. It is worse than no
tool at all, because it is trusted. For financial records, "mostly extracted" is not a
partial success — it is a silent, undetectable corruption of the source of truth.

So the real problem was never "parse a PDF". It was: **produce a spreadsheet the team can
rely on without checking it against the PDF first.**

## How I solved it

**Read the geometry, not the text.** Every word in a PDF carries its physical position on
the page. A value's *column* is a fact the file records; its *shape* is only a guess. So
the parser assigns each word to whichever column it is physically printed under, and never
asks what the value looks like.

The subtlety is where a column actually ends. Header labels are centred over their columns
while the data beneath is left-aligned, so the midpoint between two header labels lands
*inside* real values. Instead, for each page the parser scans the gap between adjacent
header anchors and splits at the **widest empty vertical strip in the printed data** —
calibrating column boundaries from the document itself rather than hard-coded coordinates.

![Column boundaries calibrated from the widest empty strip rather than the midpoint between header labels](docs/slides/column-calibration.png)

Every case in the table above then stops being a special case:

- A wrapped cell continues **under its own column**, so it merges into the right field.
- A blank certificate is just an empty column — nothing has to account for a missing field.
- Any product type reads fine, because the type is never required to be a specific word.
- Footer lines carry no serial number, so they are excluded *structurally*, not by keyword.

**Then refuse to fail silently.** Each row is validated before it is kept — it must carry a
serial number, an amount, an order number and a shipping date. Rows that fail are excluded,
and every exclusion is reported in **four places**: the console, the log file, the workbook's
`_Metadata` sheet, and the batch summary. The web UI goes further and reconciles the
extracted row count and amount total against the invoice's own `N Items - $X` header line,
showing the difference at full metric scale before you download anything.

The result is that downloading the file confirms something already visibly true, instead of
being the moment you start hoping.

## What I'd point at in a code review

| | |
|---|---|
| **The core idea** | [`ReportColumnParser._boundaries`](pdf_to_excel.py) — column calibration from the widest empty strip. Roughly 30 lines, and the reason the whole thing works. |
| **Design under constraint** | Three extraction engines tried in order, two of them optional and often absent. Availability is import-gated; missing engines degrade honestly instead of crashing. |
| **A performance fix** | `_PDFHandles` opens each document once per run instead of once per page, using explicit sentinels — a zero-page PDF is falsy, so truthiness checks leaked handles. |
| **A measurement bug worth the comment** | `Tables Found` and `Rows Extracted` are deliberately separate. Conflating them displayed a 95-row report as "95 tables". |
| **Test strategy** | 99 tests. Synthetic PDFs via reportlab for the pipeline; hand-built word-position fixtures for the geometry parser, so column logic is tested without a PDF in the loop; and a set that asserts this README still matches the code it documents. |
| **CI** | Ubuntu + Windows × Python 3.11/3.13, deliberately green *without* Camelot or Tabula. LibreOffice **is** installed on every leg, so `--to-pdf` is driven against a real `soffice` rather than a mock — and the Windows leg exercises the install-location fallback, since the installer there does not put `soffice` on `PATH`. |

## Known limitations

Stated plainly, because a tool that oversells itself is the problem this one exists to fix.

- Report shapes are configurable ([see below](#report-layouts)), but only for the
  **column-geometry** reader. If a page's geometry cannot be read at all, the
  last-resort flat-text parser only knows the built-in supplier report; other
  layouts fall through to generic text extraction rather than being guessed at.
- An encrypted PDF needs its password supplied per run; there is no keychain
  integration, and by design no way to store it in `config.json`.
- Scanned pages need Tesseract installed. Without it they are skipped and reported,
  never silently dropped.
- `java_path` in `config.json` is currently unused; put Java on `PATH` for tabula.
- Document → PDF conversion stops there; it does not chain into extraction, and
  there is no PDF → Word/PowerPoint direction. It needs LibreOffice installed.
- A spreadsheet converted to PDF is rendered at its stored column widths, so
  wide cells are clipped — which is why an `.xlsx` round-tripped through PDF
  does not come back as the original table.

---

## Features

| Capability | Detail |
|---|---|
| **PDF types** | Native text, scanned (image), mixed content |
| **Table extraction** | Auto-tries pdfplumber → camelot → tabula (last two optional) |
| **Order reports** | Reads the report's own column layout: handles wrapped cells, blank certificates and any product type. [Details](#order-reports) |
| **Report layouts** | Extra report shapes added in `config.json`, columns and validation together — no code change. [Details](#report-layouts) |
| **OCR** | Tesseract-powered for scanned pages |
| **Excel output** | Styled headers, alternating rows, auto column widths, frozen panes |
| **Batch mode** | Entire folder with progress bar + summary report |
| **Multiple formats** | `.xlsx` or `.csv` export |
| **Documents → PDF** | `.docx`/`.xlsx`/`.pptx`/`.odt`/`.rtf`/`.html` into PDF via LibreOffice (`--to-pdf`, optional) |
| **Preview** | Print table preview before saving |
| **GUI** | Optional Streamlit web app |
| **Configurable** | All settings via `config.json` |

---

## Installation

On Windows, just double-click **`install.bat`** — it does everything below and reports anything missing.

### 1. Python dependencies

```bash
pip install -r requirements.txt
```

Everything in `requirements.txt` installs from PyPI with no system software. This is all you need for digital PDFs, scanned PDFs, the CLI, the web UI and the test suite.

### 2. Tesseract OCR — only for scanned PDFs

Digital PDFs convert fine without it; scanned pages are skipped with a warning.

| OS | Command |
|---|---|
| Windows | Download installer from [UB-Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) |
| macOS | `brew install tesseract` |
| Linux | `sudo apt-get install tesseract-ocr` |

After installing on Windows, add Tesseract to PATH **or** set `tesseract_path` in `config.json`:
```json
{
  "tesseract_path": "C:/Program Files/Tesseract-OCR/tesseract.exe"
}
```

### 3. LibreOffice — only for `--to-pdf`

Needed solely to convert documents *into* PDF. Everything else works without it.

| OS | Install |
|---|---|
| Windows | [libreoffice.org/download](https://www.libreoffice.org/download/download-libreoffice/) |
| macOS | `brew install --cask libreoffice` |
| Linux | `sudo apt-get install libreoffice` |

Put `soffice` on PATH, **or** set the path in `config.json`:
```json
{ "libreoffice_path": "C:/Program Files/LibreOffice/program/soffice.exe" }
```

### 4. Optional extraction engines

**Not required.** Camelot and Tabula are fallback table engines, tried only when pdfplumber finds no tables. Each needs system software that pip cannot install, which is why they are kept separate — a failure here must not block a working install.

```bash
pip install -r requirements-optional.txt
```

| Engine | Also needs |
|---|---|
| **Camelot** — ruled/grid tables | Ghostscript: [Windows](https://www.ghostscript.com/releases/gsdnld.html) · `brew install ghostscript` · `apt-get install ghostscript` |
| **Tabula** — Java-based fallback | Java 8+ from [java.com](https://www.java.com/en/download/) |

Check what is currently active at any time:

```bash
python pdf_to_excel.py --check-deps
```

---

## Usage

### Single file

```bash
# Basic conversion — output goes to ./output/
python pdf_to_excel.py invoice.pdf

# Specify output path
python pdf_to_excel.py report.pdf -o results/report_2024.xlsx

# Export as CSV instead
python pdf_to_excel.py data.pdf --format csv -o ./csv_output/

# Preview extracted tables before saving
python pdf_to_excel.py report.pdf --preview

# Force a specific extraction method
python pdf_to_excel.py tricky.pdf --method camelot

# OCR in French
python pdf_to_excel.py french_doc.pdf --ocr-lang fra
```

### Batch processing

```bash
# Process all PDFs in a folder
python pdf_to_excel.py -d ./pdfs/ -o ./output/

# Batch with preview
python pdf_to_excel.py -d ./invoices/ -o ./results/ --preview
```

### Encrypted PDFs

```bash
# Best: the password never enters shell history or a process listing
export PDF_WORKBENCH_PASSWORD='…'          # Windows: $env:PDF_WORKBENCH_PASSWORD='…'
python pdf_to_excel.py locked.pdf

# Explicit, if you must
python pdf_to_excel.py locked.pdf --password '…'
```

The CLI never prompts. `getpass` on Windows reads the console directly rather
than stdin, so there is no reliable way to tell whether anyone is present to
answer — a scheduled or piped run would sit waiting forever. It states what to
supply and exits instead; the web UI is the interactive surface and asks there.

**There is deliberately no `config.json` key for this.** That file is committed,
so a password in it would be published by the next `git push` — and the tool
warns if it finds one there rather than quietly using it.

The two failure modes are reported separately, because they are different
problems for whoever holds the file: *"Password required"* means none was given,
*"Password incorrect"* means the one given did not unlock it.

### Documents → PDF

The one direction that writes *into* PDF, for when a report arrives as `.docx` or
`.xlsx` and has to be a PDF before anything else can read it. Needs LibreOffice
([see below](#3-libreoffice--only-for---to-pdf)).

```bash
# Single document
python pdf_to_excel.py report.docx --to-pdf -o ./pdfs/

# A whole folder, recursively
python pdf_to_excel.py -d ./documents/ --to-pdf -o ./pdfs/
```

Handles `.doc` `.docx` `.odt` `.rtf` `.txt` · `.xls` `.xlsx` `.ods` `.csv` ·
`.ppt` `.pptx` `.odp` · `.html` `.htm`. Anything else is refused by name before
LibreOffice is launched, rather than failing obscurely inside it.

> Conversion is only reported as successful if the PDF is actually on disk.
> LibreOffice exits `0` in cases where it wrote nothing at all, so the exit code
> is not trusted on its own.

The web UI does the same thing, and **there is no mode to switch**: drop a PDF and
it extracts, drop a document and it converts. Which one applies is a fact already
on disk, so the interface reads it rather than asking.

![The web UI converting a Word document into PDF](docs/demo-topdf.png)

<sub>If LibreOffice is not installed, the UI says so on the document panel — before
you click anything — and explains how to fix it, rather than failing after the fact.</sub>

### GUI (Streamlit)

The web UI lives in `app.py`, not in `pdf_to_excel.py`:

```bash
streamlit run app.py
```

On Windows, just double-click `convert.bat` — it locates Python, installs Streamlit if it is missing, starts the server and opens the browser for you.

Then open `http://localhost:8501`.

> **Note:** the `--gui` flag on `pdf_to_excel.py` does *not* launch the UI. It prints the command above and exits.

### Check dependencies

```bash
python pdf_to_excel.py --check-deps
```

---

## CLI Reference

```
usage: pdf_to_excel [-h] [-d DIR] [-o OUTPUT] [--format {excel,csv}]
                    [--config CONFIG] [--preview] [--verbose]
                    [--method {auto,pdfplumber,camelot,tabula}]
                    [--ocr-lang OCR_LANG] [--no-meta] [--keep-empty]
                    [--gui] [--check-deps]
                    [input]

Arguments:
  input                 PDF file path (omit when using -d for batch mode)

Options:
  -d DIR, --dir DIR     Input directory (batch mode)
  -o OUTPUT, --output OUTPUT
                        Output .xlsx file or output directory
  --format {excel,csv}  Output format (default: excel)
  --config CONFIG       Path to config.json (default: config.json)
  --preview             Print table preview to console before saving
  --verbose             Enable verbose debug logging
  --method METHOD       Force extraction method: auto|pdfplumber|camelot|tabula
  --ocr-lang LANG       Tesseract language code (default: eng)
  --no-meta             Omit metadata sheet from output
  --keep-empty          Keep empty rows and columns
  --password PASSWORD   Password for an encrypted PDF. Prefer the
                        PDF_WORKBENCH_PASSWORD environment variable
  --to-pdf              Convert documents INTO PDF via LibreOffice, instead of
                        reading a PDF. Works with -d for a folder
  --gui                 Print the GUI launch command and exit (does not start it)
  --check-deps          Print dependency status and exit
```

---

## Configuration (`config.json`)

All settings are optional. Missing keys use built-in defaults.

```json
{
  "output_dir": "output",
  "log_dir": "logs",

  "table_extraction_method": "auto",

  "ocr_language": "eng",
  "ocr_dpi": 300,
  "ocr_confidence_threshold": 60,

  "tesseract_path": null,
  "java_path": null,
  "libreoffice_path": null,

  "remove_empty_rows": true,
  "remove_empty_cols": true,

  "add_metadata_sheet": true,
  "freeze_header_row": true,
  "apply_table_style": true,

  "header_color": "1F4E79",
  "alternating_row_color": "D6E4F0",

  "max_col_width": 50,
  "min_col_width": 8,
  "max_sheet_name_length": 31,

  "batch_summary_report": true,
  "preview_rows": 10
}
```

| Key | Default | Description |
|---|---|---|
| `table_extraction_method` | `auto` | `auto` tries pdfplumber → camelot → tabula in order |
| `ocr_language` | `eng` | Tesseract language. `fra`=French, `deu`=German, `chi_sim`=Chinese |
| `ocr_dpi` | `300` | Render resolution for OCR. Higher = better quality, slower |
| `ocr_confidence_threshold` | `60` | Min OCR confidence (0–100) to include a word |
| `add_metadata_sheet` | `true` | Add a `_Metadata` sheet to the output (source info, counts, and the unparsed-row total) |
| `header_color` | `1F4E79` | Hex color for Excel header row (no `#`) |
| `alternating_row_color` | `D6E4F0` | Hex color for alternating data rows |
| `max_col_width` / `min_col_width` | `50` / `8` | Auto-sized column width bounds, in characters |
| `max_sheet_name_length` | `31` | Sheet-name truncation limit (Excel's own cap is 31) |
| `batch_summary_report` | `true` | Generate `_batch_summary_*.xlsx` after batch processing |
| `preview_rows` | `10` | Rows shown in the Streamlit preview table |
| `tesseract_path` | `null` | Full path to `tesseract.exe` if it is not on PATH |
| `java_path` | `null` | Reserved. Currently unused — put Java on PATH for tabula |
| `libreoffice_path` | `null` | Full path to `soffice` if it is not on PATH. Only used by `--to-pdf` |
| `report_layouts` | `null` | Extra report shapes to recognise. See [Report layouts](#report-layouts) |

---

## Report layouts

The column-geometry reader is not hard-wired to one report. A layout is **data**:
its columns, and the rules that make one of its rows valid. Add shapes in
`config.json` under `report_layouts` — no code change.

```json
{
  "report_layouts": [
    {
      "name": "parts_invoice",
      "columns": [
        {"name": "Line",        "tokens": ["Line"]},
        {"name": "Description", "tokens": ["Description"]},
        {"name": "Qty",         "tokens": ["Qty"]},
        {"name": "Unit Price",  "tokens": ["Unit", "Price"]},
        {"name": "Total",       "tokens": ["Total"]}
      ],
      "row_key": "Line",
      "structural_columns": ["Total"],
      "required": ["Description"],
      "patterns": {
        "Total": "^\\$?\\d[\\d,]*\\.\\d{2}$",
        "Qty": "^\\d+$"
      },
      "formats": {
        "Line": "integer", "Qty": "integer",
        "Unit Price": "number", "Total": "money"
      }
    }
  ]
}
```

| Key | Meaning |
|---|---|
| `columns` | Ordered. `tokens` are the header label split into **whole words**, exactly as they appear — `"Unit Price"` is `["Unit", "Price"]` |
| `row_key` | The column whose value marks a new record. Defaults to the first column |
| `row_key_pattern` | What that value must look like. Defaults to `^\d+$` |
| `required` | Columns that must be non-empty. Leave a column out if it is legitimately blank sometimes |
| `patterns` | Per-column regex a value must match. A column listed here is effectively required |
| `structural_columns` | Columns a wrapped cell never continues — an amount or a date. A "continuation" line carrying one is not a wrap, so it is reported instead of merged |
| `formats` | How each column is written to Excel. Drives both the cell format *and* the conversion, so the two cannot disagree |

**Formats.** A price read as `"4.50"` is the number `4.5`; without a declared
format Excel shows `4.5` — the right value, wrong for a price. Declaring it
fixes both halves at once: `"$4,722.30"` has its symbol and separators stripped
so the cell holds a real number, and the format puts them back on screen.

| Alias | Excel format | `"4.50"` → | `"$54.00"` → |
|---|---|---|---|
| `money` | `$#,##0.00` | `$4.50` | `$54.00` |
| `number` | `#,##0.00` | `4.50` | `54.00` |
| `integer` | `#,##0` | `5` | `54` |
| `percent` | `0.0%` | — | — |
| `text` | `@` | `4.50` verbatim | `$54.00` verbatim |

Use `text` for identifiers — it is what stops Excel renumbering an order
reference or reading `"00123"` as `123`. Anything not in this table is passed
through as a literal Excel format string, so `"\"EUR\" #,##0.00"` works too.

**How a layout is chosen.** The header test *is* the column test: a line qualifies
only if it carries every column label as whole words, consecutively and in order.
That is strict enough to identify the layout outright, so no layout needs a
separate "does this look like mine" rule.

**Two guarantees worth knowing.**

Configured layouts are tried **first**, and the built-in supplier report is always
appended **last** — so adding a layout can never stop the original from being read.

**Validation travels with the layout.** This is the point of the design, not a
detail: judging one report by another's required fields would let a new layout
parse perfectly and then have every row rejected and reported as excluded — this
tool's loudest alarm, firing at nothing. A malformed layout is reported and
skipped individually; the others, and the built-in, are unaffected.

The workbook's `_Metadata` sheet records which layout read the file.

The `_Metadata` sheet records the source file, page counts, and the **Unparsed Report Rows** total — see [Rows that don't match](#rows-that-dont-match). Set `add_metadata_sheet` to `false` to omit it.

---

## Output Structure

### Single PDF → Excel

```
output/
└── my_report_20240115.xlsx
    ├── P1_T1          ← Page 1, Table 1
    ├── P1_T2          ← Page 1, Table 2 (if exists)
    ├── P2_OCR         ← Page 2 (scanned, OCR result)
    ├── P3_Text        ← Page 3 (text, no clear tables)
    └── _Metadata      ← Source info, processing date, stats,
                          incl. Rows Extracted and Unparsed Report Rows
```

> **Reading the counts:** `Tables Found` is how many tabular datasets were written; `Rows Extracted` is how many data rows they contain. A 95-line report consolidated into one sheet reports **1 table, 95 rows**.

When every page yields the same columns, rows are combined into a **single sheet** instead of one sheet per page:

| Sheet | When |
|---|---|
| `Table_Data` | A table engine (pdfplumber/camelot/tabula) found the grid — the usual path for ruled invoice tables |
| `Text_Data` | No ruled grid was detected. Order reports land here and are read by the column parser |

### Single PDF → CSV

One CSV per dataset, named after the sheet the Excel path would have produced — so
`Text_Data` in a workbook and `report_Text_Data.csv` on disk are the same data.

```
csv_output/
├── report_Text_Data.csv     ← order report, read by the column parser
├── data_P1_T1.csv           ← page 1, table 1 (ruled grid)
└── data_P2_OCR.csv          ← page 2, scanned
```

CSV runs use the **same extraction cascade as Excel**, so they cover order reports and
scanned pages, and report excluded rows the same way. (They did not always: the CSV path
once called the table engines directly and silently produced nothing for a report with no
ruled grid.)

### Batch → Output folder

```
output/
├── invoice_20240115.xlsx
├── report_20240115.xlsx
├── contract_20240115.xlsx
└── _batch_summary_20240115_143022.xlsx   ← Success/fail, Rows Extracted
                                             and Unparsed Rows per file
```

Batch mode recurses into subdirectories, so two source PDFs can share a filename. Outputs
are **never overwritten within a run**: the second file to claim a name gets a `_2` suffix
and a `WARNING` naming both source paths. Re-running a batch still refreshes the previous
run's workbooks, which is the intended behaviour.

---

## How It Works

```
PDF Input
   │
   ├─ Classify (PyMuPDF)
   │     native text ──→ TableExtractor
   │     scanned      ──→ OCRProcessor → pytesseract
   │     mixed        ──→ both paths
   │
   ├─ Table Extraction (auto cascade):
   │     pdfplumber (layout-aware)
   │     camelot (grid + stream modes)      [optional]
   │     tabula (Java, robust fallback)     [optional]
   │
   ├─ Order report (when the header is present):
   │     ReportColumnParser reads word x-positions
   │     Column bounds calibrated per page from the layout
   │     Wrapped cells merged under their own column
   │     Each row validated before it is kept
   │       └─ falls back to text parsing if unreadable
   │
   ├─ OCR Pipeline (scanned pages):
   │     PyMuPDF renders page at 300 DPI
   │     Pillow enhances contrast/sharpness
   │     Tesseract extracts text + TSV confidence data
   │     Reconstruct table rows from positional data
   │
   └─ ExcelBuilder (openpyxl)
         Styled headers + alternating rows
         Auto column widths
         Metadata sheet
         Frozen panes
```

---

## Order Reports

Beyond the generic conversion above, this tool contains a **dedicated parser for the supplier order report**. It activates automatically — there is no flag — whenever a page contains this header line:

```
S. No. Customer Name Item Product Type Certificate No Order No Amount Manufacturer Est Shipping Date
```

### How rows are read: column geometry, not text patterns

The parser reads the report from the PDF's **physical column layout** rather than by pattern-matching the flattened text.

For each page it locates the header, then calibrates where each column begins by finding the widest empty vertical strip between adjacent header labels. (Header labels are centred over their columns while the data beneath is left-aligned, so a naive midpoint between labels lands inside real values.) Every word is then assigned to whichever column it is physically printed under.

This matters because a value's column is a fact recorded in the file, whereas its shape is a guess. Reading position directly handles cases that defeat pattern matching:

| Case | Handled by |
|---|---|
| **Wrapped cells** | A cell continuing on the next line is appended under its own column. `Bellamy Double Halo` + `Engagement Ring` → one `Item`; `IGI` + `LG_100000003` → one `Certificate No` |
| **Blank cells** | A setting or mounting has no certificate. The column is simply empty — no pattern has to account for a missing field |
| **Any product type** | `Diamond`, `Setting`, or anything new. Product type is read from its column, never required to be a specific word |
| **Varied continuation text** | `VVS2`, `Excellent`, `Round` all sit in the `Item` column and are handled identically. There is no vocabulary of expected words to keep up to date |
| **Footer lines** | Totals, item counts and page numbers carry no serial number, so they are excluded structurally rather than by keyword |

On top of that the tool **forces column types** so Excel cannot reformat `Order No` or `Certificate No` as numbers or dates, while `Amount` becomes real currency.

### Rows that don't match

Each row is validated before it is kept. It must carry the fields that make it a usable financial record:

- a numeric serial number
- an `Amount`
- an `Order No`
- a `DD Mon YYYY` shipping date

`Certificate No` and `Product Type` are deliberately **not** required — settings and mountings legitimately have neither.

**Rows failing validation are excluded from the output.** They are never silently dropped — every exclusion is reported in four places:

- a `WARNING` line in the console showing the offending row
- the full detail in the log file under `logs/`
- **`Unparsed Report Rows`** in the `_Metadata` sheet (requires `add_metadata_sheet: true`)
- an **`Unparsed Rows`** column in the batch summary workbook, highlighted amber when non-zero

If that count is not zero, **reconcile before using the data.**

### Reconciling a run

Each report states its own totals in the header (`25 Items - $5970.06`). Checking the extracted row count and the sum of `Amount` against that line is the fastest way to confirm a run is complete — it catches anything the row-level validation could not.

### Fallback

If a page's column layout cannot be read, the tool falls back to parsing the flattened text. That path is pattern-based and stricter — it expects the product type `Diamond` and a non-empty certificate — so a non-zero unparsed count on an unusual document may simply mean the fallback was used.

Skipped lines are logged at `DEBUG`, so run with `--verbose` and check `logs/` if you suspect a genuine row fragment was passed over.

---

## Troubleshooting

### "N report row(s) did not match the expected format"

Some rows of an order report were **left out of the Excel file**. Do not treat the output as complete until you have checked them.

1. Open the newest file in `logs/` and search for `Unparsed report row` — each excluded row is logged in full.
2. Check it against the four required fields in [Rows that don't match](#rows-that-dont-match): serial number, amount, order number, and a `DD Mon YYYY` date.
3. Reconcile the run against the report's own `N Items - $X` header line to see exactly what is missing.

Usual causes: a missing or reformatted amount or date, a serial number that did not extract cleanly, or a page whose column layout could not be read (in which case the stricter text fallback was used — see [Fallback](#fallback)).

A product type other than `Diamond`, or a blank certificate, is **not** a cause — those are handled normally by the column parser.

If the report's columns have genuinely been rearranged, `ReportColumnParser` needs updating — that is a code change, not a config one.

### No tables extracted from a PDF

1. Try forcing a different method: `--method camelot` or `--method tabula`
2. If it's a scanned PDF, ensure Tesseract is installed and on PATH
3. Run `--check-deps` to verify all libraries are available
4. Check `logs/` for detailed error messages with `--verbose`

### Tesseract not found

**Windows:** Either add `C:\Program Files\Tesseract-OCR` to system PATH, or set in config:
```json
{ "tesseract_path": "C:/Program Files/Tesseract-OCR/tesseract.exe" }
```

**Verify Tesseract works:**
```bash
tesseract --version
```

### `--check-deps` shows camelot / tabula as NO

That is expected unless you deliberately installed them. Both are optional fallback engines — see [Optional extraction engines](#3-optional-extraction-engines). The tool works fully without them.

If you *have* installed them and they still show NO, the system software is missing:

```bash
# Ghostscript (Camelot)
gs --version        # macOS/Linux
gswin64c --version  # Windows

# Java (Tabula)
java -version
```

### Large PDFs are slow

- Reduce `ocr_dpi` to `150` for faster (lower quality) OCR
- Use `--method pdfplumber` to skip camelot/tabula
- Process specific pages by modifying `config.json`

### Corrupted PDF

The script logs the error and continues. Check `logs/` for details. Some PDFs require repair with tools like `qpdf`:
```bash
qpdf --linearize corrupted.pdf fixed.pdf
```

---

## Running Tests

`reportlab` builds the synthetic PDFs the suite converts; it is already in
`requirements.txt`, so there is nothing extra to install.

```bash
# Run all tests
python test_pdf_to_excel.py

# Run with unittest directly
python -m unittest test_pdf_to_excel -v
```

---

## Adding OCR Language Support

1. Download the language pack from [tessdata](https://github.com/tesseract-ocr/tessdata)
2. Place `.traineddata` file in your Tesseract `tessdata/` folder
3. Set in config: `{ "ocr_language": "fra" }`

Common codes: `eng` English, `fra` French, `deu` German, `spa` Spanish,  
`chi_sim` Chinese Simplified, `jpn` Japanese, `ara` Arabic

---

## License

MIT


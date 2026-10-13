# Python Office Automation

Small, practical Python scripts that automate repetitive office work with Excel, CSV, PDF and files.
Each demo comes with sample input files and the output it produces, so you can see exactly what it does.

> All sample data in this repository is made up for demonstration.

## Demos

| # | Demo | What it does |
|---|------|--------------|
| 1 | [Merge & clean Excel/CSV into a report](01_merge_excel_report) | Combines many sales files, fixes messy data, and builds a summary report |
| 2 | [Extract PDF tables to Excel](02_pdf_tables_to_excel) | Pulls line-item tables out of PDF invoices into one spreadsheet |
| 3 | [Batch rename & organise files](03_batch_rename_files) | Renames files to a clean pattern and sorts them by type, with preview and undo |

---

### 1. Merge & clean Excel/CSV into a report

**Input:** 3 files from different stores, with inconsistent column names, mixed date formats, extra spaces, `1,200` vs `1200`, `n/a`, blanks and duplicates.

| Product | Date | Amount |
|---|---|---|
| `  apple ` | `1/1/26` | `1,200` |
| `Apple` | `2026-1-3` | *(blank)* |
| `BANANA` | `Jan 5 2026` | `850` |
| `Banana ` | `1/7/26` | `n/a` |

**Output:** one Excel file ([sample_output/report.xlsx](01_merge_excel_report/sample_output/report.xlsx)) with three sheets:

- **Cleaned Data**: every valid row, standardised, with the source file name
- **Summary**: orders and total amount per product, plus a grand total
- **Removed Rows**: every dropped row and the reason (missing amount, invalid date, duplicate)

```
Files read     : 3
Rows in        : 14
Rows kept      : 10
Rows removed   : 4
```

```bash
cd 01_merge_excel_report
python merge_report.py --input sample_input --output report.xlsx
```

### 2. Extract PDF tables to Excel

**Input:** PDF invoices with a line-item table.
**Output:** [sample_output/tables.xlsx](02_pdf_tables_to_excel/sample_output/tables.xlsx), with an **All Items** sheet (source file and page number on every row) and one sheet per PDF. Amounts like `$1,200.00` become real numbers you can sum.

```bash
cd 02_pdf_tables_to_excel
python pdf_tables_to_excel.py --input sample_input --output tables.xlsx
```

Works with PDFs that contain selectable text. Scanned PDFs need OCR first.

### 3. Batch rename & organise files

Renames files to `<prefix>_<YYYY-MM-DD>_<NNN>.<ext>` and can sort them into `Images/`, `Documents/`, `Spreadsheets/`.

- **Safe by default:** without `--apply` it only shows what it would do
- **Never overwrites** an existing file
- **Undo:** every run writes `rename_log.csv`; `--undo` puts everything back

```bash
cd 03_batch_rename_files
python batch_rename.py --folder sample_input --prefix invoice --by-type          # preview
python batch_rename.py --folder sample_input --prefix invoice --by-type --apply  # rename
python batch_rename.py --folder sample_input --undo                             # revert
```

```
DRY RUN (no files changed):
  IMG_20260103_final.jpg  ->  Images/invoice_2026-01-03_001.jpg
  scan (2).pdf            ->  Documents/invoice_2026-01-03_001.pdf
  Copy of report.xlsx     ->  Spreadsheets/invoice_2026-01-04_001.xlsx
```

The date comes from each file's last-modified time, so you will see different dates on your machine.
(The files in this demo's `sample_input` are small placeholder files.)

---

## Setup

Requires Python 3.9+.

```bash
pip install -r requirements.txt
```

To regenerate the sample files:

```bash
python tools/make_sample_data.py
```

## Need something similar?

These are examples. Real tasks usually need small adjustments for your own file layout, column names or report format.

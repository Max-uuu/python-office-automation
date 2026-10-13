"""
Merge and clean multiple Excel/CSV sales files, then build a summary report.

What it does
------------
1. Reads every .xlsx / .csv file in the input folder
2. Cleans the data:
   - trims spaces and fixes capitalisation in text columns
   - parses mixed date formats into one format (YYYY-MM-DD)
   - converts amounts like "1,200" to numbers; "n/a" or blank becomes empty
   - drops rows with no amount, and exact duplicate rows
3. Writes one Excel report with three sheets:
   - Cleaned Data   : all rows after cleaning, with the source file name
   - Summary        : total amount and order count per product
   - Removed Rows   : rows that were dropped, and why

Usage
-----
    python merge_report.py --input sample_input --output report.xlsx
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = ["Product", "Date", "Amount"]


def read_file(path: Path) -> pd.DataFrame:
    """Read one .xlsx or .csv file as text, so nothing is guessed too early."""
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip().title() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name}: missing column(s) {missing}")

    df = df[REQUIRED_COLUMNS].copy()
    df["Source File"] = path.name
    return df


def clean_product(value: str) -> str:
    value = str(value).strip()
    return value.title() if value else "Unknown"


def clean_amount(value: str):
    value = str(value).strip().replace(",", "")
    if value == "" or value.lower() in {"n/a", "na", "-", "none", "null"}:
        return None
    try:
        return round(float(value), 2)
    except ValueError:
        return None


def clean_date(value: str):
    value = str(value).strip()
    if not value:
        return None
    # Try ISO-like formats first, then day-first, then month-first.
    for dayfirst in (False, True):
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=dayfirst, format="mixed")
        if pd.notna(parsed):
            return parsed.date()
    return None


def clean(df: pd.DataFrame):
    df = df.copy()
    df["Product"] = df["Product"].map(clean_product)
    df["Amount"] = df["Amount"].map(clean_amount)
    df["Date"] = df["Date"].map(clean_date)

    removed = []

    no_amount = df["Amount"].isna()
    removed.append(df[no_amount].assign(Reason="Missing or invalid amount"))
    df = df[~no_amount]

    no_date = df["Date"].isna()
    removed.append(df[no_date].assign(Reason="Missing or invalid date"))
    df = df[~no_date]

    dupes = df.duplicated(subset=REQUIRED_COLUMNS, keep="first")
    removed.append(df[dupes].assign(Reason="Duplicate row"))
    df = df[~dupes]

    removed_df = pd.concat(removed, ignore_index=True)
    df = df.sort_values(["Date", "Product"]).reset_index(drop=True)
    return df, removed_df


def summarise(df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df.groupby("Product", as_index=False)
        .agg(Orders=("Amount", "count"), Total_Amount=("Amount", "sum"))
        .sort_values("Total_Amount", ascending=False)
        .rename(columns={"Total_Amount": "Total Amount"})
    )
    total = pd.DataFrame(
        [{"Product": "TOTAL", "Orders": summary["Orders"].sum(),
          "Total Amount": summary["Total Amount"].sum()}]
    )
    return pd.concat([summary, total], ignore_index=True)


def style_sheet(ws, money_cols=()):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for col_idx, col in enumerate(ws.columns, start=1):
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col) + 3
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 10), 40)
        if ws.cell(1, col_idx).value in money_cols:
            for c in col[1:]:
                c.number_format = "#,##0.00"
    ws.freeze_panes = "A2"


def main():
    parser = argparse.ArgumentParser(description="Merge and clean Excel/CSV files into one report.")
    parser.add_argument("--input", default="sample_input", help="Folder with .xlsx/.csv files")
    parser.add_argument("--output", default="report.xlsx", help="Output Excel file")
    args = parser.parse_args()

    input_dir = Path(args.input)
    files = sorted(p for p in input_dir.iterdir() if p.suffix.lower() in {".xlsx", ".csv"})
    if not files:
        sys.exit(f"No .xlsx or .csv files found in {input_dir}")

    frames = []
    for path in files:
        try:
            frames.append(read_file(path))
            print(f"  read  {path.name}")
        except ValueError as e:
            print(f"  skip  {e}")
    if not frames:
        sys.exit("No usable files.")

    raw = pd.concat(frames, ignore_index=True)
    cleaned, removed = clean(raw)
    summary = summarise(cleaned)

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        removed.to_excel(writer, sheet_name="Removed Rows", index=False)
        style_sheet(writer.sheets["Cleaned Data"], money_cols=("Amount",))
        style_sheet(writer.sheets["Summary"], money_cols=("Total Amount",))
        style_sheet(writer.sheets["Removed Rows"])
        last = writer.sheets["Summary"].max_row
        for cell in writer.sheets["Summary"][last]:
            cell.font = Font(bold=True)

    print(f"\nFiles read     : {len(frames)}")
    print(f"Rows in        : {len(raw)}")
    print(f"Rows kept      : {len(cleaned)}")
    print(f"Rows removed   : {len(removed)}")
    print(f"Report saved to: {args.output}")


if __name__ == "__main__":
    main()

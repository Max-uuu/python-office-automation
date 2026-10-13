"""Generate the fake sample files used by the demos. All data here is made up."""

import os
import time
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).resolve().parent.parent


def excel_samples():
    out = ROOT / "01_merge_excel_report" / "sample_input"
    out.mkdir(parents=True, exist_ok=True)

    pd.DataFrame({
        "Product": ["  apple ", "Apple", "BANANA", "apple", "Banana ", ""],
        "Date": ["1/1/26", "2026-1-3", "Jan 5 2026", "2026-01-05", "1/7/26", "2026/01/08"],
        "Amount": ["1,200", "", "850", "1200", "n/a", "430"],
    }).to_excel(out / "sales_jan_store_a.xlsx", index=False)

    pd.DataFrame({
        "product": ["orange", "Apple", "ORANGE", "banana"],
        "date": ["2026-01-10", "2026-01-12", "2026-01-10", "Jan 15 2026"],
        "amount": ["300", "1,050.50", "300", "920"],
    }).to_csv(out / "sales_jan_store_b.csv", index=False)

    pd.DataFrame({
        "Product": ["Grape", "grape ", "Apple", "Banana"],
        "Date": ["2026-01-20", "2026-01-21", "2026-01-22", "not sure"],
        "Amount": ["640", "640", "1,180", "500"],
    }).to_excel(out / "sales_jan_store_c.xlsx", index=False)


def pdf_samples():
    out = ROOT / "02_pdf_tables_to_excel" / "sample_input"
    out.mkdir(parents=True, exist_ok=True)
    styles = getSampleStyleSheet()

    invoices = {
        "invoice_1001.pdf": ("ACME Supplies (fictional)", [
            ["Item", "Qty", "Unit Price", "Total"],
            ["A4 Paper (box)", "10", "4.50", "45.00"],
            ["Ink Cartridge", "3", "22.00", "66.00"],
            ["Stapler", "2", "8.75", "17.50"],
        ]),
        "invoice_1002.pdf": ("Northwind Office (fictional)", [
            ["Item", "Qty", "Unit Price", "Total"],
            ["Desk Lamp", "4", "18.00", "72.00"],
            ["USB Cable", "12", "3.20", "38.40"],
            ["Notebook", "25", "1.80", "45.00"],
            ["Whiteboard Marker", "30", "0.90", "27.00"],
        ]),
    }

    for name, (vendor, data) in invoices.items():
        doc = SimpleDocTemplate(str(out / name), pagesize=A4)
        table = Table(data, colWidths=[200, 60, 90, 90])
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ]))
        doc.build([
            Paragraph(f"INVOICE {name[8:12]}", styles["Title"]),
            Paragraph(vendor, styles["Normal"]),
            Spacer(1, 20),
            table,
        ])


def rename_samples():
    out = ROOT / "03_batch_rename_files" / "sample_input"
    out.mkdir(parents=True, exist_ok=True)
    names = ["IMG_20260103_final.jpg", "scan (2).pdf", "Copy of report.xlsx",
             "notes.txt", "photo 1.PNG", "data_export.csv"]
    base = time.mktime((2026, 1, 3, 9, 0, 0, 0, 0, -1))
    for i, n in enumerate(names):
        p = out / n
        p.write_text(f"sample file {i}\n", encoding="utf-8")
        t = base + (i // 2) * 86400 + i * 60
        os.utime(p, (t, t))


if __name__ == "__main__":
    excel_samples()
    pdf_samples()
    rename_samples()
    print("Sample data created.")

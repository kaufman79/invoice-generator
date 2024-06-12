import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # set the header
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice # {invoice_nr}", align="L",
             ln=1, border=0)
    pdf.cell(w=0, h=12, txt= f"Date", align="L",
             ln=1, border=0)

    # set the table
    pdf.ln(14)
    # column headings
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=25, h=12, txt="Product ID", align="L",
             ln=0, border=1)
    pdf.cell(w=80, h=12, txt="Product Name", align="L",
             ln=0, border=1)
    pdf.output(f"PDFs/{filename}.pdf")


import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # set the header
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    year, month, day = date.split(".")

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice # {invoice_nr}", align="L",
             ln=1, border=0)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {month}/{day}/{year}", align="L",
             ln=1, border=0)

    # set the table
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # add column headings
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, border=1, txt=columns[0])
    pdf.cell(w=70, h=8, border=1, txt=columns[1])
    pdf.cell(w=30, h=8, border=1, txt="Amount")
    pdf.cell(w=30, h=8, border=1, txt=columns[3])
    pdf.cell(w=30, h=8, border=1, txt=columns[4], ln=1)

    # add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, border=1, txt=str(row["product_id"]))
        pdf.cell(w=70, h=8, border=1, txt=str(row["product_name"]))
        pdf.cell(w=30, h=8, border=1, txt=str(row["amount_purchased"]))
        pdf.cell(w=30, h=8, border=1, txt=str(row["price_per_unit"]))
        pdf.cell(w=30, h=8, border=1, txt=str(row["total_price"]), ln=1)

    # add sum total
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, border=1, txt=str(total_sum), ln=1)

    # add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=12, txt=f"The total price is {total_sum}", ln=1)

    # add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)


    pdf.output(f"PDFs/{filename}.pdf")

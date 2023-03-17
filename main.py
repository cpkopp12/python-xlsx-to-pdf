# when choosing where to start on app without GUI
# just consider flow from input to out put

import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for path in filepaths:
    # excel data frame
    df = pd.read_excel(path, sheet_name="Sheet 1")
    # print(df)

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Title
    pdf.set_font(family="Times", size=16, style="B")
    filename = Path(path).stem
    invoice_nr, date = filename.split('-')
    # extract file name from filepath (path)
    pdf.cell(w=50, h=8, txt=f"Invoice #: {invoice_nr}", ln=1)

    # Output the date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # header row
    header = list(df.columns)
    header = [item.replace('_', " ").title() for item in header]
    pdf.set_font(family="Times", size=9, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(header[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(header[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(header[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(header[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(header[4]), border=1, ln=1)

    # sum for total price
    sum1 = 0
    # iterate over df
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=9)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
        sum1 = row["total_price"] + sum1
    # total price
    pdf.set_font(family="Times", size=9)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, txt=str(sum1), border=1, ln=1)

    # sum sentence
    pdf.set_font(family="Times", size=9)
    pdf.cell(w=30, h=8, txt=f"The total price {sum1}.", ln=1)

    # add logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"PythonHow", ln=1)
    pdf. image("pythonhow.png", w=10)

    # output file
    pdf.output(f"pdfs/{filename}.pdf")
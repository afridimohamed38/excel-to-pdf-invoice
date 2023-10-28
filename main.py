import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    # Pdf initialization section
    pdf = FPDF(orientation="P", format="A4")
    pdf.add_page()

    # getting filename and invoice number
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Writing to pdf
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    pdf.cell(w=50,h= 8, txt=" ", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Adding headers to table
    header_columns = list(df.columns)
    header_columns = [item.replace("_", " ").title() for item in header_columns]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt= header_columns[0], border=1)
    pdf.cell(w=60, h=8, txt= header_columns[1], border=1)
    pdf.cell(w=35, h=8, txt= header_columns[2], border=1)
    pdf.cell(w=30, h=8, txt= header_columns[3], border=1)
    pdf.cell(w=30, h=8, txt= header_columns[4], border=1, ln=1)

    # inserting data in the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")

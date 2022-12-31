from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    invoice_date = filename.split("-")[1]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    headers = list(df.columns)
    headers = [header.replace("_", " ").title() for header in headers]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=headers[0], border=1)
    pdf.cell(w=65, h=8, txt=headers[1], border=1)
    pdf.cell(w=35, h=8, txt=headers[2], border=1)
    pdf.cell(w=30, h=8, txt=headers[3], border=1)
    pdf.cell(w=30, h=8, txt=headers[4], ln=1, border=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), ln=1, border=1)

    total_price = df["total_price"].sum()
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=65, h=8, border=1)
    pdf.cell(w=35, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), ln=1, border=1)

    # Add total sum
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_price}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt="Charanvir Singh Automation")
    # This would be how to add an image to a pdf
    # pdf.image("image file path")

    pdf.output(f"PDFs/{filename}.pdf")

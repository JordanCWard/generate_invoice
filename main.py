import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
# import openpyxl

# openpyxl is a Python library to read/write Excel files
# glob allows you to search for files with specific patterns

# Get all files ending with .xlsx
invoices = glob.glob("invoices/*.xlsx")

for excel_file in invoices:

    # Create pdf for each excel_file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extracting the name of the file
    filename = Path(excel_file).stem

    # Extract the invoice number and date from the filename
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")

    # ln adds 1 break line after this cell
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # These pandas command only work with openpxyl downloaded, see imports
    df = pd.read_excel(excel_file, sheet_name="Sheet 1")

    # Add a header to each table
    column_names = list(df.columns)
    column_names = [item.replace("_", " ").title() for item in column_names]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=column_names[0], border=1)
    pdf.cell(w=70, h=8, txt=column_names[1], border=1)
    pdf.cell(w=30, h=8, txt=column_names[2], border=1)
    pdf.cell(w=30, h=8, txt=column_names[3], border=1)
    pdf.cell(w=30, h=8, txt=column_names[4], border=1, ln=1)

    # Add rows to each table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)

        # str() added because row expects an integer but txt expects string
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)

        # Add ln to this to move to the next row after this
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Sum of the prices, added str because txt expects a string
    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    # placing each file in a folder
    pdf.output(f"PDFs/{filename}.pdf")

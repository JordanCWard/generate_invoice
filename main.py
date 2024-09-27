import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
# import openpyxl

# openpyxl is a Python library to read/write Excel files
# glob allows you to search for files with specific patterns

# get all files ending with .xlsx
invoices = glob.glob("invoices/*.xlsx")

for excel_file in invoices:

    # this pandas command only works with openpxyl downloaded, see imports
    df = pd.read_excel(excel_file, sheet_name="Sheet 1")

    # create pdf for each excel_file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # extracting the name of the file
    filename = Path(excel_file).stem

    # extract the invoice number and date from the filename
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")

    # ln adds 1 break line after this cell
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    # placing each file in a folder
    pdf.output(f"PDFs/{filename}.pdf")

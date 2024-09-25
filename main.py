import pandas as pd
import openpyxl
import glob

# openpyxl is a Python library to read/write Excel files
# glob allows you to search for files with specific patterns

# get all files ending with .xlsx
invoices = glob.glob("invoices/*.xlsx")

for excel_file in invoices:
    # this pandas command only works with openpxyl included
    df = pd.read_excel(excel_file, sheet_name="Sheet 1")
    print(df)

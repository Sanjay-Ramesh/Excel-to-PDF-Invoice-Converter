import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Bills\\*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    #add pages and create new files
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    #Extracting filename and modifying for Invoice and Date Line 
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    #Setting fonts for Invoice Line
    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=50, h=11,txt=f"Invoice nr.{invoice_nr}", ln=1)

    #Setting fonts for Date Line
    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=50, h=11,txt=f"Date {date}", ln=1)

    #Output
    pdf.output(f"PDFS\\{filename}.pdf")
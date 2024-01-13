import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Bills\\*.xlsx")

for filepath in filepaths:
    
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

    #Extracting datas from excel sheets
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    #Add a Header
    pdf.set_font(family="Times", size = 10, style="B")
    pdf.set_text_color(80, 80, 80)

     #Cells for contents in table
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=34, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)
    

    #Add rows to the tables
    for index, row in df.iterrows():

        #Setting Fonts and color for tables
        pdf.set_font(family="Times", size = 14)
        pdf.set_text_color(80, 80, 80)

        #Cells for contents in table
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=34, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
    
    #Total
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size = 14)
    pdf.set_text_color(80, 80, 80)

    #Cells for Total
    pdf.cell(w=30, h=8, txt="",border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=34, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    #For Outer Lines
    pdf.set_font(family="Times", size = 14, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=20, txt=f"The Total Price is {total_sum} Euros", 
             ln=1)
    
    #for company name and logo
    pdf.set_font(family="Times", size = 14, style="B")
    pdf.cell(w=25, h=10, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)
    

    #Output
    pdf.output(f"PDFS\\{filename}.pdf")
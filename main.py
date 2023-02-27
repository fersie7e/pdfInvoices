import pandas as pd
import glob
import fpdf
import pathlib


filespath = glob.glob("invoices/*.xlsx")

for filepath in filespath:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = fpdf.FPDF(orientation="P", unit="mm",format="A4")
    pdf.add_page()

    filename = pathlib.Path(filepath).stem
    nro_factura = filename.split("-")[0]
    fecha_factura = filename.split("-")[1]

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Factura nro. {nro_factura}")
    pdf.output(f"PDFs/{filename}.pdf")
    print(df)
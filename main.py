import pandas as pd
import glob
import fpdf
import pathlib


filespath = glob.glob("invoices/*.xlsx")

for filepath in filespath:
    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = pathlib.Path(filepath).stem
    nro_factura = filename.split("-")[0]
    fecha_factura = filename.split("-")[1]

    pdf.set_font(family="Arial", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"Factura nro. {nro_factura}",ln=1)
    pdf.cell(w=50, h=15, txt=f"Fecha: {fecha_factura}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    header = df.columns
    header = [item.replace("_", " ").title() for item in header]

    # Header of the table
    pdf.set_font(family="Arial", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(header[0]), border=1)
    pdf.cell(w=60, h=8, txt=str(header[1]), border=1)
    pdf.cell(w=40, h=8, txt=str(header[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(header[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(header[4]),ln=1, border=1)

    # Data of the table

    for index, row in df.iterrows():
        pdf.set_font(family="Arial", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, align="R", ln=1)

    # Total
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Arial", size=12, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=60, h=8, txt="")
    pdf.cell(w=40, h=8, txt="")
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=30, h=8, txt=f"{str(total_sum)} Euros", border=1, align="R", ln=1)

    pdf.set_font(family="Arial", size=12, style="B")
    pdf.cell(w=30, h=8,
             txt=f"El total de la factura es : {str(total_sum)} Euros", ln=1)
    pdf.cell(w=55, h=8,
             txt=f"fersie7e webDevelopment")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
    print(df)
# Import library pandas untuk pengolahan data, glob untuk mendapatkan file, FPDF untuk membuat PDF, dan Path dari pathlib untuk manipulasi path file
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Mengambil semua file Excel di folder "invoices" dengan ekstensi .xlsx
filepaths = glob.glob("invoices/*.xlsx")

# Melakukan iterasi pada setiap file Excel di dalam folder "invoices"
for filepath in filepaths:
    # Membuat objek PDF baru dengan orientasi Portrait, satuan ukuran mm, dan format A4
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()  # Menambahkan halaman baru ke PDF

    # Mendapatkan nama file tanpa ekstensi dari path file
    filename = Path(filepath).stem
    # Mengambil nomor invoice dari nama file, diasumsikan bahwa formatnya adalah 'nomorInvoice-....xlsx'
    invoice_nr, date = filename.split("-")

    # Menentukan font dengan jenis "Times", ukuran 16, dan tebal (Bold)
    pdf.set_font(family="Times", size=16, style="B")
    # Menambahkan teks ke PDF dengan lebar 50 mm dan tinggi 8 mm, berisi nomor invoice
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    # Menentukan font dengan jenis "Times", ukuran 16, dan tebal (Bold)
    pdf.set_font(family="Times", size=16, style="B")
    # Menambahkan teks ke PDF dengan lebar 50 mm dan tinggi 8 mm, berisi tanggal invoice
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Membaca file Excel dan mengambil data dari sheet "Sheet 1"
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # menambahkan header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)


    # menambahkan row table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Menyimpan file PDF ke dalam folder "PDFs" dengan nama yang sama seperti nama file Excel
    pdf.output(f"PDFs/{filename}.pdf")

import re
import pandas as pd
import xlsxwriter
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from pdfplumber import open as pdf_open

def extract_data_from_pdf(file):
    data = {}
    with pdf_open(file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

        # Mã EVN
        match = re.search(r'Mã khách hàng \(Customer\'s Code\):\s*(\S+)', text)
        data["Mã EVN"] = match.group(1) if match else ""

        # Mã tỉnh cố định
        data["Mã tỉnh"] = "YBI"

        # Kỳ
        data["Kỳ"] = "1"

        # Mã CSHT cố định hoặc lấy từ tên file
        data["Mã CSHT"] = "CSHT_YBI_00014"

        # Ngày đầu kỳ
        match = re.search(r'Điện tiêu thụ tháng \d+ năm \d+ từ ngày (\d{2}/\d{2}/\d{4}) đến ngày', text)
        data["Ngày đầu kỳ"] = match.group(1) if match else ""

        # Ngày cuối kỳ
        match = re.search(r'đến ngày (\d{2}/\d{2}/\d{4})', text)
        data["Ngày cuối kỳ"] = match.group(1) if match else ""

        # Tổng chỉ số
        match = re.search(r'kWh\s*([\d\.]+)', text)
        data["Tổng chỉ số"] = match.group(1) if match else ""

        # Số tiền
        match = re.search(r'Cộng tiền hàng.*\n.*([\d\.]+)', text)
        data["Số tiền"] = match.group(1) if match else ""

        # Thuế VAT
        match = re.search(r'Tiền thuế GTGT.*\n.*([\d\.]+)', text)
        data["Thuế VAT"] = match.group(1) if match else ""

        # Số tiền dự kiến
        match = re.search(r'Tổng cộng tiền thanh toán.*\n.*([\d\.]+)', text)
        data["Số tiền dự kiến"] = match.group(1) if match else ""

        # Số hóa đơn: chưa có cách lấy, để trống
        data["Số hóa đơn"] = ""

    return data

def create_excel(data_list):
    df = pd.DataFrame(data_list)
    writer = BytesIO()
    with pd.ExcelWriter(writer, engine='xlsxwriter') as writer_obj:
        df.to_excel(writer_obj, index=False, sheet_name='Sheet1')
        workbook = writer_obj.book
        worksheet = writer_obj.sheets['Sheet1']

        fmt_text = workbook.add_format({'num_format': '@'})
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 20, fmt_text)

    return writer.getvalue()

def create_word(data_list):
    document = Document()

    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    if not data_list:
        return document

    df = pd.DataFrame(data_list)
    table = document.add_table(rows=1, cols=len(df.columns))
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = col_name

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(df.columns):
            row_cells[i].text = str(row[col])

    return document

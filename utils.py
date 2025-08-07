import re
import pdfplumber
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def extract_data_from_pdf(file):
    data = {
        "Mã tỉnh": "YBI",  # cố định
        "Số hóa đơn": "",
        "Mã EVN": "",
        "Mã tháng (yyyyMM)": "",
        "Kỳ": "1",
        "Mã CSHT": "CSHT_YBI_00014",
        "Ngày đầu kỳ": "",
        "Ngày cuối kỳ": "",
        "Tổng chỉ số": "",
        "Số tiền": "",
        "Thuế VAT": "",
        "Số tiền dự kiến": "",
        "Ghi chú": ""
    }
    
    with pdfplumber.open(file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

        # Số hóa đơn
        m = re.search(r"Số hóa đơn\s*:\s*(\d+)", text)
        if m:
            data["Số hóa đơn"] = m.group(1).strip()

        # Mã EVN (Mã khách hàng)
        m = re.search(r"Mã khách hàng\s*:\s*([A-Z0-9]+)", text)
        if m:
            data["Mã EVN"] = m.group(1).strip()

        # Mã tháng (yyyyMM)
        m = re.search(r"Điện tiêu thụ tháng\s*(\d+)\s*năm\s*(\d{4})", text)
        if m:
            month = int(m.group(1))
            year = int(m.group(2))
            data["Mã tháng (yyyyMM)"] = f"{year}{month:02d}"

        # Ngày đầu kỳ và ngày cuối kỳ
        m = re.search(r"Điện tiêu thụ tháng \d+ năm \d{4} từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})", text)
        if m:
            data["Ngày đầu kỳ"] = m.group(1)
            data["Ngày cuối kỳ"] = m.group(2)

        # Tổng chỉ số (kWh)
        m = re.search(r'kWh\s*([\d.,]+)', text)
        if m:
            data["Tổng chỉ số"] = m.group(1).replace(',', '')

        # Số tiền (Cộng tiền hàng)
        m = re.search(r'Cộng tiền hàng.*?([\d.,]+)', text, re.DOTALL)
        if m:
            data["Số tiền"] = m.group(1).replace(',', '')

        # Thuế VAT (Tiền thuế GTGT)
        m = re.search(r'Tiền thuế GTGT.*?([\d.,]+)', text, re.DOTALL)
        if m:
            data["Thuế VAT"] = m.group(1).replace(',', '')

        # Số tiền dự kiến (Tổng cộng tiền thanh toán)
        m = re.search(r'Tổng cộng tiền thanh toán.*?([\d.,]+)', text, re.DOTALL)
        if m:
            data["Số tiền dự kiến"] = m.group(1).replace(',', '')

    return data

def create_excel(data_list):
    df = pd.DataFrame(data_list)

    columns_order = [
        "Mã tỉnh", "Số hóa đơn", "Mã EVN", "Mã tháng (yyyyMM)", "Kỳ",
        "Mã CSHT", "Ngày đầu kỳ", "Ngày cuối kỳ", "Tổng chỉ số",
        "Số tiền", "Thuế VAT", "Số tiền dự kiến", "Ghi chú"
    ]
    df = df[columns_order]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Hóa đơn')
        worksheet = writer.sheets['Hóa đơn']

        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

        # Định dạng tất cả cột dạng text
        fmt = writer.book.add_format({'num_format': '@'})
        worksheet.set_column(0, len(df.columns)-1, None, fmt)

    output.seek(0)
    return output.read()

def create_word(data_list):
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    document.add_heading('Dữ liệu hóa đơn', level=1)

    columns = [
        "Mã tỉnh", "Số hóa đơn", "Mã EVN", "Mã tháng (yyyyMM)", "Kỳ",
        "Mã CSHT", "Ngày đầu kỳ", "Ngày cuối kỳ", "Tổng chỉ số",
        "Số tiền", "Thuế VAT", "Số tiền dự kiến", "Ghi chú"
    ]

    table = document.add_table(rows=1, cols=len(columns))
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(columns):
        hdr_cells[i].text = col_name

    for data in data_list:
        row_cells = table.add_row().cells
        for i, col_name in enumerate(columns):
            val = data.get(col_name, "")
            row_cells[i].text = str(val)

    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output.read()

import re
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from io import BytesIO

def extract_data_from_pdf(pdf_file):
    # Dùng pdfplumber để trích xuất dữ liệu từ file PDF hóa đơn
    import pdfplumber

    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

    # Khởi tạo dict dữ liệu
    data = {
        "Mã tỉnh": "YBI",  # luôn YBI
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

    # Lấy số hóa đơn
    m = re.search(r'Số hóa đơn\s*:\s*(\S+)', text)
    if m:
        data["Số hóa đơn"] = m.group(1).strip()

    # Lấy mã EVN (Mã khách hàng)
    m = re.search(r'Mã khách hàng\s*:\s*(\S+)', text)
    if m:
        data["Mã EVN"] = m.group(1).strip()

    # Lấy Mã tháng (yyyyMM)
    # Lấy tháng từ đoạn 'Điện tiêu thụ tháng 7 năm 2025 ...'
    m = re.search(r'Điện tiêu thụ tháng (\d{1,2}) năm (\d{4})', text)
    if m:
        month = int(m.group(1))
        year = int(m.group(2))
        data["Mã tháng (yyyyMM)"] = f"{year}{month:02d}"

    # Kỳ luôn 1 theo yêu cầu

    # Mã CSHT luôn "CSHT_YBI_00014" theo yêu cầu

    # Lấy ngày đầu kỳ và ngày cuối kỳ từ đoạn "Điện tiêu thụ tháng ... từ ngày dd/MM/yyyy đến ngày dd/MM/yyyy"
    m = re.search(r'Điện tiêu thụ tháng \d+ năm \d+ từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})', text)
    if m:
        data["Ngày đầu kỳ"] = m.group(1)
        data["Ngày cuối kỳ"] = m.group(2)

    # Lấy Tổng chỉ số (Số lượng kWh)
    m = re.search(r'kWh\s+([\d.,]+)', text)
    if m:
        data["Tổng chỉ số"] = m.group(1).replace(',', '')

    # Lấy Số tiền (Cộng tiền hàng)
    m = re.search(r'Cộng tiền hàng.*?([\d.,]+)', text, re.DOTALL)
    if m:
        data["Số tiền"] = m.group(1).replace(',', '')

    # Lấy Thuế VAT (Tiền thuế GTGT)
    m = re.search(r'Tiền thuế GTGT.*?([\d.,]+)', text, re.DOTALL)
    if m:
        data["Thuế VAT"] = m.group(1).replace(',', '')

    # Lấy Số tiền dự kiến (Tổng cộng tiền thanh toán)
    m = re.search(r'Tổng cộng tiền thanh toán.*?([\d.,]+)', text, re.DOTALL)
    if m:
        data["Số tiền dự kiến"] = m.group(1).replace(',', '')

    # Ghi chú để trống (hoặc bổ sung nếu cần)
    data["Ghi chú"] = ""

    return data

def create_excel(data_list):
    df = pd.DataFrame(data_list)

    # Đổi thứ tự cột theo yêu cầu
    columns_order = [
        "Mã tỉnh", "Số hóa đơn", "Mã EVN", "Mã tháng (yyyyMM)", "Kỳ",
        "Mã CSHT", "Ngày đầu kỳ", "Ngày cuối kỳ", "Tổng chỉ số",
        "Số tiền", "Thuế VAT", "Số tiền dự kiến", "Ghi chú"
    ]
    df = df[columns_order]

    # Xuất Excel vào bytes
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Hóa đơn', index=False)

        workbook = writer.book
        worksheet = writer.sheets['Hóa đơn']

        # Tự động điều chỉnh độ rộng cột
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

        # Định dạng tất cả cột dạng text (tránh tự động chuyển số sang khoa học)
        fmt = workbook.add_format({'num_format': '@'})
        worksheet.set_column(0, len(df.columns) - 1, None, fmt)

    output.seek(0)
    return output.read()

def create_word(data_list):
    document = Document()

    # Đặt font Times New Roman và size 14 cho toàn văn bản
    style = document.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    # Sửa font cho toàn bộ đoạn văn trong docx
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    document.add_heading('Dữ liệu Hóa đơn', level=1)

    # Thêm bảng
    columns = [
        "Mã tỉnh", "Số hóa đơn", "Mã EVN", "Mã tháng (yyyyMM)", "Kỳ",
        "Mã CSHT", "Ngày đầu kỳ", "Ngày cuối kỳ", "Tổng chỉ số",
        "Số tiền", "Thuế VAT", "Số tiền dự kiến", "Ghi chú"
    ]
    table = document.add_table(rows=1, cols=len(columns))
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(columns):
        hdr_cells[i].text = col_name

    # Thêm dữ liệu từng dòng
    for data in data_list:
        row_cells = table.add_row().cells
        for i, col_name in enumerate(columns):
            val = data.get(col_name, "")
            if val is None:
                val = ""
            row_cells[i].text = str(val)

    return document

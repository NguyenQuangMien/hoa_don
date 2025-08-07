import pandas as pd
from io import BytesIO
from docx import Document
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def extract_period(text):
    # Chuẩn hóa chuỗi text: bỏ xuống dòng, nhiều khoảng trắng thành 1 khoảng trắng
    normalized = re.sub(r'\s+', ' ', text)

    # Tìm đoạn mô tả dạng "Điện tiêu thụ tháng 7 năm 2025 từ ngày 22/06/2025 đến ngày 21/07/2025"
    pattern = r'Điện tiêu thụ tháng \d{1,2} năm \d{4} từ ngày (\d{1,2}/\d{1,2}/\d{4}) đến ngày (\d{1,2}/\d{1,2}/\d{4})'
    match = re.search(pattern, normalized, re.IGNORECASE)
    if match:
        return match.group(1), match.group(2)
    else:
        return None, None

def extract_data_from_pdf(text):
    # DEBUG: In text để kiểm tra đầu vào
    print("=== DEBUG TEXT START ===")
    print(text)
    print("=== DEBUG TEXT END ===")

    data = {}

    data['Mã tỉnh'] = 'YBI'

    match_invoice = re.search(r'Số \(No\):\s*(\d+)', text)
    data['Số hóa đơn'] = match_invoice.group(1) if match_invoice else ''

    match_customer_code = re.search(r'Mã khách hàng \(Customer\'s Code\):\s*([A-Z0-9]+)', text)
    data['Mã EVN'] = match_customer_code.group(1) if match_customer_code else ''

    match_date = re.search(r'Ngày \(Date\) (\d{1,2}) tháng (\d{1,2}) năm (\d{4})', text)
    if match_date:
        year = match_date.group(3)
        month = match_date.group(2).zfill(2)
        data['Mã tháng (yyyyMM)'] = f"{year}{month}"
    else:
        data['Mã tháng (yyyyMM)'] = ''

    data['Kỳ'] = '1'
    data['Mã CSHT'] = 'CSHT_YBI_00014'

    # Lấy ngày đầu kỳ và ngày cuối kỳ từ chuỗi mô tả hóa đơn
    ngay_dau, ngay_cuoi = extract_period(text)
    data['Ngày đầu kỳ'] = ngay_dau if ngay_dau else ''
    data['Ngày cuối kỳ'] = ngay_cuoi if ngay_cuoi else ''

    match_kwh = re.search(r'kWh\s*([\d.,]+)', text)
    data['Tổng chỉ số'] = match_kwh.group(1) if match_kwh else ''

    match_amount = re.search(r'Cộng tiền hàng \(Total amount\):\s*([\d.,]+)', text)
    if not match_amount:
        match_amount = re.search(r'Tổng cộng tiền thanh toán\s*:\s*([\d.,]+)', text)
    data['Số tiền'] = match_amount.group(1) if match_amount else ''

    match_vat = re.search(r'Tiền thuế GTGT \(VAT amount\):\s*([\d.,]+)', text)
    data['Thuế VAT'] = match_vat.group(1) if match_vat else ''

    match_total = re.search(r'Tổng cộng tiền thanh toán\s*:\s*([\d.,]+)', text)
    data['Số tiền dự kiến'] = match_total.group(1) if match_total else ''

    data['Ghi chú'] = ''

    return data

def create_excel(df):
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = list(df.columns)
    ws.append(headers)

    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    for col_idx, col in enumerate(ws.columns, 1):
        max_length = 0
        col_letter = get_column_letter(col_idx)
        for cell in col:
            cell.number_format = '@'  # Định dạng Text
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
            cell.alignment = Alignment(horizontal='left', vertical='center')
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def create_word(df):
    doc = Document()
    doc.add_heading("Bảng dữ liệu hóa đơn", level=1)
    table = doc.add_table(rows=1, cols=len(df.columns))
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = col_name

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(df.columns):
            row_cells[i].text = str(row[col_name])

    output = BytesIO()
    doc.save(output)
    return output.getvalue()

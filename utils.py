import pandas as pd
from io import BytesIO
from docx import Document
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def extract_data_from_pdf(text):
    data = {}

    # Mã tỉnh cố định
    data['Mã tỉnh'] = 'YBI'

    # Số hóa đơn (No)
    match_invoice = re.search(r'Số \(No\):\s*(\d+)', text)
    data['Số hóa đơn'] = match_invoice.group(1) if match_invoice else ''

    # Mã EVN (Mã khách hàng)
    match_customer_code = re.search(r'Mã khách hàng \(Customer\'s Code\):\s*([A-Z0-9]+)', text)
    data['Mã EVN'] = match_customer_code.group(1) if match_customer_code else ''

    # Mã tháng yyyyMM từ ngày hóa đơn
    match_date = re.search(r'Ngày \(Date\) (\d{1,2}) tháng (\d{1,2}) năm (\d{4})', text)
    if match_date:
        year = match_date.group(3)
        month = match_date.group(2).zfill(2)
        data['Mã tháng (yyyyMM)'] = f"{year}{month}"
    else:
        data['Mã tháng (yyyyMM)'] = ''

    # Kỳ mặc định 1
    data['Kỳ'] = '1'

    # Mã CSHT cố định
    data['Mã CSHT'] = 'CSHT_YBI_00014'

    # Lấy ngày đầu kỳ, ngày cuối kỳ bằng cách tách dòng, tìm dòng chứa ngày
    lines = text.split('\n')
    date_start, date_end = '', ''
    for line in lines:
        if 'Điện tiêu thụ tháng' in line:
            m = re.search(
                r'từ ngày\s*(\d{1,2}/\d{1,2}/\d{4})\s*đến ngày\s*(\d{1,2}/\d{1,2}/\d{4})',
                line, re.IGNORECASE
            )
            if m:
                date_start = m.group(1)
                date_end = m.group(2)
                break
    data['Ngày đầu kỳ'] = date_start
    data['Ngày cuối kỳ'] = date_end

    # Debug ngày
    print("Ngày đầu kỳ:", date_start)
    print("Ngày cuối kỳ:", date_end)

    # Tổng chỉ số (giữ nguyên chuỗi trong hóa đơn)
    match_kwh = re.search(r'kWh\s*([\d.,]+)', text)
    data['Tổng chỉ số'] = match_kwh.group(1) if match_kwh else ''

    # Số tiền nguyên bản
    match_amount = re.search(r'Cộng tiền hàng \(Total amount\):\s*([\d.,]+)', text)
    if not match_amount:
        match_amount = re.search(r'Tổng cộng tiền thanh toán\s*:\s*([\d.,]+)', text)
    data['Số tiền'] = match_amount.group(1) if match_amount else ''

    # Thuế VAT nguyên bản
    match_vat = re.search(r'Tiền thuế GTGT \(VAT amount\):\s*([\d.,]+)', text)
    data['Thuế VAT'] = match_vat.group(1) if match_vat else ''

    # Số tiền dự kiến nguyên bản (Thành tiền)
    match_total = re.search(r'Thành tiền\s*:\s*([\d.,]+)', text)
    if not match_total:
        match_total = re.search(r'Tổng cộng thanh toán\s*:\s*([\d.,]+)', text)
    data['Số tiền dự kiến'] = match_total.group(1) if match_total else ''

    # Ghi chú để trống
    data['Ghi chú'] = ''

    return data


def create_excel(df):
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Viết tiêu đề giữ nguyên thứ tự cột
    headers = list(df.columns)
    ws.append(headers)

    # Viết dữ liệu
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    # Định dạng tất cả cột dạng Text, dãn cột tự động
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

    # Lưu vào BytesIO
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

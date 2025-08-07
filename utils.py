import pandas as pd
from io import BytesIO
from docx import Document
import re

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

    # Ngày đầu kỳ và ngày cuối kỳ
    match_period = re.search(r'Điện tiêu thụ tháng \d+ năm \d+ từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})', text)
    if match_period:
        data['Ngày đầu kỳ'] = match_period.group(1)
        data['Ngày cuối kỳ'] = match_period.group(2)
    else:
        data['Ngày đầu kỳ'] = ''
        data['Ngày cuối kỳ'] = ''

    # Tổng chỉ số (nguyên bản)
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

    # Số tiền dự kiến nguyên bản
    match_total = re.search(r'Thành tiền\s*:\s*([\d.,]+)', text)
    if not match_total:
        match_total = re.search(r'Tổng cộng thanh toán\s*:\s*([\d.,]+)', text)
    data['Số tiền dự kiến'] = match_total.group(1) if match_total else ''

    # Ghi chú để trống
    data['Ghi chú'] = ''

    return data

def clean_number(value):
    if not value:
        return 0
    v = value.replace('.', '').replace(',', '.')
    try:
        return float(v)
    except:
        return 0

def create_excel(df):
    numeric_cols = ['Tổng chỉ số', 'Số tiền', 'Thuế VAT', 'Số tiền dự kiến']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
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

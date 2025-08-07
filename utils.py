import pandas as pd
from io import BytesIO
from docx import Document
import re

def extract_data_from_pdf(text):
    data = {}

    # Số hóa đơn (No)
    match_invoice = re.search(r'Số \(No\):\s*(\d+)', text)
    data['Số hóa đơn'] = match_invoice.group(1) if match_invoice else ''

    # Mã khách hàng (Customer's Code) và lấy Mã tỉnh từ đây
    match_customer_code = re.search(r'Mã khách hàng \(Customer\'s Code\):\s*([A-Z0-9]+)', text)
    if match_customer_code:
        code = match_customer_code.group(1)
        data['Mã EVN'] = code
        data['Mã tỉnh'] = 'YBI' if code.startswith('PA100101') else ''
    else:
        data['Mã EVN'] = ''
        data['Mã tỉnh'] = ''

    # Mã tháng (yyyyMM) từ ngày hóa đơn
    match_date = re.search(r'Ngày \(Date\) (\d{1,2}) tháng (\d{1,2}) năm (\d{4})', text)
    if match_date:
        year = match_date.group(3)
        month = match_date.group(2).zfill(2)
        data['Mã tháng (yyyyMM)'] = f"{year}{month}"
    else:
        data['Mã tháng (yyyyMM)'] = ''

    # Kỳ (mặc định 1)
    data['Kỳ'] = '1'

    # Mã CSHT (cố định hoặc thay đổi theo nhu cầu)
    data['Mã CSHT'] = 'CSHT_YBI_00014'

    # Ngày đầu kỳ và ngày cuối kỳ
    match_period = re.search(r'Điện tiêu thụ tháng \d+ năm \d+ từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})', text)
    if match_period:
        data['Ngày đầu kỳ'] = match_period.group(1)
        data['Ngày cuối kỳ'] = match_period.group(2)
    else:
        data['Ngày đầu kỳ'] = ''
        data['Ngày cuối kỳ'] = ''

    # Tổng chỉ số kWh (lấy số trước "kWh")
    match_kwh = re.search(r'kWh\s*([\d.,]+)', text)
    if match_kwh:
        kwh_str = match_kwh.group(1).replace('.', '').replace(',', '')
        data['Tổng chỉ số'] = kwh_str
    else:
        data['Tổng chỉ số'] = ''

    # Số tiền chưa VAT (tìm "Cộng tiền hàng (Total amount)" hoặc "Tổng cộng tiền thanh toán")
    match_amount = re.search(r'Cộng tiền hàng \(Total amount\):\s*([\d.,]+)', text)
    if not match_amount:
        match_amount = re.search(r'Tổng cộng tiền thanh toán\s*:\s*([\d.,]+)', text)
    if match_amount:
        amount_str = match_amount.group(1).replace('.', '').replace(',', '')
        data['Số tiền'] = amount_str
    else:
        data['Số tiền'] = '0'

    # Thuế VAT
    match_vat = re.search(r'Tiền thuế GTGT \(VAT amount\):\s*([\d.,]+)', text)
    if match_vat:
        vat_str = match_vat.group(1).replace('.', '').replace(',', '')
        data['Thuế VAT'] = vat_str
    else:
        data['Thuế VAT'] = '0'

    # Số tiền dự kiến = Số tiền + Thuế VAT
    try:
        s = int(data['Số tiền'])
        v = int(data['Thuế VAT'])
        data['Số tiền dự kiến'] = str(s + v)
    except:
        data['Số tiền dự kiến'] = '0'

    # Ghi chú (mặc định để trống)
    data['Ghi chú'] = ''

    return data

def create_excel(df):
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

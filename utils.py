import pandas as pd
from io import BytesIO
from docx import Document
import re

def extract_data_from_pdf(text):
    data = {}

    # Ví dụ trích xuất Số hóa đơn
    match_invoice = re.search(r'Số \(No\):\s*(\d+)', text)
    data['Số hóa đơn'] = match_invoice.group(1) if match_invoice else ''

    # Trích xuất Mã tỉnh (ví dụ giả định là YBI, có thể mở rộng regex nếu cần)
    data['Mã tỉnh'] = 'YBI'

    # Mã tháng lấy theo định dạng trong file hoặc mặc định
    match_month = re.search(r'Ngày.*?(\d{2}) tháng (\d{2}) năm (\d{4})', text)
    if match_month:
        year = match_month.group(3)
        month = match_month.group(2)
        data['Mã tháng (yyyyMM)'] = f"{year}{month}"
    else:
        data['Mã tháng (yyyyMM)'] = '202507'  # Mặc định nếu không tìm thấy

    data['Kỳ'] = '1'  # Mặc định

    # Mã CSHT giả định hoặc trích xuất theo quy tắc riêng
    data['Mã CSHT'] = 'CSHT_YBI_00014'

    # Mã EVN - giả định hoặc có thể lấy từ file, hiện để cố định
    data['Mã EVN'] = 'PA10010142348'

    # Ngày đầu kỳ và cuối kỳ trích xuất hoặc mặc định
    match_date_range = re.search(r'từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})', text)
    if match_date_range:
        data['Ngày đầu kỳ'] = match_date_range.group(1)
        data['Ngày cuối kỳ'] = match_date_range.group(2)
    else:
        data['Ngày đầu kỳ'] = '22/06/2025'
        data['Ngày cuối kỳ'] = '21/07/2025'

    # Tổng chỉ số (số kWh)
    match_total_kwh = re.search(r'(\d+)[ ]*kWh', text)
    data['Tổng chỉ số'] = match_total_kwh.group(1) if match_total_kwh else ''

    # Số tiền chưa VAT
    match_amount = re.search(r'Tổng cộng tiền thanh toán\s*:\s*([\d.,]+)', text)
    if not match_amount:
        match_amount = re.search(r'Thành tiền\s*:\s*([\d.,]+)', text)
    if match_amount:
        amount_str = match_amount.group(1).replace('.', '').replace(',', '')
        data['Số tiền'] = amount_str
    else:
        data['Số tiền'] = ''

    # Thuế VAT
    match_vat = re.search(r'Tiền thuế GTGT\s*:\s*([\d.,]+)', text)
    if match_vat:
        vat_str = match_vat.group(1).replace('.', '').replace(',', '')
        data['Thuế VAT'] = vat_str
    else:
        data['Thuế VAT'] = ''

    # Số tiền dự kiến = Số tiền + Thuế VAT (cộng kiểu số)
    try:
        s = int(data['Số tiền'])
        v = int(data['Thuế VAT'])
        data['Số tiền dự kiến'] = str(s + v)
    except:
        data['Số tiền dự kiến'] = ''

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


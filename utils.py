import pandas as pd
from io import BytesIO
from docx import Document
import re

def extract_data_from_pdf(text):
    data = {}

    # Ví dụ trích xuất Số hóa đơn
    match = re.search(r'Số \(No\): (\d+)', text)
    data['Số hóa đơn'] = match.group(1) if match else ''

    # Các trường tạm thời lấy cố định theo mẫu bạn cung cấp
    data.update({
        'Mã tỉnh': 'YBI',
        'Mã tháng (yyyyMM)': '202507',
        'Kỳ': '1',
        'Mã CSHT': 'CSHT_YBI_00014',
        'Mã EVN': 'PA10010142348',
        'Ngày đầu kỳ': '22/06/2025',
        'Ngày cuối kỳ': '21/07/2025',
        'Tổng chỉ số': '1689',
        'Số tiền': '3356043',
        'Thuế VAT': '268483',
        'Số tiền dự kiến': str(3356043 + 268483),
        'Ghi chú': ''
    })

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

    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(df.columns):
            row_cells[i].text = str(row[col_name])
    
    output = BytesIO()
    doc.save(output)
    return output.getvalue()

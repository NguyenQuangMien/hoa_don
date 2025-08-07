import re
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def extract_data_from_pdf(pdf_file):
    # Hàm này gọi pdfplumber để trích xuất dữ liệu chi tiết từ file PDF hóa đơn.
    import pdfplumber
    data = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            # Cần parse text để lấy thông tin cần thiết, ví dụ:
            ma_tinh = "YBI"  # Luôn cố định
            # Giả sử parse các phần mã EVN, số hóa đơn, tổng chỉ số,... từ text
            # Đoạn code cụ thể tùy theo cấu trúc hóa đơn
            # Ví dụ giả định:
            so_hoa_don = "19098"
            ma_evn = "PA10010142348"
            ma_thang = "202507"
            ky = "1"
            ma_csht = "CSHT_YBI_00014"
            ngay_dau_ky = "22/06/2025"
            ngay_cuoi_ky = "21/07/2025"
            tong_chi_so = 1.689
            so_tien = 3356043
            thue_vat = 268483
            so_tien_du_kien = so_tien + thue_vat

            data.append({
                "Mã tỉnh": ma_tinh,
                "Số hóa đơn": so_hoa_don,
                "Mã EVN": ma_evn,
                "Mã tháng (yyyyMM)": ma_thang,
                "Kỳ": ky,
                "Mã CSHT": ma_csht,
                "Ngày đầu kỳ": ngay_dau_ky,
                "Ngày cuối kỳ": ngay_cuoi_ky,
                "Tổng chỉ số": tong_chi_so,
                "Số tiền": so_tien,
                "Thuế VAT": thue_vat,
                "Số tiền dự kiến": so_tien_du_kien,
                "Ghi chú": ""
            })
    return pd.DataFrame(data)

def extract_period(text):
    normalized = re.sub(r'\s+', ' ', text)
    pattern = r'Điện tiêu thụ tháng (\d{1,2}) năm (\d{4}) từ ngày (\d{1,2}/\d{1,2}/\d{4}) đến ngày (\d{1,2}/\d{1,2}/\d{4})'
    match = re.search(pattern, normalized, re.IGNORECASE)
    if match:
        month = match.group(1).zfill(2)
        year = match.group(2)
        ngay_dau = match.group(3)
        ngay_cuoi = match.group(4)
        print(f"Extracted period: {year}{month}, start: {ngay_dau}, end: {ngay_cuoi}")
        return f"{year}{month}", ngay_dau, ngay_cuoi
    else:
        print("Period pattern not found")
        return None, None, None

def create_excel(df):
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Hóa đơn"

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
        if r_idx == 1:
            # Định dạng header bold và căn giữa
            for cell in ws[r_idx]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Định dạng tất cả cột dạng text
    for col in ws.columns:
        for cell in col:
            cell.number_format = '@'  # Text format

    # Tự động điều chỉnh độ rộng cột
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
        adjusted_width = length + 2
        ws.column_dimensions[column_cells[0].column_letter].width = adjusted_width

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def create_word(df):
    doc = Document()
    doc.add_heading('Bảng dữ liệu hóa đơn', level=1)

    table = doc.add_table(rows=1, cols=len(df.columns))
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = str(col_name)

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

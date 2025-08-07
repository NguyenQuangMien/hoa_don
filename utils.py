import re
import json
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn


def extract_data_from_pdf(pdf_file):
    # Đọc nội dung PDF thành chuỗi text (cần có hàm đọc PDF riêng)
    # Giả sử pdfplumber đã được dùng ở app.py, truyền text vào đây
    import pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    data = {
        "Mã tỉnh": "YBI",  # Luôn cố định
        "Số hóa đơn": "",
        "Mã EVN": "",
        "Mã tháng (yyyyMM)": "",
        "Kỳ": "1",
        "Mã CSHT": "",
        "Ngày đầu kỳ": "",
        "Ngày cuối kỳ": "",
        "Tổng chỉ số": "",
        "Số tiền": "",
        "Thuế VAT": "",
        "Số tiền dự kiến": "",
        "Ghi chú": ""
    }

    # Số hóa đơn
    m_sohd = re.search(r'Số hóa đơn\s*:\s*(\d+)', text)
    if m_sohd:
        data["Số hóa đơn"] = m_sohd.group(1).strip()

    # Mã EVN
    m_maevn = re.search(r'Mã khách hàng \(Customer\'s Code\)\s*:\s*(\S+)', text)
    if m_maevn:
        data["Mã EVN"] = m_maevn.group(1).strip()

    # Mã tháng (yyyyMM) - lấy từ tên file hoặc trong text
    m_matham = re.search(r'(\d{6})', pdf_file.name)
    if m_matham:
        data["Mã tháng (yyyyMM)"] = m_matham.group(1)

    # Mã CSHT (chỉnh sửa chính xác, lấy đúng mã không lấy địa chỉ)
    m_csht = re.search(r'Mã số đơn vị có quan hệ với ngân sách \(State budget related unit code\):\s*(\S+)', text)
    if m_csht:
        data["Mã CSHT"] = m_csht.group(1).strip()

    # Ngày đầu kỳ và ngày cuối kỳ lấy từ phần mô tả tiêu thụ điện
    # Ví dụ: Điện tiêu thụ tháng 7 năm 2025 từ ngày 22/06/2025 đến ngày 21/07/2025
    m_ngay = re.search(
        r'Điện tiêu thụ tháng \d+ năm \d{4} từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})', text)
    if m_ngay:
        data["Ngày đầu kỳ"] = m_ngay.group(1)
        data["Ngày cuối kỳ"] = m_ngay.group(2)

    # Tổng chỉ số (lấy ở bảng bên dưới)
    m_chiso = re.search(r'Cộng tiền hàng.*?\n([\d\.,]+)', text, re.DOTALL)
    if m_chiso:
        chiso = m_chiso.group(1).strip().replace(",", "")
        data["Tổng chỉ số"] = chiso

    # Số tiền
    m_sotien = re.search(r'Cộng tiền hàng.*?([\d\.,]+)', text)
    if m_sotien:
        data["Số tiền"] = m_sotien.group(1).strip()

    # Thuế VAT
    m_vat = re.search(r'Tiền thuế GTGT.*?([\d\.,]+)', text)
    if m_vat:
        data["Thuế VAT"] = m_vat.group(1).strip()

    # Số tiền dự kiến = Tổng cộng tiền thanh toán
    m_tiendukien = re.search(r'Tổng cộng tiền thanh toán.*?([\d\.,]+)', text)
    if m_tiendukien:
        data["Số tiền dự kiến"] = m_tiendukien.group(1).strip()

    return data


def create_excel(data_list):
    # Tạo Excel từ danh sách dict data_list
    output = BytesIO()
    df = pd.DataFrame(data_list)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Hóa đơn')
        workbook = writer.book
        worksheet = writer.sheets['Hóa đơn']

        # Định dạng cho các cột: tự động rộng, kiểu text
        text_format = workbook.add_format({'num_format': '@'})
        for idx, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet.set_column(idx, idx, max_len, text_format)

    output.seek(0)
    return output.read()


def create_word(data_list):
    # Tạo file Word từ data_list
    document = Document()

    # Cài font Times New Roman chuẩn Unicode
    style = document.styles['Normal']
    style.font.name = 'Times New Roman'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    style.font.size = Pt(14)

    # Tiêu đề
    document.add_heading('Dữ liệu hóa đơn', level=1)

    # Tạo bảng với số cột = số trường
    if len(data_list) == 0:
        document.add_paragraph('Không có dữ liệu.')
    else:
        keys = list(data_list[0].keys())
        table = document.add_table(rows=1, cols=len(keys))
        hdr_cells = table.rows[0].cells
        for i, key in enumerate(keys):
            hdr_cells[i].text = key

        for item in data_list:
            row_cells = table.add_row().cells
            for i, key in enumerate(keys):
                row_cells[i].text = str(item.get(key, ''))

    word_buffer = BytesIO()
    document.save(word_buffer)
    word_buffer.seek(0)
    return word_buffer.read()

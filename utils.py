import re
import pandas as pd
from io import BytesIO
from pdfplumber import open as pdf_open
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def extract_data_from_pdf(pdf_bytes):
    data = {
        "Mã tỉnh": "YBI",  # Luôn cố định
        "Số hóa đơn": "",
        "Mã EVN": "",
        "Mã tháng (yyyyMM)": "",
        "Kỳ": "1",  # Cố định
        "Mã CSHT": "CSHT_YBI_00014",  # Cố định
        "Ngày đầu kỳ": "",
        "Ngày cuối kỳ": "",
        "Tổng chỉ số": "",
        "Số tiền": "",
        "Thuế VAT": "",
        "Số tiền dự kiến": "",
        "Ghi chú": ""
    }

    with pdf_open(BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()

        # In ra dữ liệu thô để debug
        print("Dữ liệu thô trích xuất:")
        print(text)

        # Số hóa đơn
        so_hd_match = re.search(r"Số\s*\(No\):\s*(\d+)", text)
        if so_hd_match:
            data["Số hóa đơn"] = so_hd_match.group(1).strip()

        # Mã EVN (Mã khách hàng)
        ma_evn_match = re.search(r"Mã khách hàng \(Customer's Code\):\s*([A-Z0-9]+)", text)
        if ma_evn_match:
            data["Mã EVN"] = ma_evn_match.group(1).strip()

        # Mã tháng yyyyMM lấy từ tên file hoặc giả định tháng 07 năm 2025
        # (Ở đây bạn có thể bổ sung nếu muốn tự động lấy từ file)
        data["Mã tháng (yyyyMM)"] = "202507"

        # Ngày đầu kỳ, ngày cuối kỳ lấy từ mô tả điện tiêu thụ
        date_range_match = re.search(
            r"Điện tiêu thụ tháng \d+ năm (\d{4}) từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})", text)
        if date_range_match:
            data["Ngày đầu kỳ"] = date_range_match.group(2)
            data["Ngày cuối kỳ"] = date_range_match.group(3)

        # Tổng chỉ số, Số tiền, Thuế VAT, Số tiền dự kiến
        # Lấy trong bảng mô tả
        # Ví dụ tìm "Cộng tiền hàng (Total amount):"
        tong_chiso_match = re.search(r"Total amount\):\s*([\d.,]+)", text)
        if tong_chiso_match:
            data["Tổng chỉ số"] = tong_chiso_match.group(1).strip().replace(",", "")

        so_tien_match = re.search(r"Cộng tiền hàng .*?([\d.,]+)", text)
        if so_tien_match:
            data["Số tiền"] = so_tien_match.group(1).strip().replace(",", "")

        thue_vat_match = re.search(r"Tiền thuế GTGT .*?([\d.,]+)", text)
        if thue_vat_match:
            data["Thuế VAT"] = thue_vat_match.group(1).strip().replace(",", "")

        so_tien_du_kien_match = re.search(r"Tổng cộng tiền thanh toán .*?([\d.,]+)", text)
        if so_tien_du_kien_match:
            data["Số tiền dự kiến"] = so_tien_du_kien_match.group(1).strip().replace(",", "")

    return data


def create_excel(data_list):
    df = pd.DataFrame(data_list)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Hóa đơn')
        worksheet = writer.sheets['Hóa đơn']

        # Tự động chỉnh độ rộng cột và định dạng text
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
            # Định dạng là text (số dạng chuỗi)
            worksheet.set_column(i, i, max_len, writer.book.add_format({'num_format': '@'}))

    output.seek(0)
    return output.read()


def create_word(data_list):
    document = Document()

    style = document.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    for data in data_list:
        document.add_paragraph("Thông tin hóa đơn:")
        for k, v in data.items():
            document.add_paragraph(f"{k}: {v}")
        document.add_paragraph("\n")

    word_buffer = BytesIO()
    document.save(word_buffer)
    word_buffer.seek(0)
    return word_buffer.read()

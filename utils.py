import re
import pandas as pd
from io import BytesIO
from docx import Document
from openpyxl import Workbook
from pdfplumber import open as pdf_open

def extract_period(text):
    normalized = re.sub(r'\s+', ' ', text)
    pattern = r'Điện tiêu thụ tháng (\d{1,2}) năm (\d{4}) từ ngày (\d{1,2}/\d{1,2}/\d{4}) đến ngày (\d{1,2}/\d{1,2}/\d{4})'
    match = re.search(pattern, normalized, re.IGNORECASE)
    if match:
        month = match.group(1).zfill(2)
        year = match.group(2)
        ngay_dau = match.group(3)
        ngay_cuoi = match.group(4)
        return f"{year}{month}", ngay_dau, ngay_cuoi
    else:
        return None, None, None

def extract_data_from_pdf(pdf_file):
    data = {
        "Mã tỉnh": "YBI",
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

    with pdf_open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

        # Số hóa đơn (No)
        so_hoa_don_match = re.search(r'Số\s*\(?No\.?\)?:?\s*(\d+)', text, re.IGNORECASE)
        if so_hoa_don_match:
            data["Số hóa đơn"] = so_hoa_don_match.group(1).strip()

        # Mã khách hàng (Customer's Code) = Mã EVN
        ma_khach_hang_match = re.search(r'Mã khách hàng\s*\(?Customer\'s Code\)?:?\s*(\S+)', text)
        if ma_khach_hang_match:
            data["Mã EVN"] = ma_khach_hang_match.group(1).strip()

        # Trích xuất tháng và ngày đầu, ngày cuối kỳ
        ma_thang, ngay_dau, ngay_cuoi = extract_period(text)
        if ma_thang:
            data["Mã tháng (yyyyMM)"] = ma_thang
        if ngay_dau:
            data["Ngày đầu kỳ"] = ngay_dau
        if ngay_cuoi:
            data["Ngày cuối kỳ"] = ngay_cuoi

        # Trích xuất bảng dữ liệu kWh, tiền, thuế VAT
        # Cách đơn giản: tìm dòng "Điện tiêu thụ tháng ..." và lấy các số theo cột kWh và Thành tiền
        lines = text.split('\n')
        for idx, line in enumerate(lines):
            if "Điện tiêu thụ tháng" in line:
                # Lấy dòng hiện tại và dòng tiếp theo (do có dòng phụ)
                line_1 = lines[idx]
                line_2 = lines[idx + 1] if idx + 1 < len(lines) else ''

                # Lấy số lượng kWh từ dòng 1 (giá trị đầu tiên dạng số thực)
                khw_match = re.search(r'(\d+\.\d+)', line_1)
                if khw_match:
                    data["Tổng chỉ số"] = khw_match.group(1)

                # Lấy thành tiền trong dòng kế hoặc dòng hiện tại
                # Thành tiền là số lớn có dấu chấm phân cách hàng nghìn
                thanh_tien_match_1 = re.findall(r'(\d{1,3}(?:\.\d{3})+)', line_1)
                thanh_tien_match_2 = re.findall(r'(\d{1,3}(?:\.\d{3})+)', line_2)

                if thanh_tien_match_2:
                    data["Số tiền"] = thanh_tien_match_2[-1]
                elif thanh_tien_match_1:
                    data["Số tiền"] = thanh_tien_match_1[-1]

        # Thuế VAT và Tổng cộng tiền thanh toán
        # Tìm dòng có "Tiền thuế GTGT" và "Tổng cộng tiền thanh toán"
        tien_thue_match = re.search(r'Tiền thuế GTGT \(VAT amount\):\s*([\d\.]+)', text)
        if tien_thue_match:
            data["Thuế VAT"] = tien_thue_match.group(1)

        tong_cong_match = re.search(r'Tổng cộng tiền thanh toán \(Total payment\):\s*([\d\.]+)', text)
        if tong_cong_match:
            data["Số tiền dự kiến"] = tong_cong_match.group(1)

    return data

def create_excel(data_list):
    df = pd.DataFrame(data_list)

    # Chuyển tất cả thành string để upload lên hệ thống và định dạng là text
    df = df.fillna("")
    for col in df.columns:
        df[col] = df[col].astype(str)

    output = BytesIO()
    df.to_excel(output, index=False, sheet_name="Hóa đơn")
    output.seek(0)
    return output

def create_word(data_list):
    doc = Document()
    doc.add_heading('Bảng dữ liệu hóa đơn trích xuất', level=1)
    for data in data_list:
        for k, v in data.items():
            doc.add_paragraph(f"{k}: {v}")
        doc.add_paragraph("\n")
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

import re
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

def extract_data_from_pdf(pdf_file):
    from pdfplumber import open as pdf_open

    data_list = []
    with pdf_open(pdf_file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()

        # Mã tỉnh cố định
        ma_tinh = "YBI"

        # Mã EVN: tìm mẫu PA + số, ví dụ PA10010142348
        ma_evn_match = re.search(r"(PA\d{8,})", text)
        ma_evn = ma_evn_match.group(1) if ma_evn_match else ""

        # Mã tháng (yyyyMM) lấy từ tên file hoặc ngày hiện tại (cần xử lý ngoài)
        # Ở đây tạm lấy mặc định hoặc cần truyền vào
        ma_thang = ""

        # Kỳ cố định 1
        ky = "1"

        # Mã CSHT có dạng CSHT_YBI_xxxx, cố định hoặc tìm kiếm
        csht_match = re.search(r"(CSHT_YBI_\d{5})", text)
        ma_csht = csht_match.group(1) if csht_match else ""

        # Ngày đầu kỳ và ngày cuối kỳ: lấy từ đoạn "Điện tiêu thụ tháng 7 năm 2025 từ ngày 22/06/2025 đến ngày 21/07/2025"
        date_match = re.search(
            r"Điện tiêu thụ tháng \d+ năm \d{4} từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})",
            text
        )
        ngay_dau_ky = date_match.group(1) if date_match else ""
        ngay_cuoi_ky = date_match.group(2) if date_match else ""

        # Tổng chỉ số, Số tiền, Thuế VAT, Số tiền dự kiến lấy từ bảng hoặc từ đoạn text
        # Cố gắng tìm mẫu số lượng (Tổng chỉ số)
        tong_chi_so_match = re.search(r"(\d+\.\d+|\d+)\s+kWh", text)
        tong_chi_so = tong_chi_so_match.group(1) if tong_chi_so_match else ""

        # Số tiền (Cộng tiền hàng) - tìm mẫu số lớn dạng số có dấu chấm hàng nghìn
        so_tien_match = re.search(r"Cộng tiền hàng.*?([\d\.]+)", text)
        so_tien = so_tien_match.group(1) if so_tien_match else ""

        # Thuế VAT (Tiền thuế GTGT)
        thue_vat_match = re.search(r"Tiền thuế GTGT.*?([\d\.]+)", text)
        thue_vat = thue_vat_match.group(1) if thue_vat_match else ""

        # Số tiền dự kiến = Tổng cộng tiền thanh toán (Tổng cộng tiền thanh toán)
        so_tien_du_kien_match = re.search(r"Tổng cộng tiền thanh toán.*?([\d\.]+)", text)
        so_tien_du_kien = so_tien_du_kien_match.group(1) if so_tien_du_kien_match else ""

        # Số hóa đơn: tìm dòng có "Số hóa đơn" hoặc lấy từ tên file (nếu cần)
        so_hoa_don_match = re.search(r"Số hóa đơn.*?(\d+)", text)
        so_hoa_don = so_hoa_don_match.group(1) if so_hoa_don_match else ""

        # Tạo dictionary dữ liệu
        data = {
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
        }
        data_list.append(data)

    return data_list

def create_excel(data_list):
    df = pd.DataFrame(data_list)
    # Định dạng các cột thành text để tránh mất số 0 đầu hoặc chuyển đổi sai
    for col in df.columns:
        df[col] = df[col].astype(str)
    output = BytesIO()
    df.to_excel(output, index=False, sheet_name="Hóa đơn")
    output.seek(0)
    return output

def create_word(data_list):
    document = Document()
    document.styles['Normal'].font.name = 'Times New Roman'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    document.styles['Normal'].font.size = Pt(12)

    document.add_heading('Bảng dữ liệu hóa đơn trích xuất', level=1)

    if not data_list:
        document.add_paragraph('Không có dữ liệu.')
        output = BytesIO()
        document.save(output)
        output.seek(0)
        return output

    df = pd.DataFrame(data_list)
    table = document.add_table(rows=1, cols=len(df.columns))
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = col_name

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(df.columns):
            row_cells[i].text = str(row[col_name])

    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output

import pandas as pd
import io
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from zipfile import ZipFile
import pdfplumber
import re

def extract_data_from_pdf(file):
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

    with pdfplumber.open(file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()

        # Trích xuất Số hóa đơn
        so_hd_match = re.search(r"Số\s*:\s*(\d+)", text)
        if so_hd_match:
            data["Số hóa đơn"] = so_hd_match.group(1).strip()

        # Trích xuất Mã EVN
        ma_kh_match = re.search(r"Mã khách hàng\s*\(Customer's Code\):\s*([A-Z0-9]+)", text)
        if ma_kh_match:
            data["Mã EVN"] = ma_kh_match.group(1).strip()

        # Trích xuất Mã tháng
        ma_thang_match = re.search(r"(\d{6})", file.name)
        if ma_thang_match:
            data["Mã tháng (yyyyMM)"] = ma_thang_match.group(1)

        # Trích xuất Ngày đầu kỳ và Ngày cuối kỳ từ chuỗi 'Điện tiêu thụ tháng ...'
        match_date_range = re.search(
            r"Điện tiêu thụ tháng \d+ năm \d+ từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})", text
        )
        if match_date_range:
            data["Ngày đầu kỳ"] = match_date_range.group(1)
            data["Ngày cuối kỳ"] = match_date_range.group(2)

        # Trích xuất Tổng chỉ số (số lượng kWh)
        quantity_match = re.search(r"kWh\s+([\d\.]+)", text)
        if quantity_match:
            data["Tổng chỉ số"] = quantity_match.group(1)

        # Trích xuất Số tiền
        so_tien_match = re.search(r"Cộng tiền hàng.*?([\d\.]+)", text, re.DOTALL)
        if so_tien_match:
            data["Số tiền"] = so_tien_match.group(1).replace(".", "")

        # Trích xuất Thuế VAT
        vat_match = re.search(r"Tiền thuế GTGT.*?([\d\.]+)", text, re.DOTALL)
        if vat_match:
            data["Thuế VAT"] = vat_match.group(1).replace(".", "")

        # Trích xuất Số tiền dự kiến (Tổng cộng tiền thanh toán)
        tien_du_kien_match = re.search(r"Tổng cộng tiền thanh toán.*?([\d\.]+)", text, re.DOTALL)
        if tien_du_kien_match:
            data["Số tiền dự kiến"] = tien_du_kien_match.group(1).replace(".", "")

    return data

def create_excel(data_list):
    df = pd.DataFrame(data_list)

    # Định dạng cột ngày tháng là text (chuỗi)
    for col in ["Ngày đầu kỳ", "Ngày cuối kỳ"]:
        if col in df.columns:
            df[col] = df[col].astype(str)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Hóa đơn")
        worksheet = writer.sheets["Hóa đơn"]
        worksheet.set_column("A:Z", 20)  # Tự động giãn cột

    return output.getvalue()

def create_word(data_list):
    document = Document()

    # Thiết lập font Times New Roman cho toàn bộ văn bản
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # Tạo bảng với số cột = số trường dữ liệu
    if data_list:
        keys = list(data_list[0].keys())
        table = document.add_table(rows=1, cols=len(keys))
        hdr_cells = table.rows[0].cells

        # Header
        for i, key in enumerate(keys):
            hdr_cells[i].text = key

        # Nội dung
        for data in data_list:
            row_cells = table.add_row().cells
            for i, key in enumerate(keys):
                row_cells[i].text = str(data.get(key, ""))

    word_buffer = io.BytesIO()
    document.save(word_buffer)
    word_buffer.seek(0)
    return word_buffer.getvalue()

import re
import pdfplumber
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def extract_data_from_pdf(file):
    data = {
        "Mã tỉnh": "YBI",  # cố định
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
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        
        # Số hóa đơn
        match_sohd = re.search(r"Số hóa đơn\s*:\s*(\d+)", text)
        if match_sohd:
            data["Số hóa đơn"] = match_sohd.group(1).strip()

        # Mã EVN (Mã khách hàng)
        match_maevn = re.search(r"Mã khách hàng\s*:\s*([A-Z0-9]+)", text)
        if match_maevn:
            data["Mã EVN"] = match_maevn.group(1).strip()

        # Mã tháng yyyyMM (lấy từ tiêu đề file hoặc chuỗi "Điện tiêu thụ tháng ... năm ...")
        match_month = re.search(r"Điện tiêu thụ tháng\s*(\d+)\s*năm\s*(\d{4})", text)
        if match_month:
            month = int(match_month.group(1))
            year = int(match_month.group(2))
            data["Mã tháng (yyyyMM)"] = f"{year}{month:02d}"

        # Ngày đầu kỳ và ngày cuối kỳ lấy từ chuỗi
        match_ngay = re.search(r"Điện tiêu thụ tháng \d+ năm \d{4} từ ngày (\d{2}/\d{2}/\d{4}) đến ngày (\d{2}/\d{2}/\d{4})", text)
        if match_ngay:
            data["Ngày đầu kỳ"] = match_ngay.group(1)
            data["Ngày cuối kỳ"] = match_ngay.group(2)

        # Tổng chỉ số, Số tiền, Thuế VAT, Số tiền dự kiến
        match_tongchi = re.search(r"Cộng chỉ số.*?([\d\.]+)", text, re.DOTALL)
        if match_tongchi:
            data["Tổng chỉ số"] = match_tongchi.group(1).replace(".", "")

        match_sotien = re.search(r"Cộng tiền hàng.*?([\d\.]+)", text, re.DOTALL)
        if match_sotien:
            data["Số tiền"] = match_sotien.group(1).strip()

        match_thuevat = re.search(r"Tiền thuế GTGT.*?([\d\.]+)", text, re.DOTALL)
        if match_thuevat:
            data["Thuế VAT"] = match_thuevat.group(1).strip()

        match_tiendukien = re.search(r"Tổng cộng tiền thanh toán.*?([\d\.]+)", text, re.DOTALL)
        if match_tiendukien:
            data["Số tiền dự kiến"] = match_tiendukien.group(1).strip()

    return data


def create_excel(data_list):
    df = pd.DataFrame(data_list)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for idx, col in enumerate(df):
            # Tự động giãn cột
            series = df[col].astype(str)
            max_len = max(series.map(len).max(), len(str(col))) + 2
            worksheet.set_column(idx, idx, max_len)
    return output.getvalue()


def create_word(data_list):
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    # Để hỗ trợ tiếng Việt
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    table = document.add_table(rows=1, cols=len(data_list[0]))
    hdr_cells = table.rows[0].cells
    for i, key in enumerate(data_list[0].keys()):
        hdr_cells[i].text = key

    for data in data_list:
        row_cells = table.add_row().cells
        for i, key in enumerate(data.keys()):
            row_cells[i].text = str(data[key]) if data[key] is not None else ''

    output = BytesIO()
    document.save(output)
    return output.getvalue()

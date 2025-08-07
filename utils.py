import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import pdfplumber
import zipfile

def extract_data_from_pdf(pdf_file):
    data = {}
    with pdfplumber.open(pdf_file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        lines = text.split('\n')

        # Tạm ví dụ: tìm các thông tin theo mẫu cụ thể
        for line in lines:
            if "Mã EVN" in line:
                data["Mã EVN"] = line.split(":")[-1].strip()
            elif "Mã CSHT" in line:
                data["Mã CSHT"] = line.split(":")[-1].strip()
            elif "Mã tháng" in line:
                data["Mã tháng (yyyyMM)"] = line.split(":")[-1].strip()
            elif "Kỳ" in line:
                data["Kỳ"] = line.split(":")[-1].strip()
            elif "Mã tỉnh" in line:
                data["Mã tỉnh"] = line.split(":")[-1].strip()
            elif "Số hóa đơn" in line:
                data["Số hóa đơn"] = line.split(":")[-1].strip()
            elif "Điện tiêu thụ tháng" in line:
                # Lấy ngày đầu kỳ và ngày cuối kỳ theo mẫu "Điện tiêu thụ tháng ... từ ngày dd/MM/yyyy đến ngày dd/MM/yyyy"
                import re
                match = re.search(r'(\d{2}/\d{2}/\d{4}).+(\d{2}/\d{2}/\d{4})', line)
                if match:
                    data["Ngày đầu kỳ"] = match.group(1)
                    data["Ngày cuối kỳ"] = match.group(2)
            elif "Tổng chỉ số" in line:
                data["Tổng chỉ số"] = line.split(":")[-1].strip()
            elif "Số tiền bằng chữ" in line:
                # Lấy tổng số tiền dự kiến (cộng tiền hàng + thuế)
                pass

        # Trích xuất bảng để lấy số tiền, thuế VAT, tiền dự kiến
        table = page.extract_table()
        if table:
            # Giả định dòng cuối hoặc dòng phù hợp có các số tiền
            for row in table:
                # Tìm dòng có 'Cộng tiền hàng' hoặc tương tự để lấy số tiền
                if any("Cộng tiền" in str(cell) for cell in row if cell):
                    idx = row.index(next(cell for cell in row if "Cộng tiền" in str(cell)))
                    data["Số tiền"] = row[idx+1].strip() if idx+1 < len(row) else ""
                if any("Thuế suất GTGT" in str(cell) for cell in row if cell):
                    idx = row.index(next(cell for cell in row if "Thuế suất GTGT" in str(cell)))
                    # Thường ở dòng khác lấy số tiền thuế GTGT
                    # Tìm trong những dòng kế tiếp
                    pass
                if any("Tiền thuế GTGT" in str(cell) for cell in row if cell):
                    idx = row.index(next(cell for cell in row if "Tiền thuế GTGT" in str(cell)))
                    data["Thuế VAT"] = row[idx+1].strip() if idx+1 < len(row) else ""
                if any("Tổng cộng tiền thanh toán" in str(cell) for cell in row if cell):
                    idx = row.index(next(cell for cell in row if "Tổng cộng tiền thanh toán" in str(cell)))
                    data["Số tiền dự kiến"] = row[idx+1].strip() if idx+1 < len(row) else ""

    # Mặc định "Mã tỉnh" luôn là "YBI"
    data["Mã tỉnh"] = "YBI"

    return data

def create_excel(data_list):
    df = pd.DataFrame(data_list)
    writer = BytesIO()
    with pd.ExcelWriter(writer, engine='xlsxwriter') as writer_obj:
        df.to_excel(writer_obj, index=False, sheet_name='Sheet1')
        worksheet = writer_obj.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
        writer_obj.save()
    writer.seek(0)
    return writer.getvalue()

def create_word(data_list):
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    if len(data_list) == 0:
        return None

    table = document.add_table(rows=1, cols=len(data_list[0]))
    hdr_cells = table.rows[0].cells
    for i, key in enumerate(data_list[0].keys()):
        hdr_cells[i].text = key

    for data in data_list:
        row_cells = table.add_row().cells
        for i, key in enumerate(data.keys()):
            row_cells[i].text = str(data[key]) if data[key] is not None else ''

    word_buffer = BytesIO()
    document.save(word_buffer)
    word_buffer.seek(0)
    return word_buffer.getvalue()

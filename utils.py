import re
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn


def extract_data_from_pdf(pdf_bytes):
    # Chuyển pdf bytes sang text (giả định có hàm get_text_from_pdf)
    text = get_text_from_pdf(pdf_bytes)

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

    # Lấy các trường dữ liệu theo mẫu hóa đơn
    data["Số hóa đơn"] = extract_pattern(text, r"Số hóa đơn\s*:\s*(\d+)")
    data["Mã EVN"] = extract_pattern(text, r"Mã khách hàng\s*:\s*(\w+)")
    data["Mã tháng (yyyyMM)"] = extract_pattern(text, r"(\d{6})")  # có thể cần hiệu chỉnh
    data["Mã CSHT"] = extract_pattern(text, r"Mã số đơn vị.*:\s*(\w+)")
    # Mã tỉnh mặc định là YBI đã set ở trên

    # Lấy Tổng chỉ số, Số tiền, Thuế VAT, Số tiền dự kiến
    data["Tổng chỉ số"] = extract_pattern(text, r"Điện tiêu thụ tháng.*?(\d+[\.,]?\d*)\s*kWh")
    data["Số tiền"] = extract_pattern(text, r"Cộng tiền hàng.*?([\d\.]+)")
    data["Thuế VAT"] = extract_pattern(text, r"Tiền thuế GTGT.*?([\d\.]+)")
    data["Số tiền dự kiến"] = extract_pattern(text, r"Tổng cộng tiền thanh toán.*?([\d\.]+)")

    # Lấy ngày đầu kỳ, ngày cuối kỳ từ chuỗi ngày tháng trong mô tả điện tiêu thụ
    date_text_match = re.search(r"Điện tiêu thụ tháng \d+ năm \d{4} từ ngày .*? đến ngày .*", text)
    if date_text_match:
        date_text = date_text_match.group(0)
        print("Chuỗi ngày tháng tìm thấy:", date_text)
        day_start_match = re.search(r"từ ngày (\d{2}/\d{2}/\d{4})", date_text)
        day_end_match = re.search(r"đến ngày (\d{2}/\d{2}/\d{4})", date_text)
        if day_start_match:
            data["Ngày đầu kỳ"] = day_start_match.group(1)
        if day_end_match:
            data["Ngày cuối kỳ"] = day_end_match.group(1)

    return data


def extract_pattern(text, pattern):
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1).strip()
    return ""


def create_excel(data_list):
    df = pd.DataFrame(data_list)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(col)
            )
            worksheet.set_column(i, i, max_len + 2)
    output.seek(0)
    return output


def create_word(data_list):
    document = Document()
    document.styles['Normal'].font.name = 'Times New Roman'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    document.styles['Normal'].font.size = Pt(14)

    table = document.add_table(rows=1, cols=len(data_list[0]))
    hdr_cells = table.rows[0].cells
    for i, key in enumerate(data_list[0].keys()):
        hdr_cells[i].text = key

    for data in data_list:
        row_cells = table.add_row().cells
        for i, value in enumerate(data.values()):
            row_cells[i].text = str(value)

    word_buffer = BytesIO()
    document.save(word_buffer)
    word_buffer.seek(0)
    return word_buffer


# Hàm giả lập đọc PDF thành text, bạn thay bằng thư viện đọc thực tế
def get_text_from_pdf(pdf_bytes):
    # TODO: sử dụng pdfplumber hoặc PyPDF2 để đọc text
    return pdf_bytes.decode('utf-8', errors='ignore')  # giả lập


if __name__ == "__main__":
    # Test nhanh hàm
    with open("CSHT_YBI_00014_07-2025.pdf", "rb") as f:
        pdf_bytes = f.read()
    data = extract_data_from_pdf(pdf_bytes)
    print(data)

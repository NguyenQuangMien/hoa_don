import re
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import pdfplumber

def extract_data_from_pdf(file):
    data = {
        "Mã tỉnh": "YBI",
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
    try:
        with pdfplumber.open(file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

            # Lấy Mã EVN
            m_ev = re.search(r'Mã khách hàng \(Customer\'s Code\):\s*(\S+)', text)
            if m_ev:
                data["Mã EVN"] = m_ev.group(1).strip()

            # Lấy Mã CSHT
            m_csht = re.search(r'Tên đơn vị.*?\n.*?Mã số đơn vị có quan hệ với ngân sách.*?\n(.*?)\n', text, re.DOTALL)
            if m_csht:
                csht_line = m_csht.group(1).strip()
                data["Mã CSHT"] = csht_line

            # Lấy Mã tháng yyyyMM
            m_month = re.search(r'Điện tiêu thụ tháng\s*(\d{1,2})\s*năm\s*(\d{4})', text)
            if m_month:
                mm = int(m_month.group(1))
                yyyy = int(m_month.group(2))
                data["Mã tháng (yyyyMM)"] = f"{yyyy}{mm:02d}"

            # Lấy Số hóa đơn
            m_sod = re.search(r'Số hóa đơn\s*:\s*(\d+)', text)
            if m_sod:
                data["Số hóa đơn"] = m_sod.group(1).strip()

            # Lấy Tổng chỉ số, Số tiền, Thuế VAT, Số tiền dự kiến
            m_table = re.search(
                r'Điện tiêu thụ tháng .*? từ ngày (.*?) đến ngày (.*?)\s.*?(\d[\d\.]*)\s*.*?([\d\.]+)\s*([\d\.]+)\s*([\d\.]+)',
                text, re.DOTALL)
            if m_table:
                data["Ngày đầu kỳ"] = m_table.group(1).strip()
                data["Ngày cuối kỳ"] = m_table.group(2).strip()
                data["Tổng chỉ số"] = m_table.group(3).strip()
                data["Số tiền"] = m_table.group(4).strip()
                data["Thuế VAT"] = m_table.group(5).strip()
                data["Số tiền dự kiến"] = m_table.group(6).strip()

            # Ghi chú để trống
            data["Ghi chú"] = ""

    except Exception as e:
        print(f"Lỗi đọc PDF: {e}")

    return data


def create_excel(data_list, return_bytes=False):
    df = pd.DataFrame(data_list)
    # Đặt kiểu dữ liệu tất cả cột thành chuỗi để giữ định dạng
    for col in df.columns:
        df[col] = df[col].astype(str)
    # Tạo bộ nhớ đệm
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Hoá đơn')
        workbook = writer.book
        worksheet = writer.sheets['Hoá đơn']

        # Tự động điều chỉnh độ rộng cột
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)
    if return_bytes:
        return output.getvalue()
    else:
        return df


def create_word(data_list, return_bytes=False):
    document = Document()
    # Thiết lập font Times New Roman
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    rFonts = style.element.rPr.rFonts
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # Thêm bảng
    headers = ["Mã tỉnh", "Số hóa đơn", "Mã EVN", "Mã tháng (yyyyMM)", "Kỳ",
               "Mã CSHT", "Ngày đầu kỳ", "Ngày cuối kỳ", "Tổng chỉ số", "Số tiền", "Thuế VAT", "Số tiền dự kiến", "Ghi chú"]
    table = document.add_table(rows=1, cols=len(headers))
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header

    for data in data_list:
        row_cells = table.add_row().cells
        for i, key in enumerate(headers):
            val = data.get(key, "")
            row_cells[i].text = val

    # Lưu file Word vào bộ nhớ đệm
    output = BytesIO()
    document.save(output)
    if return_bytes:
        return output.getvalue()
    else:
        return document

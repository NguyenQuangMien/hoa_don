import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

from utils import extract_data_from_pdf, create_excel, create_word

st.title("Ứng dụng trích xuất dữ liệu hóa đơn PDF")

uploaded_files = st.file_uploader(
    "Chọn nhiều file PDF hoặc file ZIP chứa PDF",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

data_list = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".zip"):
            # Giải nén và xử lý từng file PDF trong ZIP
            with zipfile.ZipFile(uploaded_file) as z:
                for filename in z.namelist():
                    if filename.lower().endswith(".pdf"):
                        with z.open(filename) as f:
                            data = extract_data_from_pdf(f)
                            if data:
                                data_list.append(data)
        elif uploaded_file.name.endswith(".pdf"):
            data = extract_data_from_pdf(uploaded_file)
            if data:
                data_list.append(data)

    if data_list:
        df = pd.DataFrame(data_list)
        st.dataframe(df)

        # Xuất file Excel
        excel_bytes = create_excel(data_list)
        st.download_button(
            label="Tải file Excel",
            data=excel_bytes,
            file_name="hoa_don.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Xuất file Word
        word_bytes = create_word(data_list)
        st.download_button(
            label="Tải file Word",
            data=word_bytes,
            file_name="hoa_don.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

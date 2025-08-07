import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

from utils import extract_data_from_pdf, create_excel, create_word

st.title("Xử lý nhiều hóa đơn PDF hoặc ZIP")

uploaded_files = st.file_uploader(
    "Chọn nhiều file PDF hoặc file ZIP (chứa PDF)",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

data_list = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".zip"):
            # Giải nén zip và xử lý từng file PDF bên trong
            with zipfile.ZipFile(uploaded_file) as z:
                for filename in z.namelist():
                    if filename.endswith(".pdf"):
                        with z.open(filename) as pdf_file:
                            data = extract_data_from_pdf(pdf_file)
                            if data:
                                data_list.append(data)
        elif uploaded_file.name.endswith(".pdf"):
            data = extract_data_from_pdf(uploaded_file)
            if data:
                data_list.append(data)

    if data_list:
        df = pd.DataFrame(data_list)
        st.dataframe(df)

        # Tạo và cho tải file Excel
        excel_data = create_excel(data_list)
        st.download_button(
            label="Tải file Excel",
            data=excel_data,
            file_name="hoa_don.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Tạo và cho tải file Word
        word_doc = create_word(data_list)
        word_buffer = BytesIO()
        word_doc.save(word_buffer)
        word_bytes = word_buffer.getvalue()

        st.download_button(
            label="Tải file Word",
            data=word_bytes,
            file_name="hoa_don.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

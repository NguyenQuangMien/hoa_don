import streamlit as st
import pandas as pd
import pdfplumber
import zipfile
from io import BytesIO
from utils import extract_data_from_pdf, create_excel, create_word

st.title("Xử lý nhiều hóa đơn PDF hoặc ZIP")

uploaded_files = st.file_uploader(
    "Chọn nhiều file PDF hoặc file ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

data_list = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith('.zip'):
            with zipfile.ZipFile(uploaded_file) as z:
                for filename in z.namelist():
                    if filename.endswith('.pdf'):
                        with z.open(filename) as f:
                            with pdfplumber.open(f) as pdf:
                                text = ""
                                for page in pdf.pages:
                                    text += page.extract_text() + "\n"
                            data = extract_data_from_pdf(text)
                            if data:
                                data_list.append(data)
        else:
            with pdfplumber.open(uploaded_file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
            data = extract_data_from_pdf(text)
            if data:
                data_list.append(data)

    if data_list:
        df = pd.DataFrame(data_list)

        df.fillna('', inplace=True)

        st.write(f"Đã trích xuất {len(df)} hóa đơn.")
        st.dataframe(df)

        excel_file = create_excel(df)
        st.download_button(
            "Tải file Excel tổng hợp",
            data=excel_file,
            file_name="hoa_don_tong_hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        word_file = create_word(df)
        st.download_button(
            "Tải file Word tổng hợp",
            data=word_file,
            file_name="hoa_don_tong_hop.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

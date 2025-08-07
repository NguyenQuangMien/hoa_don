import streamlit as st
import pandas as pd
import pdfplumber
from utils import extract_data_from_pdf, create_excel, create_word
from io import BytesIO

st.title("Trích xuất dữ liệu hóa đơn PDF và xuất Excel, Word")

uploaded_file = st.file_uploader("Chọn file hóa đơn PDF", type=["pdf"])

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    
    # Trích xuất dữ liệu từ text PDF
    data = extract_data_from_pdf(text)

    if data:
        df = pd.DataFrame([data])
        st.write("Dữ liệu trích xuất:")
        st.dataframe(df)

        # Tạo file Excel
        excel_file = create_excel(df)
        st.download_button(
            label="Tải file Excel",
            data=excel_file,
            file_name="data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Tạo file Word
        word_file = create_word(df)
        st.download_button(
            label="Tải file Word",
            data=word_file,
            file_name="data.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Không thể trích xuất dữ liệu từ file PDF này.")

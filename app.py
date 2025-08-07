import streamlit as st
from utils import extract_data_from_pdf, create_excel, create_word

st.title("Trích xuất dữ liệu hóa đơn PDF và xuất Excel, Word")

uploaded_files = st.file_uploader("Chọn file hóa đơn PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data_list = []
    for file in uploaded_files:
        st.write(f"Đang xử lý file: {file.name}")
        data = extract_data_from_pdf(file)
        st.json(data)
        data_list.append(data)

    if data_list:
        excel_data = create_excel(data_list)
        st.download_button(
            label="Tải file Excel",
            data=excel_data,
            file_name="Hoa_don.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        word_data = create_word(data_list)
        st.download_button(
            label="Tải file Word",
            data=word_data,
            file_name="Hoa_don.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

import streamlit as st
from utils import extract_data_from_pdf, create_excel, create_word

st.title("Trích xuất dữ liệu hóa đơn PDF và xuất Excel, Word")

uploaded_files = st.file_uploader("Chọn file hóa đơn PDF", accept_multiple_files=True, type=["pdf"])

if uploaded_files:
    data_list = []
    for uploaded_file in uploaded_files:
        data = extract_data_from_pdf(uploaded_file)
        data_list.append(data)

    st.dataframe(data_list)

    excel_data = create_excel(data_list)
    word_data = create_word(data_list)

    st.download_button(
        label="Tải file Excel",
        data=excel_data,
        file_name="hoadon_trichxuat.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Tải file Word",
        data=word_data,
        file_name="hoadon_trichxuat.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

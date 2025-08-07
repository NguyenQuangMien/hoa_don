import streamlit as st
from utils import extract_data_from_pdf, create_excel, create_word

st.title("Trích xuất dữ liệu hóa đơn PDF và xuất Excel, Word")

uploaded_files = st.file_uploader("Chọn file PDF hoặc ZIP chứa nhiều file PDF", type=["pdf", "zip"], accept_multiple_files=True)

data_list = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        # Nếu là file PDF
        if uploaded_file.name.lower().endswith(".pdf"):
            st.write(f"Đang xử lý file: {uploaded_file.name}")
            data = extract_data_from_pdf(uploaded_file)
            st.json(data)  # In dữ liệu đọc được ra màn hình
            data_list.append(data)
        else:
            st.error(f"Không hỗ trợ file: {uploaded_file.name}")

    if data_list:
        df = create_excel(data_list)
        st.dataframe(df)

        excel_bytes = create_excel(data_list, return_bytes=True)
        st.download_button(
            label="Tải file Excel",
            data=excel_bytes,
            file_name="hoadon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        word_bytes = create_word(data_list, return_bytes=True)
        st.download_button(
            label="Tải file Word",
            data=word_bytes,
            file_name="hoadon.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

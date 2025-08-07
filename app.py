import streamlit as st
from utils import extract_data_from_pdf, create_excel, create_word

st.title("Trích xuất dữ liệu hóa đơn PDF và xuất Excel, Word")

uploaded_files = st.file_uploader(
    "Chọn nhiều file PDF hoặc file ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True,
)

data_list = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        # Trích xuất dữ liệu từ từng file
        data = extract_data_from_pdf(uploaded_file)

        # In dữ liệu thô ra màn hình để kiểm tra
        st.write(f"Dữ liệu thô trích xuất từ file: {uploaded_file.name}")
        st.json(data)

        data_list.append(data)

    if data_list:
        # Hiển thị bảng dữ liệu tổng hợp
        st.dataframe(data_list)

        # Tạo file Excel và cung cấp nút tải
        excel_data = create_excel(data_list)
        st.download_button(
            label="Tải file Excel",
            data=excel_data,
            file_name="hoa_don.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Tạo file Word và cung cấp nút tải
        word_data = create_word(data_list)
        st.download_button(
            label="Tải file Word",
            data=word_data,
            file_name="hoa_don.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

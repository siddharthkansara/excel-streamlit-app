# app.py
import streamlit as st
from excel_parser import process_excel

st.title("Excel Production Data Extractor")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.info("Processing file...")
    try:
        output_excel = process_excel(uploaded_file)
        st.success("Processing complete!")
        st.download_button(
            label="Download Processed Excel",
            data=output_excel,
            file_name="processed_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error while processing: {e}")

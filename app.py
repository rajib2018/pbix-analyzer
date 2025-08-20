import streamlit as st

st.title("PBIX File Analyzer")

uploaded_file = st.file_uploader("Upload your PBIX file", type=["pbix"])

if uploaded_file is not None:
    st.success(f"File '{uploaded_file.name}' uploaded successfully!")

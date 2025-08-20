import streamlit as st
import tempfile
import os
from pbixray import PBIXRay

st.title("PBIX File Analyzer and Document Generator")
st.write("Upload your Power BI (.pbix) file to analyze its structure and generate detailed documentation.")

uploaded_file = st.file_uploader("Choose a .pbix file", type="pbix")

if uploaded_file is not None:
    st.write("File uploaded successfully!")

model = PBIXRay(uploaded_file)
tables = model.tables
print(tables)

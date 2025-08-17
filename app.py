pip install pbixray

import streamlit as st
from pbixray import PBIXRay
import pandas as pd
import os
import tempfile

st.title("PBIX File Reader")

uploaded_file = st.file_uploader("Upload your .pbix file", type=["pbix"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pbix") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name

    try:
        model = PBIXRay(tmp_file_path)
        st.success("File successfully read and parsed!")

        st.header("Tables in the PBIX file:")
        for table_name in model.tables:
            st.subheader(f"Table: {table_name}")
            table_data = model.get_table(table_name)
            st.dataframe(table_data)

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")

    finally:
        os.remove(tmp_file_path)
else:
    st.info("Please upload a .pbix file to see its contents.")

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Power BI PBIX Documentation Generator")
st.title("Power BI PBIX File Documentation Generator")

st.markdown("""
Upload your `.pbix` Power BI file below. The app will extract and display metadata, tables, Power Query code, and other documentation using the pbixray package.
""")

try:
    from pbixray import PBIXRay
    pbixray_installed = True
except ImportError:
    pbixray_installed = False

if not pbixray_installed:
    st.error("The `pbixray` package is not installed. Please install it using the command:")
    st.code("pip install pbixray")
else:
    uploaded_file = st.file_uploader("Upload PBIX file", type="pbix")

    if uploaded_file:
        file_path = f"temp_{uploaded_file.name}"
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        model = PBIXRay(file_path)

        st.subheader("General Metadata")
        st.json(model.metadata)

        st.subheader("Tables")
        if hasattr(model, "tables"):
            for table in model.tables:
                st.write(f"**Table:** {table['name']}")
                df = pd.DataFrame(table["columns"])
                st.dataframe(df)

        st.subheader("Power Query Code")
        if hasattr(model, "power_query"):
            st.dataframe(model.power_query)
        
        st.subheader("M Parameters")
        if hasattr(model, "m_parameters"):
            st.dataframe(model.m_parameters)

        st.subheader("DAX Calculated Tables")
        if hasattr(model, "dax_tables"):
            st.dataframe(model.dax_tables)

        st.subheader("DAX Measures")
        if hasattr(model, "dax_measures"):
            st.dataframe(model.dax_measures)

        st.subheader("Calculated Columns")
        if hasattr(model, "dax_columns"):
            st.dataframe(model.dax_columns)

        st.subheader("Schema")
        if hasattr(model, "schema"):
            st.dataframe(model.schema)

        st.subheader("Relationships")
        if hasattr(model, "relationships"):
            st.dataframe(model.relationships)

        st.subheader("Statistics")
        if hasattr(model, "statistics"):
            st.dataframe(model.statistics)

        import os
        os.remove(file_path)
    else:
        st.info("Please upload a PBIX file to generate documentation.")

st.markdown("""
---
This app uses the pbixray package to parse Power BI files and generate documentation. Install pbixray with `pip install pbixray` if not already installed.
""")

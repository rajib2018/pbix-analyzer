import streamlit as st
import tempfile
import os
import pbixray

st.title("PBIX File Analyzer and Document Generator")
st.write("Upload your Power BI (.pbix) file to analyze its structure and generate detailed documentation.")

uploaded_file = st.file_uploader("Choose a .pbix file", type="pbix")

if uploaded_file is not None:
    st.write("File uploaded successfully!")

    # Create a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pbix") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    try:
        # Analyze the temporary file with pbixray
        analysis_result = pbixray.read(tmp_path)
        st.success("PBIX file analyzed successfully!")
        # You can display or process analysis_result here in future steps
        # st.write(analysis_result)

    except Exception as e:
        st.error(f"Error analyzing PBIX file: {e}")

    finally:
        # Remove the temporary file
        os.remove(tmp_path)

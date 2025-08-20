# app.py
import streamlit as st
from pbixray.core import PBIXRay
import os
import tempfile
from io import BytesIO
import pandas as pd
import traceback # Import traceback module
import xlsxwriter # Required by pandas for to_excel with engine='xlsxwriter'

st.set_page_config(
layout="wide", # Enables wide mode
page_title="PBIX Analyzer Advanced", # Sets the browser tab title
page_icon="📊" # Sets a favicon or emoji
)

st.set_page_config(
layout="wide", # Enables wide mode
page_title="PBIX Analyzer Advanced", # Sets the browser tab title
page_icon="📊" # Sets a favicon or emoji
)

# Function to generate Excel document
def generate_excel_doc(report_data):
    """Generates an Excel document with multiple sheets from the extracted report data."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for key, value in report_data.items():
            # Attempt to convert various data types to DataFrame for Excel
            if isinstance(value, pd.DataFrame):
                df = value
            elif isinstance(value, list):
                # Try to create a DataFrame from a list of dictionaries
                try:
                    df = pd.DataFrame(value)
                except Exception:
                    # If list items are not dictionaries or inconsistent,
                    # represent as a single column DataFrame
                    df = pd.DataFrame({key: value})
            elif isinstance(value, dict):
                 # Convert dictionary to DataFrame (e.g., for metadata if it's a dict)
                 df = pd.DataFrame.from_dict(value, orient='index', columns=[key])
            else:
                # Handle other types, perhaps as a single value DataFrame
                df = pd.DataFrame({key: [value]})

            # Write the DataFrame to a sheet named after the key
            if not df.empty:
                # Ensure sheet name is valid (max 31 chars, no invalid characters)
                sheet_name = key[:31]
                sheet_name = "".join([c for c in sheet_name if c.isalnum() or c in (' ', '_')]).rstrip()
                if not sheet_name:
                    sheet_name = "Sheet" + str(list(report_data.keys()).index(key) + 1) # Fallback name
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output


def main():
    st.title("Power BI Report Documentation Generator")

    uploaded_file = st.file_uploader("Upload your Power BI .pbix file", type="pbix")

    if uploaded_file is not None:
        try:
            # Create a temporary file to save the uploaded pbix
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pbix") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_pbix_path = tmp_file.name

            st.success(f"File uploaded successfully: {uploaded_file.name}")

            st.subheader("Extracting Report Information:")

            # Initialize PBIXRay with the temporary file path
            pbix_ray = PBIXRay(tmp_pbix_path)

            # Extract various pieces of information using pbixray
            metadata = pbix_ray.metadata
            st.write("Metadata:", metadata)

            schema = pbix_ray.schema
            st.write("Schema:", schema)

            relationships = pbix_ray.relationships
            st.write("Relationships:", relationships)

            power_query = pbix_ray.power_query
            st.write("Power Query:", power_query)

            m_parameters = pbix_ray.m_parameters
            st.write("M Parameters:", m_parameters)

            dax_tables = pbix_ray.dax_tables
            st.write("DAX Tables:", dax_tables)

            dax_measures = pbix_ray.dax_measures
            st.write("DAX Measures:", dax_measures)


            st.success("Information extracted successfully!")

            # Store extracted information in a dictionary for later use
            report_data = {
                "metadata": metadata,
                "schema": schema,
                "relationships": relationships,
                "power_query": power_query,
                "m_parameters": m_parameters,
                "dax_tables": dax_tables,
                "dax_measures": dax_measures,
            }

            # Print types and column names for debugging
            print("\n--- Debugging report_data types and columns ---")
            for key, value in report_data.items():
                print(f"Key: {key}, Type: {type(value)}")
                if isinstance(value, pd.DataFrame):
                    print(f"  DataFrame empty: {value.empty}")
                    if not value.empty:
                         print(f"  DataFrame columns: {value.columns.tolist()}")
                         # print(f"  DataFrame head:\n{value.head()}") # Uncomment for more detailed inspection if needed
                elif isinstance(value, list):
                     print(f"  List length: {len(value)}")
                     if value:
                          print(f"  First item type: {type(value[0])}")
                          # print(f"  First item:\n{value[0]}") # Uncomment for more detailed inspection if needed
                else:
                    print(f"  Value: {value}")
            print("---------------------------------------------")


            # Clean up the temporary file
            os.remove(tmp_pbix_path)

            st.subheader("Download Documentation:")

            # Add download button for Excel
            excel_doc_stream = generate_excel_doc(report_data)
            st.download_button(
                label="Download as Excel (.xlsx)",
                data=excel_doc_stream,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_documentation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.error(traceback.format_exc()) # Print the full traceback
            # Ensure temporary file is removed even if an error occurs
            if 'tmp_pbix_path' in locals() and os.path.exists(tmp_pbix_path):
                os.remove(tmp_pbix_path)


if __name__ == "__main__":
    main()

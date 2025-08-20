import streamlit as st
import tempfile
import os
import io # Import io for BytesIO

def sizeof_fmt(num, suffix='B'):
    """Formats a number into a human-readable byte string."""
    for unit in ['', 'Ki', 'Mi', 'Gi', 'Ti', 'Pi', 'Ei', 'Zi']:
        if abs(num) < 1024.0:
            return f'{num:3.1f}{unit}{suffix}'
        num /= 1024.0
    return f'{num:.1f}Yi{suffix}'

# Assume analyze_pbix_file and generate_documentation functions are defined elsewhere
# (from previous steps)

def st_app():
    """Main Streamlit application function for PBIX Analyzer."""
    st.set_page_config(page_title="PBIX Analyzer", layout="wide")
    st.title("PBIX Analyzer")

    uploaded_file = st.file_uploader("Upload a PBIX file", type=["pbix"])

    if uploaded_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pbix") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name

        try:
            st.info("Analyzing PBIX file...")
            # Assuming analyze_pbix_file is available from previous steps
            analysis_results = analyze_pbix_file(tmp_file_path)

            if analysis_results:
                st.success("Analysis complete!")

                # Display analysis results (as implemented in previous steps)
                # Display metadata
                st.header("Metadata")
                st.write(f"**File Size:** {sizeof_fmt(analysis_results.get('size_bytes', 0))}")
                if analysis_results.get("metadata"):
                    st.json(analysis_results["metadata"])

                # Display tables
                st.header("Tables")
                tables = analysis_results.get("tables", {})
                if tables:
                    st.write(f"Number of Tables: {len(tables)}")
                    for table_name, table_info in tables.items():
                        with st.expander(f"Table: {table_name} (Rows: {table_info.get('row_count', 'N/A')})"):
                            st.subheader("Columns")
                            columns = table_info.get("columns", {})
                            if columns:
                                column_data = [{"Name": col_name, **col_info} for col_name, col_info in columns.items()]
                                st.dataframe(column_data)
                            else:
                                st.write("No columns found for this table.")


                # Display measures
                st.header("Measures")
                measures = analysis_results.get("measures", {})
                if measures:
                    st.write(f"Number of Measures: {len(measures)}")
                    for measure_name, measure_info in measures.items():
                         with st.expander(f"Measure: {measure_name}"):
                            st.write(f"**Expression:**")
                            st.code(measure_info.get('expression', 'N/A'), language='dax')
                            st.write(f"**Display Folder:** {measure_info.get('display_folder', 'N/A')}")
                            st.write(f"**Format String:** {measure_info.get('format_string', 'N/A')}")
                            st.write(f"**Is Hidden:** {measure_info.get('is_hidden', 'N/A')}")

                # Display relationships
                st.header("Relationships")
                relationships = analysis_results.get("relationships", [])
                if relationships:
                    st.write(f"Number of Relationships: {len(relationships)}")
                    relationship_data = []
                    for rel in relationships:
                        relationship_data.append({
                            "Name": rel.get('name', 'N/A'),
                            "From Table": rel.get('from_table', 'N/A'),
                            "From Column": rel.get('from_column', 'N/A'),
                            "To Table": rel.get('to_table', 'N/A'),
                            "To Column": rel.get('to_column', 'N/A'),
                            "Model": rel.get('model', 'N/A'),
                            "Active": rel.get('active', 'N/A'),
                            "Cross Filter Direction": rel.get('cross_filter_direction', 'N/A'),
                            "Security Filter Table": rel.get('security_filter_table', 'N/A')
                        })
                    st.dataframe(relationship_data)
                else:
                    st.write("No relationships found.")

                # Display Power Query (M code)
                st.header("Power Query (M Code)")
                power_query = analysis_results.get("power_query", "No M code found.")
                if power_query and power_query.strip():
                     with st.expander("View M Code"):
                        st.code(power_query, language='m')
                else:
                    st.write("No Power Query (M Code) found.")


                # Display M Parameters
                st.header("M Parameters")
                m_parameters = analysis_results.get("m_parameters", [])
                if m_parameters:
                    st.write(f"Number of M Parameters: {len(m_parameters)}")
                    st.json(m_parameters)
                else:
                    st.write("No M Parameters found.")

                # Display DAX Tables (calculated tables)
                st.header("DAX Tables")
                dax_tables = analysis_results.get("dax_tables", {})
                if dax_tables:
                    st.write(f"Number of DAX Tables: {len(dax_tables)}")
                    for dax_table_name, dax_table_info in dax_tables.items():
                        with st.expander(f"DAX Table: {dax_table_name}"):
                             st.write(f"**Expression:**")
                             st.code(dax_table_info.get('expression', 'N/A'), language='dax')
                             st.write(f"**Display Folder:** {dax_table_info.get('display_folder', 'N/A')}")
                             st.write(f"**Is Hidden:** {dax_table_info.get('is_hidden', 'N/A')}")
                else:
                    st.write("No DAX Tables found.")


                # Documentation Generation Section
                st.header("Generate Documentation")
                output_format = st.radio("Select Output Format", ('Word', 'PDF'))

                if st.button("Generate Documentation"):
                    st.info(f"Generating {output_format} documentation...")
                    # Assuming generate_documentation is available from previous steps
                    doc_buffer = generate_documentation(analysis_results, output_format)

                    if doc_buffer:
                        file_name = f"pbix_documentation.{'docx' if output_format == 'Word' else 'pdf'}"
                        mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' if output_format == 'Word' else 'application/pdf'

                        st.download_button(
                            label=f"Download {output_format} Document",
                            data=doc_buffer,
                            file_name=file_name,
                            mime=mime_type
                        )
                        st.success(f"{output_format} documentation generated successfully!")
                    else:
                        st.error(f"Failed to generate {output_format} documentation.")

            else:
                st.error("Failed to analyze the PBIX file. Please ensure it is a valid PBIX file.")

        except Exception as e:
            st.error(f"An error occurred during analysis or documentation generation: {e}")
        finally:
            # Clean up the temporary file
            if os.path.exists(tmp_file_path):
                os.remove(tmp_file_path)

    else:
        st.info("Please upload a PBIX file to analyze.")

if __name__ == "__main__":
    pass

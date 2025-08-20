import streamlit as st
import pbixray
import pandas as pd
import os
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import io

def process_pbix_generic(pbix_input):
    """
    Processes a PBIX file using pbixray from a file path or file-like object
    and returns a dictionary of extracted information.

    Args:
        pbix_input: A file path (string) or a file-like object.

    Returns:
        A dictionary containing processed information from the PBIX file,
        or an error dictionary if processing fails.
    """
    processed_info = {}
    try:
        if isinstance(pbix_input, str) and os.path.exists(pbix_input):
            # If input is a file path, open it
            with open(pbix_input, 'rb') as f:
                pbix_file = pbixray.read_file(f)
        elif hasattr(pbix_input, 'read'):
            # If input is a file-like object (e.g., from st.file_uploader)
            pbix_file = pbixray.read_file(pbix_input)
        else:
            processed_info["error"] = "Invalid input: must be a file path or file-like object."
            return processed_info

        # Extract information using pbixray and store in the dictionary
        processed_info["Metadata"] = pbix_file.metadata
        processed_info["Model Size (bytes)"] = pbix_file.model_size_bytes
        processed_info["Number of Tables"] = len(pbix_file.model.tables)

        # Convert DataFrames to HTML for easier collation and display
        processed_info["Schema"] = pd.DataFrame(pbix_file.model.schema).to_html()
        processed_info["Statistics"] = pd.DataFrame(pbix_file.model.statistics).to_html()
        processed_info["Relationships"] = pd.DataFrame(pbix_file.model.relationships).to_html()
        processed_info["Power Query (M Code)"] = pbix_file.model.power_query_code
        processed_info["M Parameters"] = pd.DataFrame(pbix_file.model.m_parameters).to_html()
        processed_info["DAX Tables"] = pd.DataFrame(pbix_file.model.dax_tables).to_html()
        processed_info["DAX Measures"] = pd.DataFrame(pbix_file.model.dax_measures).to_html()
        processed_info["Calculated Columns"] = pd.DataFrame(pbix_file.model.calculated_columns).to_html()

        # Get data for the first table (for demonstration)
        first_table_name = "N/A"
        if pbix_file.model.tables:
            first_table_name = pbix_file.model.tables[0]['Name']
            try:
                 # Try to read the table data, handle potential errors
                table_data = pbix_file.model.read_table(first_table_name)
                processed_info[f"Data Preview: {first_table_name}"] = table_data.to_html()
            except Exception as e:
                processed_info[f"Data Preview: {first_table_name}"] = f"<p>Could not read data for table '{first_table_name}': {e}</p>"
        else:
             processed_info[f"Data Preview: {first_table_name}"] = "<p>No tables found in the model.</p>"


        return processed_info

    except Exception as e:
        processed_info["error"] = f"An error occurred during processing: {e}"
        return processed_info

def create_word_document(data):
    """Creates a Word document from the collated data."""
    document = Document()
    document.add_heading('PBIX File Analysis Report', 0)

    for title, content in data.items():
        document.add_heading(title, level=1)
        if isinstance(content, dict):
            for key, value in content.items():
                document.add_paragraph(f"{key}: {value}")
        elif isinstance(content, (int, float)):
             document.add_paragraph(str(content))
        else:
            # Simple attempt to handle HTML - this is basic and might not render complex HTML
            # For better HTML rendering in docx, consider libraries like `docx2python`
            # or converting HTML to text/markdown first.
            # Here, we'll just add the string content.
            document.add_paragraph(str(content))
        document.add_paragraph('') # Add a blank line between sections

    # Save document to a bytes buffer
    byte_io = io.BytesIO()
    document.save(byte_io)
    byte_io.seek(0)
    return byte_io

def create_pdf_document(data):
    """Creates a PDF document from the collated data using ReportLab."""
    byte_io = io.BytesIO()
    doc = SimpleDocTemplate(byte_io, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()

    story.append(Paragraph("PBIX File Analysis Report", styles['h1']))
    story.append(Spacer(1, 12))

    for title, content in data.items():
        story.append(Paragraph(title, styles['h2']))
        story.append(Spacer(1, 6))

        if isinstance(content, dict):
            for key, value in content.items():
                story.append(Paragraph(f"<b>{key}:</b> {value}", styles['Normal']))
        elif isinstance(content, (int, float)):
            story.append(Paragraph(str(content), styles['Normal']))
        else:
            # For HTML content, ReportLab doesn't directly render HTML easily.
            # We'll just add the string content as a paragraph.
            # For proper HTML rendering, consider libraries like `xhtml2pdf` or `WeasyPrint`.
            story.append(Paragraph(str(content), styles['Normal']))

        story.append(Spacer(1, 12)) # Add space after each section

    try:
        doc.build(story)
    except Exception as e:
        st.error(f"Error generating PDF: {e}")
        return None

    byte_io.seek(0)
    return byte_io

# --- Streamlit App ---
st.title("PBIX File Analyzer")

# Option to upload file or specify path
input_method = st.radio("Choose input method:", ("Upload File", "Specify File Path"))

pbix_input_source = None

if input_method == "Upload File":
    uploaded_file = st.file_uploader("Upload a .pbix file", type="pbix")
    if uploaded_file is not None:
        pbix_input_source = uploaded_file
elif input_method == "Specify File Path":
    file_path = st.text_input("Enter the path to the .pbix file:")
    if file_path:
        pbix_input_source = file_path

collated_data = {} # Initialize an empty dictionary to store collated data

if pbix_input_source:
    st.info("Processing file...")
    collated_data = process_pbix_generic(pbix_input_source) # Store returned dictionary

    if "error" in collated_data:
        st.error(collated_data["error"])
    else:
        st.success("File processed successfully!")

        # Display information by iterating through the collated_data dictionary
        for section_title, content in collated_data.items():
            st.subheader(section_title)
            if isinstance(content, str) and content.startswith('<'): # Check if it's likely HTML
                 st.write(content, unsafe_allow_html=True)
            elif isinstance(content, (int, float)): # Handle metrics differently if needed
                 if section_title == "Model Size (bytes)":
                      st.metric(section_title, f"{content:,} bytes")
                 elif section_title == "Number of Tables":
                      st.metric(section_title, content)
                 else:
                      st.write(content)
            else:
                st.write(content)

        st.subheader("Export Options")

        # Add download buttons
        if collated_data and "error" not in collated_data: # Ensure data is processed before offering export
            word_buffer = create_word_document(collated_data)
            if word_buffer:
                st.download_button(
                    label="Export to Word",
                    data=word_buffer,
                    file_name="pbix_analysis_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            pdf_buffer = create_pdf_document(collated_data)
            if pdf_buffer:
                 st.download_button(
                    label="Export to PDF",
                    data=pdf_buffer,
                    file_name="pbix_analysis_report.pdf",
                    mime="application/pdf"
                )
        else:
            st.warning("Process a file first to enable export options.")

# app.py
import streamlit as st
import os
from io import BytesIO
import pandas as pd
import traceback
import xlsxwriter
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

st.set_page_config(
layout="wide", # Enables wide mode
page_title="PBIX Analyzer Advanced", # Sets the browser tab title
page_icon="📊" # Sets a favicon or emoji
)

def generate_word_doc(report_data):
    """Generates a Word document from the extracted report data."""
    document = Document()

    document.add_heading('Excel File Documentation', 0)

    for sheet_name, df in report_data.items():
        document.add_heading(f"Sheet: {sheet_name}", level=1)
        if isinstance(df, pd.DataFrame) and not df.empty:
            document.add_paragraph("Table Data:")
            # Add DataFrame to Word document as a table
            table = document.add_table(rows=1, cols=len(df.columns))
            table.style = 'Table Grid'

            # Add header row
            hdr_cells = table.rows[0].cells
            for i, col_name in enumerate(df.columns):
                hdr_cells[i].text = str(col_name)

            # Add data rows
            for index, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, col_data in enumerate(row):
                    row_cells[i].text = str(col_data)
        else:
            document.add_paragraph("No data available for this sheet.")
        document.add_page_break() # Add page break after each sheet


    document_stream = BytesIO()
    document.save(document_stream)
    document_stream.seek(0)
    return document_stream

def generate_pdf_doc(report_data):
    """Generates a PDF document from the extracted report data using ReportLab Flowables."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=18)
    story = []
    styles = getSampleStyleSheet()

    # Add title
    story.append(Paragraph("Excel File Documentation", styles['Title']))
    story.append(Spacer(1, 0.2 * inch))

    for sheet_name, df in report_data.items():
        story.append(Paragraph(f"Sheet: {sheet_name}", styles['Heading1']))
        story.append(Spacer(1, 0.1 * inch))

        if isinstance(df, pd.DataFrame) and not df.empty:
            story.append(Paragraph("Table Data:", styles['Heading2']))
            # Convert DataFrame to a list of lists for the table
            data = [df.columns.tolist()] + df.values.tolist()

            # Create a ReportLab table (requires reportlab.platypus.tables)
            from reportlab.platypus import Table, TableStyle
            from reportlab.lib import colors

            table = Table(data)

            # Add table style
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ])
            table.setStyle(style)

            story.append(table)
        else:
            story.append(Paragraph("No data available for this sheet.", styles['Normal']))

        story.append(Spacer(1, 0.5 * inch)) # Space after each sheet data
        # story.append(PageBreak()) # Add page break after each sheet (optional)


    doc.build(story)
    buffer.seek(0)
    return buffer


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
    st.title("Excel File Documentation Generator")

    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

    if uploaded_file is not None:
        try:
            st.success(f"File uploaded successfully: {uploaded_file.name}")

            st.subheader("Reading Excel Data:")

            # Read the Excel file into a dictionary of DataFrames, one per sheet
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)

            # Store extracted information in a dictionary (using sheet names as keys)
            report_data = excel_data

            st.success("Data read from Excel successfully!")


            st.subheader("Download Documentation:")

            # Add download buttons for Word, PDF, and Excel
            word_doc_stream = generate_word_doc(report_data)
            st.download_button(
                label="Download as Word (.docx)",
                data=word_doc_stream,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_documentation.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            pdf_doc_stream = generate_pdf_doc(report_data)
            st.download_button(
                label="Download as PDF (.pdf)",
                data=pdf_doc_stream,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_documentation.pdf",
                mime="application/pdf"
            )

            # Add download button for the processed Excel file (optional, but keeps consistency)
            excel_output_stream = generate_excel_doc(report_data) # Reuse the excel generator
            st.download_button(
                 label="Download as Excel (.xlsx)",
                 data=excel_output_stream,
                 file_name=f"{os.path.splitext(uploaded_file.name)[0]}_processed.xlsx",
                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.error(traceback.format_exc())


if __name__ == "__main__":
    main()

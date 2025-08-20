# app.py
import streamlit as st
from pbixray.core import PBIXRay
import os
import tempfile
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
import pandas as pd
import traceback # Import traceback module
import xlsxwriter # Ensure xlsxwriter is imported if needed explicitly, though pandas uses it internally

st.set_page_config(
layout="wide", # Enables wide mode
page_title="PBIX Analyzer Advanced", # Sets the browser tab title
page_icon="📊" # Sets a favicon or emoji
)


# Helper functions for PDF generation (moved outside to be accessible by generate_pdf_doc)
def draw_text(c, text, x, y, size=12, bold=False):
    c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
    c.drawString(x, y, text)
    return y - size - 5 # Return new y position

def draw_paragraph(c, text, x, y, size=10):
    c.setFont("Helvetica", size)
    # Simple text wrapping (can be improved for ReportLab canvas)
    # For direct canvas drawing, we need manual line breaking or use flowables
    # in a different way. Let's stick to basic drawing for now but improve key access.

    # Basic truncation for simplicity, as direct drawing doesn't wrap easily.
    # A more robust solution would involve splitting the text into lines.
    lines = text.split('\n')
    current_y = y
    for line in lines:
        c.drawString(x, current_y, line[:100] + "..." if len(line) > 100 else line)
        current_y -= size + 2 # Adjust spacing between lines
        if current_y < 50: # Simple page break
             c.showPage()
             current_y = letter[1] - 50 # Reset y position for new page

    return current_y # Return the final y position


def generate_word_doc(report_data):
    """Generates a Word document from the extracted report data."""
    document = Document()

    document.add_heading('Power BI Report Documentation', 0)

    # Add Metadata
    document.add_heading('Metadata', level=1)
    metadata = report_data.get("metadata")
    # Step-by-step check for metadata type
    if metadata is not None:
        if isinstance(metadata, pd.DataFrame) and not metadata.empty:
             document.add_paragraph("Metadata (DataFrame format):")
             for index, row in metadata.iterrows():
                 for col, value in row.items():
                      document.add_paragraph(f"{col}: {value if pd.notna(value) else 'N/A'}") # Handle potential NaN in DataFrame
        elif isinstance(metadata, dict):
            document.add_paragraph("Metadata (Dictionary format):")
            for key, value in metadata.items():
                document.add_paragraph(f"{key}: {value if value else 'N/A'}") # Handle potential empty strings or None in dict
        else:
             document.add_paragraph(f"Metadata available in unexpected format: {type(metadata)}")
    else:
        document.add_paragraph("No metadata available.")


    # Add Schema
    document.add_heading('Schema', level=1)
    schema = report_data.get("schema")
    if schema is not None:
        if isinstance(schema, pd.DataFrame) and not schema.empty: # Handle case where schema might be a DataFrame
             document.add_paragraph("Schema Information (DataFrame format):")
             for index, row in schema.iterrows():
                 # print(f"Word Doc - Processing Schema Row (DataFrame): {row.to_dict()}") # Debug print
                 document.add_paragraph(f"  Table: {row.get('name', row.get('Name', 'N/A'))}") # Added common variations
                 # Assuming DataFrame schema has 'columns' as a string representation or similar
                 columns_info = row.get('columns', row.get('Columns', 'N/A')) # Added common variations
                 document.add_paragraph(f"  Columns: {columns_info if columns_info else 'N/A'}") # Handle potential empty columns info
        elif isinstance(schema, list): # Schema is expected to be a list of dictionaries
            document.add_paragraph("Schema Information (List format):")
            for table in schema:
                # print(f"Word Doc - Processing Schema Table (List): {table}") # Debug print
                document.add_heading(f"Table: {table.get('name', table.get('Name', 'N/A'))}", level=2) # Added common variations
                if table.get("columns"):
                    document.add_paragraph("Columns:")
                    for column in table["columns"]:
                        # print(f"Word Doc - Processing Schema Column (List): {column}") # Debug print
                        document.add_paragraph(f"- {column.get('name', column.get('Name', 'N/A'))} ({column.get('dataType', column.get('DataType', 'N/A'))})") # Added common variations
                else:
                    document.add_paragraph("No columns found for this table.")
        else:
            document.add_paragraph(f"No schema information available or in unexpected format: {type(schema)}")
    else:
        document.add_paragraph("No schema information available.")


    # Add Relationships
    document.add_heading('Relationships', level=1)
    relationships = report_data.get("relationships")
    if relationships is not None and hasattr(relationships, 'empty') and not relationships.empty:
        document.add_paragraph("Relationships:")
        for rel in relationships.to_dict('records'):
            # print(f"Word Doc - Processing Relationship: {rel}") # Debug print
            # Refine key access based on potential pbixray output structure
            from_table = rel.get('from_table', rel.get('From Table', rel.get('fromTable', 'N/A')))
            from_column = rel.get('from_column', rel.get('From Column', rel.get('fromColumn', 'N/A')))
            to_table = rel.get('to_table', rel.get('To Table', rel.get('toTable', 'N/A')))
            to_column = rel.get('to_column', rel.get('To Column', rel.get('toColumn', 'N/A')))
            document.add_paragraph(f"From Table: {from_table}, From Column: {from_column}, To Table: {to_table}, To Column: {to_column}")
    else:
        document.add_paragraph("No relationships available.")

    # Add Power Query Code
    document.add_heading('Power Query Code', level=1)
    power_query = report_data.get("power_query")
    if power_query is not None and hasattr(power_query, 'empty') and not power_query.empty:
        document.add_paragraph("Power Query Code:")
        for pq in power_query.to_dict('records'):
             # print(f"Word Doc - Processing Power Query: {pq}") # Debug print
             # Refine key access based on potential pbixray output structure
             name = pq.get('name', pq.get('Name', 'N/A'))
             expression = pq.get('expression', pq.get('Expression', 'N/A'))
             document.add_paragraph(f"Name: {name}")
             document.add_paragraph(f"Expression: {expression}") # Always print expression with N/A if missing
    else:
        document.add_paragraph("No Power Query code available.")

    # Add M Parameters
    document.add_heading('M Parameters', level=1)
    m_parameters = report_data.get("m_parameters")
    if m_parameters is not None and hasattr(m_parameters, 'empty') and not m_parameters.empty:
        document.add_paragraph("M Parameters:")
        for param in m_parameters.to_dict('records'):
            # print(f"Word Doc - Processing M Parameter: {param}") # Debug print
            # Refine key access based on potential pbixray output structure
            name = param.get('name', param.get('Name', 'N/A'))
            value = param.get('value', param.get('Value', 'N/A'))
            document.add_paragraph(f"Name: {name}, Value: {value}")
    else:
        document.add_paragraph("No M parameters available.")

    # Add DAX Tables
    document.add_heading('DAX Tables', level=1)
    dax_tables = report_data.get("dax_tables")
    if dax_tables is not None and hasattr(dax_tables, 'empty') and not dax_tables.empty:
         document.add_paragraph("DAX Tables:")
         for table in dax_tables.to_dict('records'):
            # print(f"Word Doc - Processing DAX Table: {table}") # Debug print
            # Refine key access based on potential pbixray output structure
            name = table.get('name', table.get('Name', 'N/A'))
            expression = table.get('expression', table.get('Expression', 'N/A'))
            document.add_paragraph(f"Name: {name}")
            document.add_paragraph(f"Expression: {expression}") # Always print expression with N/A if missing
    else:
        document.add_paragraph("No DAX tables available.")

    # Add DAX Measures
    document.add_heading('DAX Measures', level=1)
    dax_measures = report_data.get("dax_measures")
    if dax_measures is not None and hasattr(dax_measures, 'empty') and not dax_measures.empty:
        document.add_paragraph("DAX Measures:")
        for measure in dax_measures.to_dict('records'):
            # print(f"Word Doc - Processing DAX Measure: {measure}") # Debug print
            # Refine key access based on potential pbixray output structure
            name = measure.get('name', measure.get('Name', 'N/A'))
            expression = measure.get('expression', measure.get('Expression', 'N/A'))
            document.add_paragraph(f"Name: {name}")
            document.add_paragraph(f"Expression: {expression}") # Always print expression with N/A if missing
    else:
        document.add_paragraph("No DAX measures available.")

    document_stream = BytesIO()
    document.save(document_stream)
    document_stream.seek(0)
    return document_stream

def generate_pdf_doc(report_data):
    """Generates a PDF document from the extracted report data."""
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    c.drawString(100, height - 50, "Power BI Report Documentation")
    y_position = height - 100

    # Add Metadata
    y_position = draw_text(c, "Metadata", 100, y_position, size=14, bold=True)
    metadata = report_data.get("metadata")
    # Step-by-step check for metadata type
    if metadata is not None:
        if isinstance(metadata, pd.DataFrame) and not metadata.empty:
             y_position = draw_paragraph(c, "Metadata (DataFrame format):", 100, y_position)
             for index, row in metadata.iterrows():
                 # print(f"PDF Doc - Processing Metadata Row (DataFrame): {row.to_dict()}") # Debug print
                 for col, value in row.items():
                      y_position = draw_paragraph(c, f"{col}: {value if pd.notna(value) else 'N/A'}", 100, y_position) # Handle potential NaN in DataFrame
                 y_position -= 10 # Add some space between records
        elif isinstance(metadata, dict):
            y_position = draw_paragraph(c, "Metadata (Dictionary format):", 100, y_position)
            for key, value in metadata.items():
                 # print(f"PDF Doc - Processing Metadata Item (Dict): {key}: {value}") # Debug print
                 y_position = draw_paragraph(c, f"{key}: {value if value else 'N/A'}", 100, y_position) # Handle potential empty strings or None in dict
        else:
            y_position = draw_paragraph(c, f"Metadata available in unexpected format: {type(metadata)}", 100, y_position)
    else:
        y_position = draw_paragraph(c, "No metadata available.", 100, y_position)

    # Add Schema
    y_position = draw_text(c, "Schema", 100, y_position, size=14, bold=True)
    schema = report_data.get("schema")
    if schema is not None:
        if isinstance(schema, pd.DataFrame) and not schema.empty: # Handle case where schema might be a DataFrame
             y_position = draw_paragraph(c, "Schema Information (DataFrame format):", 100, y_position)
             for index, row in schema.iterrows():
                 # print(f"PDF Doc - Processing Schema Row (DataFrame): {row.to_dict()}") # Debug print
                 y_position = draw_paragraph(c, f"  Table: {row.get('name', row.get('Name', 'N/A'))}", 100, y_position) # Added common variations
                 # Assuming DataFrame schema has 'columns' as a string representation or similar
                 columns_info = row.get('columns', row.get('Columns', 'N/A')) # Added common variations
                 y_position = draw_paragraph(c, f"  Columns: {columns_info if columns_info else 'N/A'}", 100, y_position) # Handle potential empty columns info
                 y_position -= 10 # Add some space between records
        elif isinstance(schema, list): # Schema is expected to be a list of dictionaries
            y_position = draw_paragraph(c, "Schema Information (List format):", 100, y_position)
            for table in schema:
                # print(f"PDF Doc - Processing Schema Table (List): {table}") # Debug print
                y_position = draw_text(c, f"Table: {table.get('name', table.get('Name', 'N/A'))}", 100, y_position, size=12, bold=True) # Added common variations
                if table.get("columns"):
                    y_position = draw_text(c, "Columns:", 100, y_position, size=10)
                    for column in table["columns"]:
                        # print(f"PDF Doc - Processing Schema Column (List): {column}") # Debug print
                        y_position = draw_paragraph(c, f"- {column.get('name', column.get('Name', 'N/A'))} ({column.get('dataType', column.get('DataType', 'N/A'))})", 100, y_position) # Added common variations
                    y_position -= 10 # Add some space after columns
                else:
                     y_position = draw_paragraph(c, "No columns found for this table.", 100, y_position)
        else:
            y_position = draw_paragraph(c, f"No schema information available or in unexpected format: {type(schema)}", 100, y_position)
    else:
        y_position = draw_paragraph(c, "No schema information available.", 100, y_position)


    # Add Relationships
    y_position = draw_text(c, "Relationships", 100, y_position, size=14, bold=True)
    relationships = report_data.get("relationships")
    if relationships is not None and hasattr(relationships, 'empty') and not relationships.empty:
        y_position = draw_paragraph(c, "Relationships:", 100, y_position)
        for rel in relationships.to_dict('records'):
            # print(f"PDF Doc - Processing Relationship: {rel}") # Debug print
            # Refine key access based on potential pbixray output structure
            from_table = rel.get('from_table', rel.get('From Table', rel.get('fromTable', 'N/A')))
            from_column = rel.get('from_column', rel.get('From Column', rel.get('fromColumn', 'N/A')))
            to_table = rel.get('to_table', rel.get('To Table', rel.get('toTable', 'N/A')))
            to_column = rel.get('to_column', rel.get('To Column', rel.get('toColumn', 'N/A')))
            y_position = draw_paragraph(c, f"From Table: {from_table}, From Column: {from_column}, To Table: {to_table}, To Column: {to_column}", 100, y_position)
            y_position -= 10 # Add some space between records
    else:
        y_position = draw_paragraph(c, "No relationships available.", 100, y_position)

    # Add Power Query Code
    y_position = draw_text(c, "Power Query Code", 100, y_position, size=14, bold=True)
    power_query = report_data.get("power_query")
    if power_query is not None and hasattr(power_query, 'empty') and not power_query.empty:
        y_position = draw_paragraph(c, "Power Query Code:", 100, y_position)
        for pq in power_query.to_dict('records'):
             # print(f"PDF Doc - Processing Power Query: {pq}") # Debug print
             # Refine key access based on potential pbixray output structure
             name = pq.get('name', pq.get('Name', 'N/A'))
             expression = pq.get('expression', pq.get('Expression', 'N/A'))
             y_position = draw_paragraph(c, f"Name: {name}", 100, y_position)
             y_position = draw_paragraph(c, f"Expression: {expression}", 100, y_position) # Always print expression with N/A if missing
             y_position -= 10 # Add some space between records
    else:
        y_position = draw_paragraph(c, "No Power Query code available.", 100, y_position)

    # Add M Parameters
    y_position = draw_text(c, "M Parameters", 100, y_position, size=14, bold=True)
    m_parameters = report_data.get("m_parameters")
    if m_parameters is not None and hasattr(m_parameters, 'empty') and not m_parameters.empty:
        y_position = draw_paragraph(c, "M Parameters:", 100, y_position)
        for param in m_parameters.to_dict('records'):
            # print(f"PDF Doc - Processing M Parameter: {param}") # Debug print
            # Refine key access based on potential pbixray output structure
            name = param.get('name', param.get('Name', 'N/A'))
            value = param.get('value', param.get('Value', 'N/A'))
            y_position = draw_paragraph(c, f"Name: {name}, Value: {value}", 100, y_position)
            y_position -= 10 # Add some space between records
    else:
        y_position = draw_paragraph(c, "No M parameters available.", 100, y_position)

    # Add DAX Tables
    y_position = draw_text(c, "DAX Tables", 100, y_position, size=14, bold=True)
    dax_tables = report_data.get("dax_tables")
    if dax_tables is not None and hasattr(dax_tables, 'empty') and not dax_tables.empty:
         y_position = draw_paragraph(c, "DAX Tables:", 100, y_position)
         for table in dax_tables.to_dict('records'):
            # print(f"PDF Doc - Processing DAX Table: {table}") # Debug print
            # Refine key access based on potential pbixray output structure
            name = table.get('name', table.get('Name', 'N/A'))
            expression = table.get('expression', table.get('Expression', 'N/A'))
            y_position = draw_paragraph(c, f"Name: {name}", 100, y_position)
            y_position = draw_paragraph(c, f"Expression: {expression}", 100, y_position) # Always print expression with N/A if missing
            y_position -= 10 # Add some space between records
    else:
        y_position = draw_paragraph(c, "No DAX tables available.", 100, y_position)

    # Add DAX Measures
    y_position = draw_text(c, "DAX Measures", 100, y_position, size=14, bold=True)
    dax_measures = report_data.get("dax_measures")
    if dax_measures is not None and hasattr(dax_measures, 'empty') and not dax_measures.empty:
        y_position = draw_paragraph(c, "DAX Measures:", 100, y_position)
        for measure in dax_measures.to_dict('records'):
            # print(f"PDF Doc - Processing DAX Measure: {measure}") # Debug print
            # Refine key access based on potential pbixray output structure
            name = measure.get('name', measure.get('Name', 'N/A'))
            expression = measure.get('expression', measure.get('Expression', 'N/A'))
            y_position = draw_paragraph(c, f"Name: {name}", 100, y_position)
            y_position = draw_paragraph(c, f"Expression: {expression}", 100, y_position) # Always print expression with N/A if missing
            y_position -= 10 # Add some space between records
    else:
        y_position = draw_paragraph(c, "No DAX measures available.", 100, y_position)


    c.save()
    buffer.seek(0)
    return buffer

# Function to generate Excel document (copied from previous cell)
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

            # Initialize PBIXRay with the temporary file path
            pbix_ray = PBIXRay(tmp_pbix_path)

            st.subheader("Extracting Report Information:")

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

            # Add download buttons for Word and PDF
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

import streamlit as st
import subprocess
import json
import pandas as pd
import os
from docx import Document
from io import BytesIO
from fpdf import FPDF
from pbixray.core import PBIXRay

# Function to run pbixray
def run_pbixray(pbix_file_path):
    """Runs pbixray on the given file and returns the parsed JSON output."""
    try:
        # Construct the command to run pbixray
        command = ["PBIXRay", "-f", pbix_file_path]

        # Run the command and capture the output
        result = subprocess.run(command, capture_output=True, text=True, check=True)

        # Parse the JSON output
        parsed_output = json.loads(result.stdout)
        return parsed_output

    except FileNotFoundError:
        st.error("Error: pbixray command not found. Make sure pbixray is installed and in your PATH.")
        return None
    except subprocess.CalledProcessError as e:
        st.error(f"Error running pbixray: {e}")
        st.error(f"Stderr: {e.stderr}")
        return None
    except json.JSONDecodeError:
        st.error("Error decoding JSON output from pbixray.")
        return None

# Functions to display different sections
def display_model_info(model_data):
    """Displays general model information."""
    st.header("Model Information")
    if model_data:
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Name:** {model_data.get('name', 'N/A')}")
            st.write(f"**Culture:** {model_data.get('culture', 'N/A')}")
            st.write(f"**Default Mode:** {model_data.get('defaultMode', 'N/A')}")
            st.write(f"**Default Power BI Abstract Visual:** {model_data.get('defaultPowerBIAbstractVisual', 'N/A')}")
        with col2:
            st.write(f"**Description:** {model_data.get('description', 'N/A')}")
            st.write(f"**Modified Time:** {model_data.get('modifiedTime', 'N/A')}")
            st.write(f"**Created Time:** {model_data.get('createdTime', 'N/A')}")
            st.write(f"**Last Data Refresh:** {model_data.get('lastDataRefresh', 'N/A')}")
    else:
        st.info("No model information available.")


def display_tables(tables_data):
    """Displays the tables information in a user-friendly format with a dropdown."""
    st.header("Tables")

    if tables_data:
        table_names = [table['name'] for table in tables_data]
        selected_table_name = st.selectbox("Select a Table", table_names)

        selected_table = next((table for table in tables_data if table['name'] == selected_table_name), None)

        if selected_table:
            st.subheader(f"Details for Table: {selected_table['name']}")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Description:** {selected_table.get('description', 'N/A')}")
                st.write(f"**Hidden:** {selected_table.get('isHidden', 'N/A')}")
            with col2:
                st.write(f"**Data Category:** {selected_table.get('dataCategory', 'N/A')}")
                st.write(f"**Lineage Tag:** {selected_table.get('lineageTag', 'N/A')}")


            if 'columns' in selected_table and selected_table['columns']:
                st.write("---") # Separator
                st.subheader("Columns:")
                columns_df = pd.DataFrame(selected_table['columns'])
                # Reorder columns for better readability if needed
                column_order = ['name', 'dataType', 'description', 'isHidden', 'dataCategory', 'summarizeBy', 'displayFolder', 'formatString', 'lineageTag']
                existing_columns = [col for col in column_order if col in columns_df.columns]
                st.dataframe(columns_df[existing_columns])
            else:
                st.info("No columns found for this table.")
    else:
        st.info("No table data available.")


def display_measures(measures_data):
    """Displays the measures information."""
    st.header("Measures")
    if measures_data:
        for measure in measures_data:
            with st.expander(f"Measure: {measure['name']}"):
                st.write(f"**Expression:** ```\n{measure['expression']}\n```")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Description:** {measure.get('description', 'N/A')}")
                    st.write(f"**Display Folder:** {measure.get('displayFolder', 'N/A')}")
                    st.write(f"**Hidden:** {measure.get('isHidden', 'N/A')}")
                with col2:
                    st.write(f"**Format String:** {measure.get('formatString', 'N/A')}")
                    st.write(f"**Data Category:** {measure.get('dataCategory', 'N/A')}")
                    st.write(f"**Display Format:** {measure.get('displayFormat', 'N/A')}")
                    st.write(f"**Lineage Tag:** {measure.get('lineageTag', 'N/A')}")
    else:
        st.info("No measures available.")


def display_relationships(relationships_data):
    """Displays the relationships information."""
    st.header("Relationships")
    if relationships_data:
        for relationship in relationships_data:
            with st.expander(f"Relationship: {relationship.get('name', 'N/A')}"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**From Table:** {relationship.get('fromTable', 'N/A')}")
                    st.write(f"**From Column:** {relationship.get('fromColumn', 'N/A')}")
                    st.write(f"**To Table:** {relationship.get('toTable', 'N/A')}")
                    st.write(f"**To Column:** {relationship.get('toColumn', 'N/A')}")
                with col2:
                    st.write(f"**Join On Date Behavior:** {relationship.get('joinOnDateBehavior', 'N/A')}")
                    st.write(f"**State:** {relationship.get('state', 'N/A')}")
                    st.write(f"**Type:** {relationship.get('type', 'N/A')}")
                    st.write(f"**Cross Filtering Behavior:** {relationship.get('crossFilteringBehavior', 'N/A')}")
                    st.write(f"**Active:** {relationship.get('isActive', 'N/A')}")
                    st.write(f"**Security Filtering Behavior:** {relationship.get('securityFilteringBehavior', 'N/A')}")
                    st.write(f"**Lineage Tag:** {relationship.get('lineageTag', 'N/A')}")
    else:
        st.info("No relationships available.")


def display_cultures(cultures_data):
    """Displays cultures information."""
    st.header("Cultures")
    if cultures_data:
        for culture in cultures_data:
            with st.expander(f"Culture: {culture.get('name', 'N/A')}"):
                st.write(f"**Language:** {culture.get('language', 'N/A')}")
                if 'linguisticMetadata' in culture and culture['linguisticMetadata']:
                    st.write("**Linguistic Metadata:**")
                    st.json(culture['linguisticMetadata'])
                else:
                    st.info("No linguistic metadata available for this culture.")
    else:
        st.info("No cultures data available.")

def display_datasources(datasources_data):
    """Displays datasources information."""
    st.header("Data Sources")
    if datasources_data:
        for datasource in datasources_data:
            with st.expander(f"Data Source: {datasource.get('name', 'N/A')}"):
                st.write(f"**Type:** {datasource.get('type', 'N/A')}")
                st.write(f"**Connection String:** {datasource.get('connectionString', 'N/A')}")
                st.write(f"**Datasource Type:** {datasource.get('datasourceType', 'N/A')}")
                st.write(f"**Gateway:** {datasource.get('gateway', 'N/A')}")
                # Add more datasource details if available
                st.write(f"**Server:** {datasource.get('server', 'N/A')}")
                st.write(f"**Database:** {datasource.get('database', 'N/A')}")
    else:
        st.info("No data sources available.")

def display_roles(roles_data):
    """Displays roles information."""
    st.header("Roles")
    if roles_data:
        for role in roles_data:
            with st.expander(f"Role: {role.get('name', 'N/A')}"):
                st.write(f"**Model Permission:** {role.get('modelPermission', 'N/A')}")
                if 'tablePermissions' in role and role['tablePermissions']:
                    st.write("**Table Permissions:**")
                    for tp in role['tablePermissions']:
                        st.write(f"- **Table:** {tp.get('name', 'N/A')}")
                        st.write(f"  - **Filter Expression:** ```\n{tp.get('filterExpression', 'N/A')}\n```")
                        st.write(f"  - **Rows Access:** {tp.get('rowsAccess', 'N/A')}")
                else:
                    st.info("No table permissions defined for this role.")
    else:
        st.info("No roles available.")

def display_expressions(expressions_data):
    """Displays expressions information (e.g., M or DAX expressions)."""
    st.header("Expressions")
    if expressions_data:
        for expression in expressions_data:
            with st.expander(f"Expression: {expression.get('name', 'N/A')}"):
                st.write(f"**Kind:** {expression.get('kind', 'N/A')}")
                st.write(f"**Expression:** ```\n{expression.get('expression', 'N/A')}\n```")
                st.write(f"**Hidden:** {expression.get('isHidden', 'N/A')}")
                st.write(f"**Lineage Tag:** {expression.get('lineageTag', 'N/A')}")
    else:
        st.info("No expressions available.")

def display_annotations(annotations_data):
    """Displays annotations information."""
    st.header("Annotations")
    if annotations_data:
        for annotation in annotations_data:
            with st.expander(f"Annotation: {annotation.get('name', 'N/A')}"):
                st.write(f"**Value:** {annotation.get('value', 'N/A')}")
    else:
        st.info("No annotations available.")


def generate_word_doc(data):
    """Generates a Word document from the parsed PBIX data."""
    document = Document()
    document.add_heading('PBIX Documentation', 0)

    if 'model' in data and data['model']:
        document.add_heading('Model Information', level=1)
        model_data = data['model']
        document.add_paragraph(f"Name: {model_data.get('name', 'N/A')}")
        document.add_paragraph(f"Culture: {model_data.get('culture', 'N/A')}")
        document.add_paragraph(f"Default Mode: {model_data.get('defaultMode', 'N/A')}")
        document.add_paragraph(f"Description: {model_data.get('description', 'N/A')}")
        document.add_paragraph(f"Modified Time: {model_data.get('modifiedTime', 'N/A')}")


    if 'tables' in data and data['tables']:
        document.add_heading('Tables', level=1)
        for table in data['tables']:
            document.add_heading(f"Table: {table['name']}", level=2)
            document.add_paragraph(f"Description: {table.get('description', 'N/A')}")
            document.add_paragraph(f"Hidden: {table.get('isHidden', 'N/A')}")
            document.add_paragraph(f"Data Category: {table.get('dataCategory', 'N/A')}")
            document.add_paragraph(f"Lineage Tag: {table.get('lineageTag', 'N/A')}")

            if 'columns' in table and table['columns']:
                document.add_paragraph("Columns:")
                df_columns = pd.DataFrame(table['columns'])
                column_order = ['name', 'dataType', 'description', 'isHidden', 'dataCategory', 'summarizeBy', 'displayFolder', 'formatString', 'lineageTag']
                existing_columns = [col for col in column_order if col in df_columns.columns]
                document.add_paragraph(df_columns[existing_columns].to_markdown(index=False))

    if 'measures' in data and data['measures']:
        document.add_heading('Measures', level=1)
        for measure in data['measures']:
            document.add_heading(f"Measure: {measure['name']}", level=2)
            document.add_paragraph(f"Expression: {measure['expression']}")
            document.add_paragraph(f"Description: {measure.get('description', 'N/A')}")
            document.add_paragraph(f"Display Folder: {measure.get('displayFolder', 'N/A')}")
            document.add_paragraph(f"Format String: {measure.get('formatString', 'N/A')}")
            document.add_paragraph(f"Hidden: {measure.get('isHidden', 'N/A')}")
            document.add_paragraph(f"Data Category: {measure.get('dataCategory', 'N/A')}")
            document.add_paragraph(f"Display Format: {measure.get('displayFormat', 'N/A')}")
            document.add_paragraph(f"Lineage Tag: {measure.get('lineageTag', 'N/A')}")


    if 'relationships' in data and data['relationships']:
        document.add_heading('Relationships', level=1)
        for relationship in data['relationships']:
            document.add_heading(f"Relationship: {relationship.get('name', 'N/A')}", level=2)
            document.add_paragraph(f"From Table: {relationship.get('fromTable', 'N/A')}")
            document.add_paragraph(f"From Column: {relationship.get('fromColumn', 'N/A')}")
            document.add_paragraph(f"To Table: {relationship.get('toTable', 'N/A')}")
            document.add_paragraph(f"To Column: {relationship.get('toColumn', 'N/A')}")
            document.add_paragraph(f"Join On Date Behavior: {relationship.get('joinOnDateBehavior', 'N/A')}")
            document.add_paragraph(f"State: {relationship.get('state', 'N/A')}")
            document.add_paragraph(f"Type: {relationship.get('type', 'N/A')}")
            document.add_paragraph(f"Cross Filtering Behavior: {relationship.get('crossFilteringBehavior', 'N/A')}")
            document.add_paragraph(f"Active: {relationship.get('isActive', 'N/A')}")
            document.add_paragraph(f"Security Filtering Behavior: {relationship.get('securityFilteringBehavior', 'N/A')}")
            document.add_paragraph(f"Lineage Tag: {relationship.get('lineageTag', 'N/A')}")


    if 'cultures' in data and data['cultures']:
        document.add_heading('Cultures', level=1)
        for culture in data['cultures']:
            document.add_heading(f"Culture: {culture.get('name', 'N/A')}", level=2)
            document.add_paragraph(f"Language: {culture.get('language', 'N/A')}")
            if 'linguisticMetadata' in culture and culture['linguisticMetadata']:
                document.add_paragraph("Linguistic Metadata:")
                document.add_paragraph(json.dumps(culture['linguisticMetadata'], indent=2))

    if 'datasources' in data and data['datasources']:
        document.add_heading('Data Sources', level=1)
        for datasource in data['datasources']:
            document.add_heading(f"Data Source: {datasource.get('name', 'N/A')}", level=2)
            document.add_paragraph(f"Type: {datasource.get('type', 'N/A')}")
            document.add_paragraph(f"Connection String: {datasource.get('connectionString', 'N/A')}")
            document.add_paragraph(f"Datasource Type: {datasource.get('datasourceType', 'N/A')}")
            document.add_paragraph(f"Gateway: {datasource.get('gateway', 'N/A')}")
            document.add_paragraph(f"Server: {datasource.get('server', 'N/A')}")
            document.add_paragraph(f"Database: {datasource.get('database', 'N/A')}")


    if 'roles' in data and data['roles']:
        document.add_heading('Roles', level=1)
        for role in data['roles']:
            document.add_heading(f"Role: {role.get('name', 'N/A')}", level=2)
            document.add_paragraph(f"Model Permission: {role.get('modelPermission', 'N/A')}")
            if 'tablePermissions' in role and role['tablePermissions']:
                document.add_paragraph("Table Permissions:")
                for tp in role['tablePermissions']:
                    document.add_paragraph(f"- Table: {tp.get('name', 'N/A')}")
                    document.add_paragraph(f"  - Filter Expression: {tp.get('filterExpression', 'N/A')}")
                    document.add_paragraph(f"  - Rows Access: {tp.get('rowsAccess', 'N/A')}")

    if 'expressions' in data and data['expressions']:
        document.add_heading('Expressions', level=1)
        for expression in data['expressions']:
            document.add_heading(f"Expression: {expression.get('name', 'N/A')}", level=2)
            document.add_paragraph(f"Kind: {expression.get('kind', 'N/A')}")
            document.add_paragraph(f"Expression: {expression.get('expression', 'N/A')}")
            document.add_paragraph(f"Hidden: {expression.get('isHidden', 'N/A')}")
            document.add_paragraph(f"Lineage Tag: {expression.get('lineageTag', 'N/A')}")


    if 'annotations' in data and data['annotations']:
        document.add_heading('Annotations', level=1)
        for annotation in data['annotations']:
            document.add_heading(f"Annotation: {annotation.get('name', 'N/A')}", level=2)
            document.add_paragraph(f"Value: {annotation.get('value', 'N/A')}")

    # Save the document to a BytesIO object
    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer

def generate_pdf_doc(data):
    """Generates a PDF document from the parsed PBIX data."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size = 12)

    pdf.cell(200, 10, txt = "PBIX Documentation", ln = 1, align = 'C')

    if 'model' in data and data['model']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Model Information", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        model_data = data['model']
        pdf.multi_cell(0, 10, txt = f"Name: {model_data.get('name', 'N/A')}")
        pdf.multi_cell(0, 10, txt = f"Culture: {model_data.get('culture', 'N/A')}")
        pdf.multi_cell(0, 10, txt = f"Default Mode: {model_data.get('defaultMode', 'N/A')}")
        pdf.multi_cell(0, 10, txt = f"Description: {model_data.get('description', 'N/A')}")
        pdf.multi_cell(0, 10, txt = f"Modified Time: {model_data.get('modifiedTime', 'N/A')}")
        pdf.ln(5)


    if 'tables' in data and data['tables']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Tables", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for table in data['tables']:
            pdf.multi_cell(0, 10, txt = f"Table: {table['name']}")
            pdf.multi_cell(0, 10, txt = f"Description: {table.get('description', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Hidden: {table.get('isHidden', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Data Category: {table.get('dataCategory', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Lineage Tag: {table.get('lineageTag', 'N/A')}")

            if 'columns' in table and table['columns']:
                pdf.multi_cell(0, 10, txt = "Columns:")
                df_columns = pd.DataFrame(table['columns'])
                 # Convert dataframe to string and add to pdf
                pdf.multi_cell(0, 10, txt = df_columns.to_string(index=False))
            pdf.ln(5)

    if 'measures' in data and data['measures']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Measures", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for measure in data['measures']:
            pdf.multi_cell(0, 10, txt = f"Measure: {measure['name']}")
            pdf.multi_cell(0, 10, txt = f"Expression: {measure['expression']}")
            pdf.multi_cell(0, 10, txt = f"Description: {measure.get('description', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Display Folder: {measure.get('displayFolder', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Format String: {measure.get('formatString', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Hidden: {measure.get('isHidden', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Data Category: {measure.get('dataCategory', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Display Format: {measure.get('displayFormat', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Lineage Tag: {measure.get('lineageTag', 'N/A')}")
            pdf.ln(5)


    if 'relationships' in data and data['relationships']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Relationships", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for relationship in data['relationships']:
            pdf.multi_cell(0, 10, txt = f"Relationship: {relationship.get('name', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"From Table: {relationship.get('fromTable', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"From Column: {relationship.get('fromColumn', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"To Table: {relationship.get('toTable', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"To Column: {relationship.get('toColumn', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Join On Date Behavior: {relationship.get('joinOnDateBehavior', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"State: {relationship.get('state', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Type: {relationship.get('type', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Cross Filtering Behavior: {relationship.get('crossFilteringBehavior', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Active: {relationship.get('isActive', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Security Filtering Behavior: {relationship.get('securityFilteringBehavior', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Lineage Tag: {relationship.get('lineageTag', 'N/A')}")
            pdf.ln(5)

    if 'cultures' in data and data['cultures']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Cultures", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for culture in data['cultures']:
            pdf.multi_cell(0, 10, txt = f"Culture: {culture.get('name', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Language: {culture.get('language', 'N/A')}")
            if 'linguisticMetadata' in culture and culture['linguisticMetadata']:
                pdf.multi_cell(0, 10, txt = "Linguistic Metadata:")
                pdf.multi_cell(0, 10, txt = json.dumps(culture['linguisticMetadata'], indent=2))
            pdf.ln(5)

    if 'datasources' in data and data['datasources']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Data Sources", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for datasource in data['datasources']:
            pdf.multi_cell(0, 10, txt = f"Data Source: {datasource.get('name', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Type: {datasource.get('type', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Connection String: {datasource.get('connectionString', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Datasource Type: {datasource.get('datasourceType', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Gateway: {datasource.get('gateway', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Server: {datasource.get('server', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Database: {datasource.get('database', 'N/A')}")
            pdf.ln(5)

    if 'roles' in data and data['roles']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Roles", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for role in data['roles']:
            pdf.multi_cell(0, 10, txt = f"Role: {role.get('name', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Model Permission: {role.get('modelPermission', 'N/A')}")
            if 'tablePermissions' in role and role['tablePermissions']:
                pdf.multi_cell(0, 10, txt = "Table Permissions:")
                for tp in role['tablePermissions']:
                    pdf.multi_cell(0, 10, txt = f"- Table: {tp.get('name', 'N/A')}")
                    pdf.multi_cell(0, 10, txt = f"  - Filter Expression: {tp.get('filterExpression', 'N/A')}")
                    pdf.multi_cell(0, 10, txt = f"  - Rows Access: {tp.get('rowsAccess', 'N/A')}")
            pdf.ln(5)

    if 'expressions' in data and data['expressions']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Expressions", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for expression in data['expressions']:
            pdf.multi_cell(0, 10, txt = f"Expression: {expression.get('name', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Kind: {expression.get('kind', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Expression: {expression.get('expression', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Hidden: {expression.get('isHidden', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Lineage Tag: {expression.get('lineageTag', 'N/A')}")
            pdf.ln(5)


    if 'annotations' in data and data['annotations']:
        pdf.add_page()
        pdf.set_font("Arial", style='B', size = 12)
        pdf.cell(200, 10, txt = "Annotations", ln = 1, align = 'L')
        pdf.set_font("Arial", size = 10)
        for annotation in data['annotations']:
            pdf.multi_cell(0, 10, txt = f"Annotation: {annotation.get('name', 'N/A')}")
            pdf.multi_cell(0, 10, txt = f"Value: {annotation.get('value', 'N/A')}")
            pdf.ln(5)


    # Save the PDF to a BytesIO object
    buffer = BytesIO()
    pdf.output(buffer, 'S')
    buffer.seek(0)
    return buffer


st.title("PBIX File Analyzer")
st.write("Upload a PBIX file to analyze its structure, including tables, measures, relationships, and more.")

uploaded_file = st.file_uploader("Choose a .pbix file", type="pbix")

if uploaded_file is not None:
    # Save the uploaded file to a temporary location
    temp_file_path = "temp.pbix"
    with open(temp_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("File uploaded successfully! Analyzing...")

    # Run pbixray on the uploaded file
    pbix_data = run_pbixray(temp_file_path)

    # Clean up the temporary file
    os.remove(temp_file_path)

    if pbix_data:
        st.sidebar.header("Sections")
        sections = []
        if 'model' in pbix_data and pbix_data['model']:
            sections.append('Model Information')
        if 'tables' in pbix_data and pbix_data['tables']:
            sections.append('Tables')
        if 'measures' in pbix_data and pbix_data['measures']:
            sections.append('Measures')
        if 'relationships' in pbix_data and pbix_data['relationships']:
            sections.append('Relationships')
        if 'cultures' in pbix_data and pbix_data['cultures']:
             sections.append('Cultures')
        if 'datasources' in pbix_data and pbix_data['datasources']:
            sections.append('Data Sources')
        if 'roles' in pbix_data and pbix_data['roles']:
            sections.append('Roles')
        if 'expressions' in pbix_data and pbix_data['expressions']:
            sections.append('Expressions')
        if 'annotations' in pbix_data and pbix_data['annotations']:
            sections.append('Annotations')

        if sections:
            selected_section = st.sidebar.radio("Navigate to:", sections)

            st.header("Analysis Results")
            st.markdown("---") # Add a separator

            if selected_section == 'Model Information' and 'model' in pbix_data:
                display_model_info(pbix_data['model'])
            elif selected_section == 'Tables' and 'tables' in pbix_data:
                display_tables(pbix_data['tables'])
            elif selected_section == 'Measures' and 'measures' in pbix_data:
                display_measures(pbix_data['measures'])
            elif selected_section == 'Relationships' and 'relationships' in pbix_data:
                display_relationships(pbix_data['relationships'])
            elif selected_section == 'Cultures' and 'cultures' in pbix_data:
                display_cultures(pbix_data['cultures'])
            elif selected_section == 'Data Sources' and 'datasources' in pbix_data:
                display_datasources(pbix_data['datasources'])
            elif selected_section == 'Roles' and 'roles' in pbix_data:
                display_roles(pbix_data['roles'])
            elif selected_section == 'Expressions' and 'expressions' in pbix_data:
                display_expressions(pbix_data['expressions'])
            elif selected_section == 'Annotations' and 'annotations' in pbix_data:
                display_annotations(pbix_data['annotations'])

            st.markdown("---") # Add a separator
            st.header("Export Options")
            col1, col2 = st.columns(2)

            with col1:
                word_buffer = generate_word_doc(pbix_data)
                st.download_button(
                    label="Export to Word",
                    data=word_buffer,
                    file_name="pbix_documentation.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            with col2:
                pdf_buffer = generate_pdf_doc(pbix_data)
                st.download_button(
                    label="Export to PDF",
                    data=pdf_buffer,
                    file_name="pbix_documentation.pdf",
                    mime="application/pdf"
                )
        else:
            st.warning("No data found in the PBIX file to display.")

    else:
        st.error("Could not analyze the .pbix file.")
else:
    st.info("Please upload a .pbix file to get started.")

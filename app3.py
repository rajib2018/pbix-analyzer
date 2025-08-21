import streamlit as st
import subprocess
import json
import pandas as pd
import os

def run_pbixray(pbix_file_path):
    """Runs pbixray on the given file and returns the parsed JSON output."""
    try:
        # Construct the command to run pbixray
        command = ["pbixray", "-f", pbix_file_path]

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

def display_tables(tables_data):
    """Displays the tables information in a user-friendly format."""
    st.header("Tables")
    for table in tables_data:
        st.subheader(f"Table: {table['name']}")
        st.write(f"Description: {table.get('description', 'N/A')}")
        st.write(f"Hidden: {table.get('isHidden', 'N/A')}")
        st.write(f"Data Category: {table.get('dataCategory', 'N/A')}")

        if 'columns' in table:
            st.write("Columns:")
            columns_df = pd.DataFrame(table['columns'])
            st.dataframe(columns_df)
        else:
            st.write("No columns found for this table.")

def display_measures(measures_data):
    """Displays the measures information."""
    st.header("Measures")
    for measure in measures_data:
        st.subheader(f"Measure: {measure['name']}")
        st.write(f"Expression: ```\n{measure['expression']}\n```")
        st.write(f"Description: {measure.get('description', 'N/A')}")
        st.write(f"Display Folder: {measure.get('displayFolder', 'N/A')}")
        st.write(f"Format String: {measure.get('formatString', 'N/A')}")
        st.write(f"Hidden: {measure.get('isHidden', 'N/A')}")

def display_relationships(relationships_data):
    """Displays the relationships information."""
    st.header("Relationships")
    for relationship in relationships_data:
        st.write(f"Name: {relationship.get('name', 'N/A')}")
        st.write(f"From Table: {relationship.get('fromTable', 'N/A')}")
        st.write(f"From Column: {relationship.get('fromColumn', 'N/A')}")
        st.write(f"To Table: {relationship.get('toTable', 'N/A')}")
        st.write(f"To Column: {relationship.get('toColumn', 'N/A')}")
        st.write(f"Join On Date Behavior: {relationship.get('joinOnDateBehavior', 'N/A')}")
        st.write(f"State: {relationship.get('state', 'N/A')}")
        st.write(f"Type: {relationship.get('type', 'N/A')}")
        st.write("---")


st.title("PBIX File Analyzer")

uploaded_file = st.file_uploader("Upload a .pbix file", type="pbix")

if uploaded_file is not None:
    # Save the uploaded file to a temporary location
    with open("temp.pbix", "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("File uploaded successfully!")

    # Run pbixray on the uploaded file
    pbix_data = run_pbixray("temp.pbix")

    # Clean up the temporary file
    os.remove("temp.pbix")

    if pbix_data:
        st.header("Analysis Results")

        # Display different sections based on the parsed data
        if 'tables' in pbix_data:
            display_tables(pbix_data['tables'])

        if 'measures' in pbix_data:
            display_measures(pbix_data['measures'])

        if 'relationships' in pbix_data:
            display_relationships(pbix_data['relationships'])

        # You can add more sections here based on the pbixray output structure
        # e.g., display_pages, display_visuals, etc.
    else:
        st.error("Could not analyze the .pbix file.")

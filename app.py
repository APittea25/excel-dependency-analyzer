import streamlit as st
import pandas as pd
import openpyxl
import graphviz
import re
import io
from pathlib import Path

# App title
st.title("📂 Multi-File Excel Dependency Analyzer")

# File uploader for multiple Excel files
uploaded_files = st.file_uploader(
    "Upload multiple Excel files", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} files uploaded successfully!")

    # Store sheet dependencies
    file_dependencies = {}

    # Read and process each uploaded file
    excel_data = {}
    file_names = [uploaded_file.name for uploaded_file in uploaded_files]  # Store full file names
    file_index_map = {f"[{i+1}]": uploaded_file.name for i, uploaded_file in enumerate(uploaded_files)}  # Map [1], [2] to real filenames

    st.write("📂 Uploaded files detected:", file_names)  # Debugging: List uploaded files
    st.write("📊 Excel Reference Mapping:", file_index_map)  # Debugging: Show [1] -> filename mapping

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_dependencies[file_name] = set()  # Initialize dependency storage

        # Read the Excel file
        file_stream = io.BytesIO(uploaded_file.read())  # Convert to BytesIO for openpyxl
        wb = openpyxl.load_workbook(file_stream, data_only=False)
        excel_data[file_name] = wb

    # Analyze formulas and detect dependencies
    for file_name, wb in excel_data.items():
        st.write(f"🔍 Scanning file: {file_name}")  # Debugging: Show file being scanned

        for sheet in wb.sheetnames:
            ws = wb[sheet]

            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_text = cell.value

                        # Debugging: Show detected formula
                        st.write(f"📊 Formula found in {file_name} - {sheet}: `{formula_text}`")

                        # Extract file references (handle both filenames and numeric placeholders [1])
                        match = re.search(r"\[(.*?)\]", formula_text)
                        if match:
                            referenced_file = match.group(1)  # Extracted reference (could be filename or [1])

                            # **Map [1], [2], etc., to actual filenames**
                            resolved_filename = file_index_map.get(f"[{referenced_file}]", referenced_file)

                            # Debugging: Show extracted file reference
                            st.write(f"🔗 Formula references `{referenced_file}`, resolved to `{resolved_filename}`")

                            # Ensure the referenced file exists in uploaded files
                            if resolved_filename in file_names and resolved_filename != file_name:
                                file_dependencies[file_name].add(resolved_filename)  # Store dependency

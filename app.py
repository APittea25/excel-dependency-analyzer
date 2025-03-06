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

    st.write("📂 Uploaded files detected:", file_names)  # Debugging: List uploaded files

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

                        # Check if the formula references any of the other uploaded files
                        for potential_reference in file_names:
                            file_stem = Path(potential_reference).stem  # Get filename without extension
                            if file_stem.lower() in formula_text.lower() and potential_reference != file_name:
                                file_dependencies[file_name].add(potential_reference)  # Store dependency
                                st.write(f"✅ Link created: `{file_name}` → `{potential_reference}` (Partial match found)")

    # **Ensure all files appear in Graphviz, even isolated ones**
    all_files = set(file_dependencies.keys()).union(*file_dependencies.values())

    # **Check and debug dependencies before drawing graph**
    st.write("📋 Final Detected Dependencies:", file_dependencies)

    # **Create Dependency Flowchart**
    st.write("### 🔄 Spreadsheet Dependency Flowchart")
    flow = graphviz.Digraph(format="png")

    # **Ensure every file is added as a node**
    for file in all_files:
        flow.node(file)

    # **Force Graphviz to draw edges properly**
    has_edges = False  # Track if edges exist
    for file, dependencies in file_dependencies.items():
        for dependency in dependencies:
            flow.edge(dependency, file)  # Draw arrows
            has_edges = True  # Confirm edges exist

    # **If no edges were created, show a message**
    if not has_edges:
        st.warning("⚠️ No dependencies detected between uploaded spreadsheets.")
    else:
        st.graphviz_chart(flow)

    # **Show dependency table**
    st.write("### 📊 Dependency Table")
    dependency_df = pd.DataFrame(
        [(file, dep) for file, deps in file_dependencies.items() for dep in deps],
        columns=["File", "Depends On"]
    )

    if dependency_df.empty:
        st.write("✅ No direct dependencies found between uploaded Excel files.")
    else:
        st.dataframe(dependency_df)

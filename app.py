import streamlit as st
import pandas as pd
import openpyxl
import graphviz
import zipfile
import os
import re

# App title
st.title("📂 Excel Dependency Analyzer")

# File uploader for ZIP folder
uploaded_zip = st.file_uploader("Upload a ZIP file containing Excel spreadsheets", type=["zip"])

if uploaded_zip:
    st.success(f"✅ File '{uploaded_zip.name}' uploaded successfully!")

    # Extract ZIP file
    extract_folder = "extracted_files"
    os.makedirs(extract_folder, exist_ok=True)

    # Extract all contents
    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        zip_ref.extractall(extract_folder)

    # Debug: Show extracted files
    extracted_files = os.listdir(extract_folder)
    st.write("📁 Extracted files:", extracted_files)  # Debugging step

    # Find all Excel files (ensure correct extensions)
    excel_files = [f for f in extracted_files if f.lower().endswith((".xlsx", ".xls"))]

    if not excel_files:
        st.error("⚠️ No Excel files found in the uploaded ZIP. Please ensure your ZIP contains valid .xlsx or .xls files at the top level.")
        st.stop()

    # Store sheet dependencies
    file_dependencies = {}

    # Process each Excel file
    for file in excel_files:
        file_path = os.path.join(extract_folder, file)
        wb = openpyxl.load_workbook(file_path, data_only=False)
        file_dependencies[file] = set()

        # Check for sheet-to-sheet references within and across files
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        # Check for references to other Excel files
                        for other_file in excel_files:
                            if other_file != file and re.search(rf'\b{other_file[:-5]}!', cell.value, re.IGNORECASE):
                                file_dependencies[file].add(other_file)

    # Generate dependency flowchart
    st.write("### 🔄 Spreadsheet Dependency Flowchart")
    flow = graphviz.Digraph()

    # Add nodes and edges
    for file in file_dependencies:
        flow.node(file)
        for dependency in file_dependencies[file]:
            flow.edge(dependency, file)

    # Display the flowchart
    st.graphviz_chart(flow)

    # Show detected dependencies
    st.write("### 📊 Dependency Table")
    dependency_df = pd.DataFrame(
        [(file, dep) for file, deps in file_dependencies.items() for dep in deps],
        columns=["File", "Depends On"]
    )
    if dependency_df.empty:
        st.write("✅ No direct dependencies found between Excel files.")
    else:
        st.dataframe(dependency_df)

    # Clean up extracted files after execution (Optional)
    import shutil
    shutil.rmtree(extract_folder)

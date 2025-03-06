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
    file_stems = {Path(name).stem: name for name in file_names}  # Create mapping of filename without extension

    st.write("📂 Uploaded files detected:", file_names)  # Debugging: List uploaded files

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_dependencies[file_name] = set()

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

                        # Extract only the spreadsheet name from the full SharePoint/OneDrive path
                        match = re.search(r"\[(.*?)\]", formula_text)  # Extract part inside square brackets [filename.xlsx]
                        if match:
                            referenced_file = match.group(1)  # Extracted file name
                            referenced_stem = Path(referenced_file).stem  # Get just the name without extension

                            # Debugging: Show extracted file reference
                            st.write(f"🔗 Formula references file: `{referenced_file}` (Detected as `{referenced_stem}`)")

                            # Check if this referenced file exists in uploaded files
                            if referenced_stem in file_stems and file_stems[referenced_stem] != file_name:
                                file_dependencies[file_name].add(file_stems[referenced_stem])
                                st.write(f"✅ Link created: `{file_name}` → `{file_stems[referenced_stem]}`")

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
        st.write("✅ No direct dependencies found between uploaded Excel files.")
    else:
        st.dataframe(dependency_df)

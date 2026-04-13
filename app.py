import streamlit as st
import pandas as pd
import io
import os

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    # ---------------- READ FILE ----------------
    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.dataframe(df.head())

    # ---------------- CLEAN COLUMNS ----------------
    df.columns = [str(c).strip() for c in df.columns]

    wire_col = "Wireno"
    connections = {}

    # ---------------- PROCESS DATA ----------------
    for _, row in df.iterrows():

        wire = row[wire_col]

        if pd.isna(wire):
            continue

        wire = str(wire).strip()
        row_id = row.name

        if wire not in connections:
            connections[wire] = set()

        # ---------------- START SIDE ----------------
        start_component = row["Name"]
        start_conn = row["C.name"]

        if pd.notna(start_component):
            start_component = str(start_component).strip()

            if pd.notna(start_conn):
                start_conn = str(start_conn).strip()
                connections[wire].add(f"{start_component}|{start_conn}")
            else:
                connections[wire].add(f"{start_component}|MISSING_START_{row_id}")

        # ---------------- END SIDE ----------------
        end_component = row["Name.1"] if "Name.1" in df.columns else row["Name"]
        end_conn = row["C.name.1"] if "C.name.1" in df.columns else row["C.name"]

        if pd.notna(end_component):
            end_component = str(end_component).strip()

            if pd.notna(end_conn):
                end_conn = str(end_conn).strip()
                connections[wire].add(f"{end_component}|{end_conn}")
            else:
                connections[wire].add(f"{end_component}|MISSING_END_{row_id}")

    # ---------------- RESULT ----------------
    result = pd.DataFrame([
        {"Wire": wire, "Markings": len(values)}
        for wire, values in connections.items()
    ])

    result = result.sort_values("Wire")

    st.subheader("Result")
    st.dataframe(result)

    st.success(f"Total wires: {len(result)}")
    st.success(f"Total markings needed: {result['Markings'].sum()}")

    # ---------------- DOWNLOAD EXCEL ----------------
    original_name = uploaded_file.name
    base_name = os.path.splitext(original_name)[0]
    project_code = base_name.split()[0]

    download_name = f"{project_code} laidų žymėjimai.xlsx"

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Markings")

    output.seek(0)

    st.download_button(
        "Download Excel",
        output,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

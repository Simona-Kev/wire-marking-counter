import streamlit as st
import pandas as pd

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    # Read Excel safely
    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.dataframe(df.head())

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    wire_col = "Wireno"

    connections = {}

    for _, row in df.iterrows():

        wire = row[wire_col]

        if pd.isna(wire):
            continue

        wire = str(wire).strip()
        row_id = row.name  # important for unique missing values

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
                # unique missing marker (prevents collapsing)
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
                # unique missing marker (prevents collapsing)
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

    st.download_button(
        "Download CSV",
        result.to_csv(index=False),
        "wire_markings.csv",
        "text/csv"
    )

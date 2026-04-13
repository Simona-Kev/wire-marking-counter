import streamlit as st
import pandas as pd

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

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

        if wire not in connections:
            connections[wire] = set()

        # START side (B + C)
        start_component = row["Name"]
        start_conn = row["C.name"]

        if pd.notna(start_component) or pd.notna(start_conn):
            connections[wire].add(
                f"{str(start_component).strip()}|{str(start_conn).strip()}"
            )

        # END side (D + E)
        end_component = row["Name.1"] if "Name.1" in df.columns else row["Name"]
        end_conn = row["C.name.1"] if "C.name.1" in df.columns else row["C.name"]

        if pd.notna(end_component) or pd.notna(end_conn):
            connections[wire].add(
                f"{str(end_component).strip()}|{str(end_conn).strip()}"
            )

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

import streamlit as st
import pandas as pd

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    # Read Excel safely (xlsx + xls)
    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.write(df.head())

    # First column = wire ID
    wire_col = df.columns[0]

    # Only columns called "Name"
    name_cols = [col for col in df.columns if col.strip() == "Name"]

    if not name_cols:
        st.error("No column named 'Name' found in file.")
        st.stop()

    connections = {}

    for _, row in df.iterrows():
        wire = row[wire_col]

        if pd.isna(wire):
            continue

        wire = str(wire).strip()

        if wire not in connections:
            connections[wire] = set()

        for col in name_cols:
            value = row[col]

            if pd.notna(value):
                connections[wire].add(str(value).strip())

    # Build result table
    result = pd.DataFrame([
        {"Wire": wire, "Markings": len(values)}
        for wire, values in connections.items()
    ])

    result = result.sort_values("Wire")

    st.subheader("Result")
    st.dataframe(result)

    st.success(f"Total wires: {len(result)}")
    st.success(f"Total markings needed: {result['Markings'].sum()}")

    # Download
    st.download_button(
        "Download CSV",
        result.to_csv(index=False),
        "wire_markings.csv",
        "text/csv"
    )

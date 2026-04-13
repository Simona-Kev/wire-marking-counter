import streamlit as st
import pandas as pd

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    # Read file
    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.dataframe(df.head())

    # Wire column = first column
    wire_col = df.columns[0]

    # CLEAN column names (important fix!)
    df.columns = [str(c).strip() for c in df.columns]

    # Find ALL "Name" columns (case-insensitive)
    name_cols = [col for col in df.columns if str(col).strip().lower() == "name"]

    st.write("Detected Name columns:", name_cols)

    if not name_cols:
        st.error("No 'Name' columns found.")
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
                val = str(value).strip()
                if val:
                    connections[wire].add(val)

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

import streamlit as st
import pandas as pd

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"]
)

if uploaded_file.name.endswith(".xls"):
    df = pd.read_excel(uploaded_file, engine="xlrd")
else:
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.write(df)

    # First column = wire number
    wire_col = df.columns[0]

    # All other columns = connection points
    connection_cols = df.columns[1:]

    connections = {}

for _, row in df.iterrows():
    wire = row[wire_col]

    if wire not in connections:
        connections[wire] = set()

    # Process columns in pairs (Name + C name)
    cols = list(connection_cols)

    for i in range(0, len(cols), 2):
        val1 = row[cols[i]]
        val2 = row[cols[i+1]] if i+1 < len(cols) else None

        if pd.notna(val1) or pd.notna(val2):
            # Treat pair as one connection
            pair = f"{val1}-{val2}"
            connections[wire].add(pair)
    # Build result table
    result = pd.DataFrame([
        {"Wire": wire, "Markings": len(points)}
        for wire, points in connections.items()
    ])

    st.subheader("Result")
    st.write(result)

    total = result["Markings"].sum()
    st.success(f"Total markings needed: {total}")

    # Download button
    st.download_button(
        "Download results",
        result.to_csv(index=False),
        "wire_markings.csv",
        "text/csv"
    )

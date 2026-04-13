import streamlit as st
import pandas as pd

st.title("Wire Marking Counter")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

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

        for col in connection_cols:
            value = row[col]

            if pd.notna(value):  # ignore empty cells
                connections[wire].add(str(value).strip())

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

import streamlit as st
import pandas as pd

st.title("Laidų žymėjimų skaičiuoklė")

uploaded_file = st.file_uploader(
    "Įkelkite excel failą:",
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
    connections = {}

# only keep columns called "Name"
name_cols = [col for col in df.columns if col.strip() == "Name"]

for _, row in df.iterrows():
    wire = row[df.columns[0]]  # first column = wire number

    if wire not in connections:
        connections[wire] = set()

    for col in name_cols:
        value = row[col]

        if pd.notna(value):
            connections[wire].add(str(value).strip())

# build result
result = pd.DataFrame([
    {"Wire": wire, "Markings": len(values)}
    for wire, values in connections.items()
])

st.write(result)
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

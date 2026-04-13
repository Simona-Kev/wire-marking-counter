import streamlit as st
import pandas as pd
import io
import os
import re
from streamlit_sortables import sort_items

st.title("Wire Marking Counter (Drag & Drop Sorting)")

# ---------------- DEFAULT SORT ORDER ----------------
if "rules" not in st.session_state:
    st.session_state.rules = [
        "1L",
        "L",
        "N",
        "24V",
        "0V",
        "S_0V",
        "A",
        "X",
        "Y"
    ]

# ---------------- DRAG & DROP UI ----------------
st.sidebar.header("Sorting Priority (Drag to reorder)")

st.session_state.rules = sort_items(
    st.session_state.rules,
    direction="vertical"
)

st.sidebar.write("Current order:")
st.sidebar.write(st.session_state.rules)

# Convert list → priority map
priority_map = {prefix: i for i, prefix in enumerate(st.session_state.rules)}

# ---------------- SORT FUNCTION ----------------
def natural_key(wire):
    wire = str(wire).strip().upper()

    def extract_numbers(text):
        nums = re.findall(r"\d+", text)
        return tuple(map(int, nums)) if nums else (0,)

    # match by prefix order from UI
    for prefix, priority in priority_map.items():
        if wire.startswith(prefix):
            return (priority, extract_numbers(wire))

    return (99, wire)


# ---------------- FILE UPLOAD ----------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.dataframe(df.head())

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

        # START
        start_component = row["Name"]
        start_conn = row["C.name"]

        if pd.notna(start_component):
            start_component = str(start_component).strip()

            if pd.notna(start_conn):
                start_conn = str(start_conn).strip()
                connections[wire].add(f"{start_component}|{start_conn}")
            else:
                connections[wire].add(f"{start_component}|MISSING_START_{row_id}")

        # END
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

    result["sort_key"] = result["Wire"].apply(natural_key)
    result = result.sort_values("sort_key").drop(columns=["sort_key"])

    st.subheader("Result")
    st.dataframe(result)

    st.success(f"Total wires: {len(result)}")
    st.success(f"Total markings needed: {result['Markings'].sum()}")

    # ---------------- DOWNLOAD FILE NAME ----------------
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

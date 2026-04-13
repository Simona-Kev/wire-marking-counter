import streamlit as st
import pandas as pd
import io
import os
import re

st.title("Wire Marking Counter")

# ---------------- NATURAL SORT FUNCTION ----------------
def natural_key(wire):
    wire = str(wire).strip().upper()

    def extract_numbers(text):
        nums = re.findall(r"\d+", text)
        return tuple(map(int, nums)) if nums else (0,)

    # 1L group
    if wire.startswith("1L"):
        return (1, extract_numbers(wire))

    # L group
    if wire.startswith("L") and not wire.startswith("1L"):
        return (2, extract_numbers(wire))

    # N group
    if wire.startswith("N"):
        return (3, extract_numbers(wire))

    # 24V group
    if wire.startswith("24V"):
        return (4, extract_numbers(wire))

    # 0V (always first in this group)
if wire == "0V":
    return (5, 0)

# S_0V base (without numbers)
if wire == "S_0V":
    return (6, 0)

# S_0V with numbers (S_0V1, S_0V2...)
if wire.startswith("S_0V"):
    nums = re.findall(r"\d+", wire)
    return (7, int(nums[0]) if nums else 0)

    # pure numbers
    if wire.isdigit():
        return (6, int(wire))

    # A group
    if wire.startswith("A"):
        return (7, extract_numbers(wire))

    # X group
    if wire.startswith("X"):
        return (8, extract_numbers(wire))

    # Y group
    if wire.startswith("Y"):
        return (9, extract_numbers(wire))

    # everything else
    return (99, wire)


# ---------------- UPLOAD ----------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    # ---------------- READ FILE ----------------
    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.subheader("Preview")
    st.dataframe(df.head())

    df.columns = [str(c).strip() for c in df.columns]

    wire_col = "Wireno"
    connections = {}

    # ---------------- PROCESS ----------------
    for _, row in df.iterrows():

        wire = row[wire_col]

        if pd.isna(wire):
            continue

        wire = str(wire).strip()
        row_id = row.name

        if wire not in connections:
            connections[wire] = set()

        # START side
        start_component = row["Name"]
        start_conn = row["C.name"]

        if pd.notna(start_component):
            start_component = str(start_component).strip()

            if pd.notna(start_conn):
                start_conn = str(start_conn).strip()
                connections[wire].add(f"{start_component}|{start_conn}")
            else:
                connections[wire].add(f"{start_component}|MISSING_START_{row_id}")

        # END side
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

    # ---------------- SORT (IMPORTANT UPGRADE) ----------------
    result["sort_key"] = result["Wire"].apply(natural_key)
    result = result.sort_values("sort_key").drop(columns=["sort_key"])

    st.subheader("Result")
    st.dataframe(result)

    st.success(f"Total wires: {len(result)}")
    st.success(f"Total markings needed: {result['Markings'].sum()}")

    # ---------------- FILE NAME ----------------
    original_name = uploaded_file.name
    base_name = os.path.splitext(original_name)[0]
    project_code = base_name.split()[0]

    download_name = f"{project_code} laidų žymėjimai.xlsx"

    # ---------------- EXPORT EXCEL ----------------
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Markings")

    output.seek(0)

    # ---------------- DOWNLOAD ----------------
    st.download_button(
        "Download Excel",
        output,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

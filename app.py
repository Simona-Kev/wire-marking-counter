import streamlit as st
import pandas as pd
import io
import os
import re

st.title("Wire Marking Counter (Dynamic Sorting)")

# ---------------- SESSION STATE (STORE RULES) ----------------
if "rules" not in st.session_state:
    st.session_state.rules = {
        "1L": 1,
        "L": 2,
        "N": 3,
        "24V": 4,
        "0V": 5,
        "S_0V": 6,
        "A": 7,
        "X": 8,
        "Y": 9
    }

# ---------------- SIDEBAR UI ----------------
st.sidebar.header("Sorting Rules")

st.sidebar.write("Edit priority (lower = higher priority)")

for key in list(st.session_state.rules.keys()):
    st.session_state.rules[key] = st.sidebar.number_input(
        f"{key}",
        value=st.session_state.rules[key],
        step=1
    )

# ➕ Add new rule
st.sidebar.subheader("Add new rule")

new_prefix = st.sidebar.text_input("Prefix (e.g. 24V, Z, PWR)")
new_priority = st.sidebar.number_input("Priority", value=10, step=1)

if st.sidebar.button("Add rule"):
    if new_prefix:
        st.session_state.rules[new_prefix.upper()] = new_priority
        st.rerun()

# ---------------- NATURAL SORT FUNCTION ----------------
def natural_key(wire):
    wire = str(wire).strip().upper()

    def extract_numbers(text):
        nums = re.findall(r"\d+", text)
        return tuple(map(int, nums)) if nums else (0,)

    # match rules dynamically
    for prefix, priority in st.session_state.rules.items():
        if wire.startswith(prefix):
            return (priority, extract_numbers(wire))

    return (99, wire)


# ---------------- UPLOAD ----------------
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

    # ---------------- PROCESS ----------------
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
                connections[wire].add(f"{start_component}|{str(start_conn).strip()}")
            else:
                connections[wire].add(f"{start_component}|MISSING_START_{row_id}")

        # END
        end_component = row["Name.1"] if "Name.1" in df.columns else row["Name"]
        end_conn = row["C.name.1"] if "C.name.1" in df.columns else row["C.name"]

        if pd.notna(end_component):
            end_component = str(end_component).strip()

            if pd.notna(end_conn):
                connections[wire].add(f"{end_component}|{str(end_conn).strip()}")
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

    # ---------------- EXPORT ----------------
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

import streamlit as st
import pandas as pd
import io
import os
import re
import json

st.title("Wire Marking Counter (Final Stable Version)")

# ---------------- STORAGE ----------------
RULES_FILE = "rules.json"

DEFAULT_RULES = [
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

def load_rules():
    if os.path.exists(RULES_FILE):
        with open(RULES_FILE, "r") as f:
            return json.load(f)
    return DEFAULT_RULES

def save_rules(rules):
    with open(RULES_FILE, "w") as f:
        json.dump(rules, f)

if "rules" not in st.session_state:
    st.session_state.rules = load_rules()

# ---------------- PRIORITY MAP ----------------
priority_map = {prefix: i for i, prefix in enumerate(st.session_state.rules)}

# ---------------- SAFE SORT FUNCTION ----------------
def natural_key(wire):
    wire = str(wire).strip().upper()

    def nums(text):
        found = re.findall(r"\d+", text)
        return [int(x) for x in found] if found else []

    # detect prefix group
    for prefix, priority in priority_map.items():
        if wire.startswith(prefix):

            n = nums(wire)

            # X / Y special handling
            if prefix in ["X", "Y"]:
                if len(n) >= 2:
                    return (priority, n[0], n[1])
                elif len(n) == 1:
                    return (priority, n[0], 0)
                else:
                    return (priority, 0, 0)

            # normal groups
            if len(n) > 0:
                return (priority, n[0], 0)

            return (priority, 0, 0)

    # numbers last
    if wire.isdigit():
        return (999, int(wire), 0)

    return (999, 0, wire)

# ---------------- UPLOAD ----------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    if uploaded_file.name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    df.columns = [str(c).strip() for c in df.columns]

    wire_col = "Wireno"
    connections = {}

    # ---------------- PROCESS ----------------
    for _, row in df.iterrows():

        wire = row[wire_col]
        if pd.isna(wire):
            continue

        wire = str(wire).strip().upper()
        row_id = row.name

        if wire not in connections:
            connections[wire] = set()

        # START
        start_component = row["Name"]
        start_conn = row["C.name"]

        if pd.notna(start_component):
            if pd.notna(start_conn):
                connections[wire].add(f"{start_component}|{start_conn}")
            else:
                connections[wire].add(f"{start_component}|MISSING_START_{row_id}")

        # END
        end_component = row["Name.1"] if "Name.1" in df.columns else row["Name"]
        end_conn = row["C.name.1"] if "C.name.1" in df.columns else row["C.name"]

        if pd.notna(end_component):
            if pd.notna(end_conn):
                connections[wire].add(f"{end_component}|{end_conn}")
            else:
                connections[wire].add(f"{end_component}|MISSING_END_{row_id}")

    # ---------------- RESULT ----------------
    result = pd.DataFrame([
        {"Wire": wire, "Markings": len(values)}
        for wire, values in connections.items()
    ])

    # ---------------- SAFE SORT (FINAL FIX) ----------------
    result["sort_key"] = result["Wire"].apply(natural_key)

    # IMPORTANT: normalize EVERY key into tuple (guaranteed safe)
    def safe_key(x):
        k = x["sort_key"]
        if isinstance(k, tuple):
            return k
        if isinstance(k, list):
            return tuple(k)
        return (999, 0, str(k))

    sorted_rows = sorted(
        result.to_dict("records"),
        key=safe_key
    )

    result = pd.DataFrame(sorted_rows).drop(columns=["sort_key"])

    # ---------------- OUTPUT ----------------
    st.subheader("Result")
    st.dataframe(result)

    st.success(f"Total wires: {len(result)}")
    st.success(f"Total markings needed: {result['Markings'].sum()}")

    # ---------------- EXPORT ----------------
    original_name = uploaded_file.name
    base_name = os.path.splitext(original_name)[0]
    project_code = base_name.split()[0]

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Markings")

    output.seek(0)

    st.download_button(
        "Download Excel",
        output,
        file_name=f"{project_code} laidų žymėjimai.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

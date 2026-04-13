import streamlit as st
import pandas as pd
import io
import os
import re
import json
from streamlit_sortables import sort_items

st.title("Wire Marking Counter (Stable Version)")

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

# ---------------- SIDEBAR ----------------
st.sidebar.header("Sorting Rules (Drag & Drop)")

st.session_state.rules = sort_items(
    st.session_state.rules,
    direction="vertical"
)

st.sidebar.write("Current order:")
st.sidebar.write(st.session_state.rules)

# Add rule
new_rule = st.sidebar.text_input("Add new prefix")

if st.sidebar.button("➕ Add rule"):
    if new_rule:
        new_rule = new_rule.upper()
        if new_rule not in st.session_state.rules:
            st.session_state.rules.append(new_rule)
            st.rerun()

# Remove rule
remove_rule = st.sidebar.selectbox("Remove rule", st.session_state.rules)

if st.sidebar.button("❌ Remove rule"):
    st.session_state.rules.remove(remove_rule)
    st.rerun()

# Save / Reset
if st.sidebar.button("💾 Save rules"):
    save_rules(st.session_state.rules)
    st.sidebar.success("Saved!")

if st.sidebar.button("🔄 Reset"):
    st.session_state.rules = DEFAULT_RULES.copy()
    save_rules(st.session_state.rules)
    st.rerun()

# ---------------- PRIORITY MAP ----------------
priority_map = {prefix: i for i, prefix in enumerate(st.session_state.rules)}

# ---------------- SORT KEY (SAFE SIMPLE VERSION) ----------------
def natural_key(wire):
    wire = str(wire).strip().upper()

    def extract_all_numbers(text):
        nums = re.findall(r"\d+", text)
        return [int(n) for n in nums] if nums else []

    # ---------------- PREFIX MATCH ----------------
    for prefix, priority in priority_map.items():
        if wire.startswith(prefix):

            nums = extract_all_numbers(wire)

            # X / Y special handling (multi-number safe)
            if prefix in ["X", "Y"]:
                if len(nums) >= 2:
                    return (priority, nums[0], nums[1], wire)
                elif len(nums) == 1:
                    return (priority, nums[0], 0, wire)
                else:
                    return (priority, 0, 0, wire)

            # normal groups (24V, S_0V, A, etc.)
            if nums:
                return (priority, nums[0], wire)

            return (priority, 0, wire)

    # ---------------- NUMBERS ONLY (LAST GROUP) ----------------
    if wire.isdigit():
        return (999, int(wire))

    return (999, wire)


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

    # ---------------- SAFE SORT (BACK TO SIMPLE PANDAS) ----------------
    result["sort_key"] = result["Wire"].apply(natural_key)
    result = result.sort_values("sort_key", kind="stable").drop(columns=["sort_key"])

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

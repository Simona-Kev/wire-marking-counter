import streamlit as st
import pandas as pd
import io
import os
import re
import json
from streamlit_sortables import sort_items

st.title("Wire Marking Counter (Fully Fixed)")

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

# ---------------- SESSION INIT ----------------
if "rules" not in st.session_state:
    st.session_state.rules = load_rules()

# ---------------- FORCE UPDATE HELPERS ----------------
def force_rerun():
    st.session_state["_rerun_trigger"] = not st.session_state.get("_rerun_trigger", False)

# ---------------- SIDEBAR ----------------
st.sidebar.header("Sorting Rules")

new_rules = sort_items(
    st.session_state.rules,
    direction="vertical"
)

# IMPORTANT: update ONLY if changed
if new_rules != st.session_state.rules:
    st.session_state.rules = new_rules
    force_rerun()

st.sidebar.write("Current order:")
st.sidebar.write(st.session_state.rules)

# Add rule
new_rule = st.sidebar.text_input("Add rule")

if st.sidebar.button("➕ Add"):
    if new_rule:
        r = new_rule.upper()
        if r not in st.session_state.rules:
            st.session_state.rules.append(r)
            force_rerun()

# Remove rule
remove_rule = st.sidebar.selectbox("Remove rule", st.session_state.rules)

if st.sidebar.button("❌ Remove"):
    st.session_state.rules.remove(remove_rule)
    force_rerun()

# Save
if st.sidebar.button("💾 Save"):
    save_rules(st.session_state.rules)
    st.sidebar.success("Saved!")

# Reset (FIXED)
if st.sidebar.button("🔄 Reset"):
    st.session_state.rules = DEFAULT_RULES.copy()
    save_rules(st.session_state.rules)
    force_rerun()

# ---------------- PRIORITY MAP ----------------
priority_map = {prefix: i for i, prefix in enumerate(st.session_state.rules)}

# ---------------- FIXED SORT FUNCTION ----------------
def natural_key(wire):
    wire = str(wire).strip().upper()

    def nums(text):
        found = re.findall(r"\d+", text)
        return [int(x) for x in found] if found else []

    for prefix, priority in priority_map.items():
        if wire.startswith(prefix):

            n = nums(wire)

            # ---------------- FIX 24V / S_0V ----------------
            # base first, then _1 _2 _10 correctly
            if prefix in ["24V", "S_0V"]:

                if wire == prefix:
                    return (priority, 0, 0)

                if n:
                    # extract suffix order safely
                    suffix = n[0]
                    return (priority, 1, suffix)

                return (priority, 0, 0)

            # ---------------- X / Y ----------------
            if prefix in ["X", "Y"]:
                if len(n) >= 2:
                    return (priority, n[0], n[1])
                elif len(n) == 1:
                    return (priority, n[0], 0)
                return (priority, 0, 0)

            # ---------------- NORMAL ----------------
            if n:
                return (priority, n[0], 0)

            return (priority, 0, 0)

    # numbers last (true numeric sort)
    if wire.isdigit():
        return (999, int(wire), 0)

    return (999, 0, wire)

# ---------------- UPLOAD ----------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:

    df = pd.read_excel(uploaded_file, engine="xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    wire_col = "Wireno"
    connections = {}

    for _, row in df.iterrows():

        wire = row[wire_col]
        if pd.isna(wire):
            continue

        wire = str(wire).strip().upper()
        row_id = row.name

        if wire not in connections:
            connections[wire] = set()

        # START
        sc = row["Name"]
        scc = row["C.name"]

        if pd.notna(sc):
            connections[wire].add(f"{sc}|{scc if pd.notna(scc) else 'MISSING_START_'+str(row_id)}")

        # END
        ec = row["Name.1"] if "Name.1" in df.columns else row["Name"]
        ecc = row["C.name.1"] if "C.name.1" in df.columns else row["C.name"]

        if pd.notna(ec):
            connections[wire].add(f"{ec}|{ecc if pd.notna(ecc) else 'MISSING_END_'+str(row_id)}")

    # ---------------- RESULT ----------------
    result = pd.DataFrame([
        {"Wire": w, "Markings": len(v)}
        for w, v in connections.items()
    ])

    result["sort_key"] = result["Wire"].apply(natural_key)

    result = pd.DataFrame(
        sorted(result.to_dict("records"), key=lambda x: x["sort_key"])
    ).drop(columns=["sort_key"])

    # ---------------- OUTPUT ----------------
    st.subheader("Result")
    st.dataframe(result)

    st.success(f"Total wires: {len(result)}")
    st.success(f"Total markings: {result['Markings'].sum()}")

    # ---------------- EXPORT ----------------
    base = os.path.splitext(uploaded_file.name)[0].split()[0]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False)

    output.seek(0)

    st.download_button(
        "Download Excel",
        output,
        file_name=f"{base} laidų žymėjimai.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

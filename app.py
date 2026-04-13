import streamlit as st
import pandas as pd
import io
import os
import re
import json
from streamlit_sortables import sort_items

st.title("Wire & Component Tools")

# ---------------- MODE ----------------
mode = st.radio(
    "Select tool",
    ["Wire Marking Counter", "Component Marking Cleaner"]
)

# =========================================================
# 🔌 WIRE TOOL
# =========================================================
if mode == "Wire Marking Counter":

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
    st.sidebar.header("Sorting Rules")

    new_rules = sort_items(
        st.session_state.rules,
        direction="vertical"
    )

    if new_rules != st.session_state.rules:
        st.session_state.rules = new_rules
        st.rerun()

    st.sidebar.write(st.session_state.rules)

    # ---------------- PRIORITY MAP ----------------
    priority_map = {p: i for i, p in enumerate(st.session_state.rules)}

    # ---------------- SORT FUNCTION ----------------
    def natural_key(wire):
        wire = str(wire).strip().upper()

        def nums(text):
            found = re.findall(r"\d+", text)
            return [int(x) for x in found] if found else []

        for prefix, priority in priority_map.items():
            if wire.startswith(prefix):

                n = nums(wire)

                if prefix in ["24V", "S_0V"]:
                    if wire == prefix:
                        return (priority, 0, 0)
                    if n:
                        return (priority, 1, n[0])
                    return (priority, 0, 0)

                if prefix in ["X", "Y"]:
                    if len(n) >= 2:
                        return (priority, n[0], n[1])
                    elif len(n) == 1:
                        return (priority, n[0], 0)
                    return (priority, 0, 0)

                if n:
                    return (priority, n[0], 0)

                return (priority, 0, 0)

        if wire.isdigit():
            return (999, int(wire), 0)

        return (999, 0, wire)

    # ---------------- UPLOAD ----------------
    uploaded_file = st.file_uploader("Upload Wire Excel", type=["xlsx", "xls"])

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

            if wire not in connections:
                connections[wire] = set()

            start_component = row["Name"]
            start_conn = row["C.name"]

            if pd.notna(start_component):
                connections[wire].add(
                    f"{start_component}|{start_conn if pd.notna(start_conn) else 'MISSING'}"
                )

            end_component = row["Name.1"] if "Name.1" in df.columns else row["Name"]
            end_conn = row["C.name.1"] if "C.name.1" in df.columns else row["C.name"]

            if pd.notna(end_component):
                connections[wire].add(
                    f"{end_component}|{end_conn if pd.notna(end_conn) else 'MISSING'}"
                )

        result = pd.DataFrame([
            {"Wire": w, "Markings": len(v)}
            for w, v in connections.items()
        ])

        result["sort_key"] = result["Wire"].apply(natural_key)

        result = pd.DataFrame(
            sorted(result.to_dict("records"), key=lambda x: x["sort_key"])
        ).drop(columns=["sort_key"])

        st.subheader("Result")
        st.dataframe(result)

        st.success(f"Total wires: {len(result)}")
        st.success(f"Total markings: {result['Markings'].sum()}")

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(writer, index=False)

        output.seek(0)

        base = os.path.splitext(uploaded_file.name)[0].split()[0]

        st.download_button(
            "Download Excel",
            output,
            file_name=f"{base} laidų žymėjimai.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =========================================================
# 🧩 COMPONENT TOOL (UNIQUE MARKINGS ONLY)
# =========================================================
if mode == "Component Marking Cleaner":

    st.subheader("Component Marking Cleaner")

    uploaded_file = st.file_uploader(
        "Upload Component Excel",
        type=["xlsx", "xls"],
        key="comp_upload"
    )

    if uploaded_file:

        df = pd.read_excel(uploaded_file, engine="xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl")
        df.columns = [str(c).strip() for c in df.columns]

        if "Name" not in df.columns:
            st.error("File must contain 'Name' column")
            st.stop()

        values = (
            df["Name"]
            .dropna()
            .astype(str)
            .str.strip()
        )

        unique_values = sorted(set(values))

        result = pd.DataFrame({"Marking": unique_values})

        st.dataframe(result)

        st.success(f"Total unique markings: {len(result)}")

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(writer, index=False)

        output.seek(0)

        st.download_button(
            "Download Unique Markings",
            output,
            file_name=f"{base} komponentų žymėjimai.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

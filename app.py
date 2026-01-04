# =========================================================
# SMART PURCHASE ORDER ENGINE – FINAL WORKING VERSION
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="Smart PO Engine", layout="wide")
st.title("📦 Smart Purchase Order Engine")

# ---------------------------------------------------------
# SIDEBAR CONFIG
# ---------------------------------------------------------
st.sidebar.header("⚙️ Configuration")

w7  = st.sidebar.slider("7 Days Weight", 0.0, 1.0, 0.35, 0.01)
w15 = st.sidebar.slider("15 Days Weight", 0.0, 1.0, 0.25, 0.01)
w30 = st.sidebar.slider("30 Days Weight", 0.0, 1.0, 0.20, 0.01)
w45 = st.sidebar.slider("45 Days Weight", 0.0, 1.0, 0.12, 0.01)
w60 = st.sidebar.slider("60 Days Weight", 0.0, 1.0, 0.08, 0.01)

plan_top = st.sidebar.number_input("Top200 / Hotcake Days", 10, 90, 45)
plan_pos = st.sidebar.number_input("Positive / New SKU Days", 10, 90, 38)
plan_def = st.sidebar.number_input("Default Days", 10, 90, 30)

box_threshold = st.sidebar.slider("Box Rounding Threshold", 0.5, 1.0, 0.8)

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if file:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    def num(col, default=0):
        return pd.to_numeric(df.get(col, default), errors="coerce").fillna(default)

    sales_cols = ["7 Days Sales","15 Days Sales","30 Days Sales","45 Days Sales","60 Days Sales"]
    for c in sales_cols:
        df[c] = num(c)

    df["Box Qty"] = num("Box Qty")
    df["Current Stock"] = num("Current Stock")
    df["Hold & Unbilled Stock"] = num("Hold & Unbilled Stock")
    df["TOTAL_STOCK"] = df["Current Stock"] + df["Hold & Unbilled Stock"]
    df["Rank"] = num("Top 500 SKU Rank", 9999)

    df["Review"] = df.get("Review","").astype(str)
    df["MOS"] = df.get("MOS-WH Available","").astype(str)

    # -----------------------------------------------------
    # AUTO SYSTEM QTY
    # -----------------------------------------------------
    def system_po(row):
        if row["MOS"] != "Yes" or row["Box Qty"] <= 0:
            return 0

        S7,S15,S30,S45,S60 = row[sales_cols]
        stock, box = row["TOTAL_STOCK"], row["Box Qty"]
        review, rank = row["Review"].lower(), row["Rank"]

        is_top = rank <= 200 or "hot" in review
        is_pos = review in ["positive","new sku"]

        daily = (
            (S7/7)*w7 + (S15/15)*w15 + (S30/30)*w30 +
            (S45/45)*w45 + (S60/60)*w60
        )

        days = plan_top if is_top else plan_pos if is_pos else plan_def
        target = daily * days
        shortage = max(target - stock, 0)

        if shortage <= 0:
            return 0

        rounded = np.ceil(shortage / box) * box
        return int(rounded)

    df["System Required Qty"] = df.apply(system_po, axis=1)

    # -----------------------------------------------------
    # MANUAL OVERRIDE COLUMN (USER INPUT)
    # -----------------------------------------------------
    if "Manual Required Qty" not in df.columns:
        df["Manual Required Qty"] = 0

    st.success("✅ Purchase Order Calculated Successfully")

    df = st.data_editor(
        df,
        use_container_width=True,
        disabled=[c for c in df.columns if c != "Manual Required Qty"]
    )

    # -----------------------------------------------------
    # FINAL QTY LOGIC (THIS WAS BUG – NOW FIXED)
    # -----------------------------------------------------
    df["Final Order Qty"] = np.where(
        df["Manual Required Qty"].astype(float) > 0,
        df["Manual Required Qty"],
        df["System Required Qty"]
    ).astype(int)

    # -----------------------------------------------------
    # DOWNLOAD
    # -----------------------------------------------------
    output = BytesIO()
    df.fillna("").to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        "⬇️ Download Final PO Excel",
        data=output,
        file_name="SMART_PO_FINAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Upload Excel to begin")

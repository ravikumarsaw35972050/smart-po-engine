# =========================================================
# SMART PO ENGINE + CONFIG UI (STREAMLIT – FINAL MATCH)
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
# SIDEBAR CONFIG (EXACT COLAB MATCH)
# ---------------------------------------------------------
st.sidebar.header("⚙️ Configuration")

w7  = st.sidebar.slider("7D wt", 0.0, 1.0, 0.35, 0.01)
w15 = st.sidebar.slider("15D wt", 0.0, 1.0, 0.25, 0.01)
w30 = st.sidebar.slider("30D wt", 0.0, 1.0, 0.20, 0.01)
w45 = st.sidebar.slider("45D wt", 0.0, 1.0, 0.12, 0.01)
w60 = st.sidebar.slider("60D wt", 0.0, 1.0, 0.08, 0.01)

ew15 = st.sidebar.slider("15D ex", 0.0, 1.0, 0.40, 0.01)
ew30 = st.sidebar.slider("30D ex", 0.0, 1.0, 0.30, 0.01)
ew45 = st.sidebar.slider("45D ex", 0.0, 1.0, 0.20, 0.01)
ew60 = st.sidebar.slider("60D ex", 0.0, 1.0, 0.10, 0.01)

threshold = st.sidebar.slider("Box Threshold", 0.5, 1.0, 0.8)

plan_top = st.sidebar.slider("Top200 Days", 15, 60, 45)
plan_pos = st.sidebar.slider("Positive Days", 15, 60, 38)
plan_def = st.sidebar.slider("Default Days", 15, 60, 30)

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if file:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    def num(col, default=0):
        return pd.to_numeric(df.get(col, default), errors="coerce").fillna(default)

    sales_cols = ['7 Days Sales','15 Days Sales','30 Days Sales','45 Days Sales','60 Days Sales']
    for c in sales_cols:
        df[c] = num(c)

    df['Box Qty'] = num('Box Qty')
    df['Current Stock'] = num('Current Stock')
    df['Hold & Unbilled Stock'] = num('Hold & Unbilled Stock')
    df['TOTAL_STOCK'] = df['Current Stock'] + df['Hold & Unbilled Stock']
    df['Rank'] = num('Top 500 SKU Rank', 9999)

    df['Review'] = df.get('Review', '').astype(str)
    df['MOS'] = df.get('MOS-WH Available', '').astype(str)

    # -----------------------------------------------------
    # EXACT COLAB AUTO_PO FUNCTION
    # -----------------------------------------------------
    def auto_po(row):

        S7,S15,S30,S45,S60 = row[sales_cols]
        Stock = row['TOTAL_STOCK']
        Box = row['Box Qty']
        Review = row['Review']
        Rank = row['Rank']
        MOS = row['MOS']

        if MOS != 'Yes' or Box <= 0:
            return 0

        IsTopHotcake = Review in ['Top-HotCake','Top-Hotcake','Hot Cake']
        IsPositive = Review == 'Positive'
        IsNewSKU = Review == 'New SKU'
        IsTop200 = Rank <= 200

        DailyIncl7 = (
            (S7/7)*w7 + (S15/15)*w15 + (S30/30)*w30 +
            (S45/45)*w45 + (S60/60)*w60
        )

        DailyExcl7 = (
            (S15/15)*ew15 + (S30/30)*ew30 +
            (S45/45)*ew45 + (S60/60)*ew60
        )

        FinalDaily = DailyExcl7 if (IsTop200 or IsTopHotcake or IsPositive) else DailyIncl7

        if IsTop200 or IsTopHotcake:
            PlanDays = plan_top
        elif IsPositive or IsNewSKU:
            PlanDays = plan_pos
        else:
            PlanDays = plan_def

        Target = FinalDaily * PlanDays
        Shortage = max(Target - Stock, 0)

        if Shortage <= 0:
            return 0

        Remainder = Shortage % Box
        QtyRaw = (
            np.ceil(Shortage / Box) * Box
            if Remainder >= threshold * Box
            else np.floor(Shortage / Box) * Box
        )

        MaxStock = FinalDaily * 60
        ForceMin = Shortage > 0 and (IsTopHotcake or IsPositive or IsNewSKU or IsTop200)

        FinalQty = Box if (Stock > MaxStock or QtyRaw == 0) and ForceMin else QtyRaw
        FinalQty = min(FinalQty, 10*Box, 120)

        if S7==S15==S30==S45==S60==0:
            return 0

        return int(FinalQty)

    # -----------------------------------------------------
    # AUTO-FILL MANUAL REQUIRED QTY (KEY FIX)
    # -----------------------------------------------------
    df['Manual Required Qty'] = df.apply(auto_po, axis=1)

    st.success("✅ PO Calculated (Manual Qty Auto-Filled)")

    # -----------------------------------------------------
    # EDITABLE TABLE (ONLY MANUAL COLUMN)
    # -----------------------------------------------------
    df = st.data_editor(
        df,
        disabled=[c for c in df.columns if c != 'Manual Required Qty'],
        use_container_width=True
    )

    # -----------------------------------------------------
    # DOWNLOAD
    # -----------------------------------------------------
    output = BytesIO()
    df.fillna("").to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        "⬇️ Download Final PO Excel",
        data=output,
        file_name="FINAL_PO_WITH_UI_CONFIG.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Upload Excel to start")

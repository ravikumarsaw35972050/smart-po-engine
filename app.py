# =========================================================
# SMART PO ENGINE – FINAL FIXED VERSION (NO NONE ISSUE)
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
    # AUTO PO LOGIC (ALWAYS RETURNS INT)
    # -----------------------------------------------------
    def auto_po(row):

        try:
            if row['MOS'] != 'Yes' or row['Box Qty'] <= 0:
                return 0

            S7,S15,S30,S45,S60 = row[sales_cols]
            stock = row['TOTAL_STOCK']
            box = row['Box Qty']
            review = row['Review']
            rank = row['Rank']

            is_top200 = rank <= 200
            is_hotcake = review in ['Top-HotCake','Top-Hotcake','Hot Cake']
            is_positive = review == 'Positive'
            is_newsku = review == 'New SKU'

            daily_incl7 = (
                (S7/7)*w7 + (S15/15)*w15 + (S30/30)*w30 +
                (S45/45)*w45 + (S60/60)*w60
            )

            daily_excl7 = (
                (S15/15)*ew15 + (S30/30)*ew30 +
                (S45/45)*ew45 + (S60/60)*ew60
            )

            final_daily = daily_excl7 if (is_top200 or is_hotcake or is_positive) else daily_incl7

            if is_top200 or is_hotcake:
                plan_days = plan_top
            elif is_positive or is_newsku:
                plan_days = plan_pos
            else:
                plan_days = plan_def

            target = final_daily * plan_days
            shortage = max(target - stock, 0)

            if shortage <= 0:
                return 0

            remainder = shortage % box
            qty_raw = (
                np.ceil(shortage / box) * box
                if remainder >= threshold * box
                else np.floor(shortage / box) * box
            )

            max_stock = final_daily * 60
            force_min = shortage > 0 and (is_top200 or is_hotcake or is_positive or is_newsku)

            final_qty = box if (stock > max_stock or qty_raw == 0) and force_min else qty_raw
            final_qty = min(final_qty, 10*box, 120)

            return int(max(final_qty, 0))

        except Exception:
            return 0   # 🔥 GUARANTEE NO NONE

    # -----------------------------------------------------
    # AUTO-FILL MANUAL REQUIRED QTY
    # -----------------------------------------------------
    df['Manual Required Qty'] = df.apply(auto_po, axis=1).fillna(0).astype(int)

    st.success("✅ Manual Required Qty auto-calculated")

    # -----------------------------------------------------
    # EDITABLE TABLE (OPTIONAL OVERRIDE)
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
        file_name="FINAL_PO_WITH_AUTO_MANUAL_QTY.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Upload Excel to start")

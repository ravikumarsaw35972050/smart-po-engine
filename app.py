# =========================================================
# SMART PURCHASE ORDER ENGINE – ALL IN ONE STREAMLIT APP
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np

# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="Smart PO Engine", layout="wide")
st.title("📦 Smart Purchase Order Engine")

# ---------------------------------------------------------
# SIDEBAR CONFIGURATION
# ---------------------------------------------------------
st.sidebar.header("⚙️ Configuration")

st.sidebar.subheader("Weighted Average (Include 7 Days)")
w7  = st.sidebar.slider("7 Days Weight", 0.0, 1.0, 0.35)
w15 = st.sidebar.slider("15 Days Weight", 0.0, 1.0, 0.25)
w30 = st.sidebar.slider("30 Days Weight", 0.0, 1.0, 0.20)
w45 = st.sidebar.slider("45 Days Weight", 0.0, 1.0, 0.12)
w60 = st.sidebar.slider("60 Days Weight", 0.0, 1.0, 0.08)

st.sidebar.subheader("Weighted Average (Exclude 7 Days)")
ew15 = st.sidebar.slider("15 Days (Exclude)", 0.0, 1.0, 0.40)
ew30 = st.sidebar.slider("30 Days (Exclude)", 0.0, 1.0, 0.30)
ew45 = st.sidebar.slider("45 Days (Exclude)", 0.0, 1.0, 0.20)
ew60 = st.sidebar.slider("60 Days (Exclude)", 0.0, 1.0, 0.10)

st.sidebar.subheader("Plan Days")
plan_top = st.sidebar.number_input("Top200 / Hotcake", 10, 90, 45)
plan_pos = st.sidebar.number_input("Positive / New SKU", 10, 90, 38)
plan_def = st.sidebar.number_input("Default", 10, 90, 30)

st.sidebar.subheader("Box Rounding Rule")
box_threshold = st.sidebar.slider("Rounding Threshold (80%)", 0.5, 1.0, 0.8)

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # -----------------------------------------------------
    # SAFE NUMERIC FUNCTION
    # -----------------------------------------------------
    def num(col, default=0):
        if col in df.columns:
            return pd.to_numeric(df[col], errors="coerce").fillna(default)
        return pd.Series([default]*len(df))

    sales_cols = [
        '7 Days Sales','15 Days Sales','30 Days Sales',
        '45 Days Sales','60 Days Sales'
    ]

    for c in sales_cols:
        df[c] = num(c)

    df['Box Qty'] = num('Box Qty')
    df['Current Stock'] = num('Current Stock')
    df['Hold & Unbilled Stock'] = num('Hold & Unbilled Stock')
    df['TOTAL_STOCK'] = df['Current Stock'] + df['Hold & Unbilled Stock']
    df['Rank'] = num('Top 500 SKU Rank', 9999)

    df['Review'] = df['Review'].astype(str).str.strip()
    df['MOS'] = df['MOS-WH Available'].astype(str).str.strip()

    # -----------------------------------------------------
    # AUTO PO LOGIC (EXACT BUSINESS RULE)
    # -----------------------------------------------------
    def auto_po(row):

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
        if remainder >= box_threshold * box:
            qty_raw = np.ceil(shortage / box) * box
        else:
            qty_raw = np.floor(shortage / box) * box

        max_stock = final_daily * 60
        is_overstock = stock > max_stock
        force_min = shortage > 0 and (is_top200 or is_hotcake or is_positive or is_newsku)

        final_qty = box if (is_overstock or qty_raw == 0) and force_min else qty_raw
        limit_qty = min(10*box, 120)

        result = np.floor(min(final_qty, limit_qty) / box) * box

        if S7==S15==S30==S45==S60==0:
            return 0

        return int(result)

    # -----------------------------------------------------
    # APPLY & OUTPUT
    # -----------------------------------------------------
    df['Manual Required Qty'] = df.apply(auto_po, axis=1)

    st.success("✅ Purchase Order Calculated Successfully")
    st.dataframe(df.head(20))

    st.download_button(
        "⬇️ Download Result Excel",
        data=df.to_excel(index=False),
        file_name="SMART_PO_RESULT.xlsx"
    )

# =========================================================
# END OF APP
# =========================================================

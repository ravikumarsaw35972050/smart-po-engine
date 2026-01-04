# =========================================================
# SMART PURCHASE ORDER ENGINE – STREAMLIT APP
# FULLY REVISED | STREAMLIT CLOUD SAFE | EXCEL EXPORT FIXED
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(
    page_title="Smart Purchase Order Engine",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Smart Purchase Order Engine")

# ---------------------------------------------------------
# SIDEBAR – CONFIGURATION (WEIGHTS)
# ---------------------------------------------------------
st.sidebar.header("⚙️ Configuration")

st.sidebar.subheader("Weighted Average (Include 7 Days)")
w7  = st.sidebar.slider("7 Days Weight",  0.0, 1.0, 0.35, 0.01)
w15 = st.sidebar.slider("15 Days Weight", 0.0, 1.0, 0.25, 0.01)
w30 = st.sidebar.slider("30 Days Weight", 0.0, 1.0, 0.20, 0.01)
w45 = st.sidebar.slider("45 Days Weight", 0.0, 1.0, 0.12, 0.01)
w60 = st.sidebar.slider("60 Days Weight", 0.0, 1.0, 0.08, 0.01)

st.sidebar.markdown("---")
st.sidebar.info("💡 Adjust weights to control sales trend sensitivity")

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
uploaded_file = st.file_uploader(
    "📤 Upload Excel File",
    type=["xlsx"]
)

if uploaded_file is None:
    st.stop()

# ---------------------------------------------------------
# READ EXCEL
# ---------------------------------------------------------
df = pd.read_excel(uploaded_file)
df.columns = df.columns.str.strip()

# ---------------------------------------------------------
# SAFE NUMERIC FUNCTION
# ---------------------------------------------------------
def num(series, default=0):
    return pd.to_numeric(series, errors="coerce").fillna(default)

# ---------------------------------------------------------
# REQUIRED COLUMNS (SAFE LOAD)
# ---------------------------------------------------------
sales_cols = [
    "7 Days Sales",
    "15 Days Sales",
    "30 Days Sales",
    "45 Days Sales",
    "60 Days Sales"
]

for c in sales_cols:
    if c not in df.columns:
        df[c] = 0
    df[c] = num(df[c])

df["Box Qty"] = num(df.get("Box Qty", 0))
df["Current Stock"] = num(df.get("Current Stock", 0))
df["Hold & Unbilled Stock"] = num(df.get("Hold & Unbilled Stock", 0))
df["TOTAL_STOCK"] = df["Current Stock"] + df["Hold & Unbilled Stock"]

df["Rank"] = num(df.get("Top 500 SKU Rank", 9999), 9999)
df["Review"] = df.get("Review", "").astype(str).str.strip()
df["MOS"] = df.get("MOS-WH Available", "").astype(str).str.strip()

# ---------------------------------------------------------
# AUTO PO LOGIC (WEIGHTED + BUSINESS RULES)
# ---------------------------------------------------------
def calculate_po(row):
    S7, S15, S30, S45, S60 = row[sales_cols]
    stock = row["TOTAL_STOCK"]
    box = row["Box Qty"]
    review = row["Review"]
    rank = row["Rank"]
    mos = row["MOS"]

    if mos != "Yes" or box <= 0:
        return 0

    # Flags
    is_hotcake = review in ["Top-HotCake", "Top-Hotcake", "Hot Cake"]
    is_positive = review == "Positive"
    is_new = review == "New SKU"
    is_top200 = rank <= 200

    # Weighted Daily Sales
    daily_incl7 = (
        (S7 / 7)  * w7  +
        (S15 / 15) * w15 +
        (S30 / 30) * w30 +
        (S45 / 45) * w45 +
        (S60 / 60) * w60
    )

    daily_excl7 = max(
        S15 / 15 if S15 else 0,
        S30 / 30 if S30 else 0,
        S45 / 45 if S45 else 0,
        S60 / 60 if S60 else 0
    )

    final_daily = daily_excl7 if (is_top200 or is_hotcake or is_positive) else daily_incl7

    # Plan Days
    if is_top200 or is_hotcake:
        plan_days = 45
    elif is_positive or is_new:
        plan_days = 38
    else:
        plan_days = 30

    target = final_daily * plan_days
    shortage = max(target - stock, 0)

    if shortage <= 0:
        return 0

    # Box rounding (80% rule)
    remainder = shortage % box
    qty_raw = (
        np.ceil(shortage / box) * box
        if remainder >= 0.8 * box
        else np.floor(shortage / box) * box
    )

    max_stock = final_daily * 60
    is_overstock = stock > max_stock
    force_min_box = shortage > 0 and (is_hotcake or is_positive or is_new or is_top200)

    final_qty = box if (is_overstock or qty_raw == 0) and force_min_box else qty_raw

    limit_qty = min(10 * box, 120)
    result = np.floor(min(final_qty, limit_qty) / box) * box

    if S7 == S15 == S30 == S45 == S60 == 0:
        return 0

    return int(result)

# ---------------------------------------------------------
# APPLY LOGIC
# ---------------------------------------------------------
df["Manual Required Qty"] = df.apply(calculate_po, axis=1)

st.success("✅ Purchase Order Calculated Successfully")

# ---------------------------------------------------------
# SHOW DATA
# ---------------------------------------------------------
st.dataframe(df, use_container_width=True)

# ---------------------------------------------------------
# EXCEL DOWNLOAD (STREAMLIT CLOUD SAFE)
# ---------------------------------------------------------
output = BytesIO()
df.to_excel(output, index=False, engine="openpyxl")
output.seek(0)

st.download_button(
    label="📥 Download PO Excel",
    data=output,
    file_name="Smart_PO_Output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

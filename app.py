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
# SIDEBAR – CONFIGURATION
# ---------------------------------------------------------
st.sidebar.header("⚙️ Configuration")

st.sidebar.subheader("Weighted Average (Include 7 Days)")
w7  = st.sidebar.slider("7 Days Weight",  0.0, 1.0, 0.35, 0.01)
w15 = st.sidebar.slider("15 Days Weight", 0.0, 1.0, 0.25, 0.01)
w30 = st.sidebar.slider("30 Days Weight", 0.0, 1.0, 0.20, 0.01)
w45 = st.sidebar.slider("45 Days Weight", 0.0, 1.0, 0.12, 0.01)
w60 = st.sidebar.slider("60 Days Weight", 0.0, 1.0, 0.08, 0.01)

st.sidebar.markdown("---")
st.sidebar.subheader("Weighted Average (Exclude 7 Days)")
ew15 = st.sidebar.slider("15 Days Weight (Ex)", 0.0, 1.0, 0.40, 0.01)
ew30 = st.sidebar.slider("30 Days Weight (Ex)", 0.0, 1.0, 0.30, 0.01)
ew45 = st.sidebar.slider("45 Days Weight (Ex)", 0.0, 1.0, 0.20, 0.01)
ew60 = st.sidebar.slider("60 Days Weight (Ex)", 0.0, 1.0, 0.10, 0.01)

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
uploaded_file = st.file_uploader(
    "📤 Upload Excel File",
    type=["xlsx"],
    help="Upload Order Sheet Excel (.xlsx)"
)

# ---------------------------------------------------------
# HELPER FUNCTION
# ---------------------------------------------------------
def num(series, default=0):
    return pd.to_numeric(series, errors="coerce").fillna(default)

# ---------------------------------------------------------
# CORE LOGIC FUNCTION
# ---------------------------------------------------------
def calculate_po(df):

    # Required columns mapping
    sales_cols = {
        "7 Days Sales": 7,
        "15 Days Sales": 15,
        "30 Days Sales": 30,
        "45 Days Sales": 45,
        "60 Days Sales": 60,
    }

    # Safe numeric conversions
    for col in sales_cols:
        if col not in df.columns:
            df[col] = 0
        df[col] = num(df[col])

    df["Box Qty"] = num(df.get("Box Qty", 0))
    df["Current Stock"] = num(df.get("Current Stock", 0))
    df["Hold & Unbilled Stock"] = num(df.get("Hold & Unbilled Stock", 0))
    df["TOTAL_STOCK"] = df["Current Stock"] + df["Hold & Unbilled Stock"]

    df["Rank"] = num(df.get("Top 500 SKU Rank", 9999), 9999)
    df["Review"] = df.get("Review", "").astype(str).str.strip()
    df["MOS"] = df.get("MOS-WH Available", "").astype(str).str.strip()

    # -----------------------------------------------------
    # AUTO PO ROW FUNCTION
    # -----------------------------------------------------
    def auto_po(row):

        S7, S15, S30, S45, S60 = (
            row["7 Days Sales"],
            row["15 Days Sales"],
            row["30 Days Sales"],
            row["45 Days Sales"],
            row["60 Days Sales"],
        )

        Box = row["Box Qty"]
        Stock = row["TOTAL_STOCK"]
        Review = row["Review"]
        Rank = row["Rank"]
        MOS = row["MOS"]

        if MOS != "Yes" or Box <= 0:
            return 0

        IsTopHotcake = Review in ["Top-HotCake", "Top-Hotcake", "Hot Cake"]
        IsPositive = Review == "Positive"
        IsNewSKU = Review == "New SKU"
        IsTop200 = Rank <= 200

        DailyIncl7 = (
            (S7 / 7) * w7 +
            (S15 / 15) * w15 +
            (S30 / 30) * w30 +
            (S45 / 45) * w45 +
            (S60 / 60) * w60
        )

        DailyExcl7 = (
            (S15 / 15) * ew15 +
            (S30 / 30) * ew30 +
            (S45 / 45) * ew45 +
            (S60 / 60) * ew60
        )

        FinalDaily = DailyExcl7 if (IsTop200 or IsTopHotcake or IsPositive) else DailyIncl7

        if IsTop200 or IsTopHotcake:
            PlanDays = 45
        elif IsPositive or IsNewSKU:
            PlanDays = 38
        else:
            PlanDays = 30

        Target = FinalDaily * PlanDays
        Shortage = max(Target - Stock, 0)

        if Shortage <= 0:
            return 0

        Remainder = Shortage % Box
        QtyRaw = (
            np.ceil(Shortage / Box) * Box
            if Remainder >= (0.8 * Box)
            else np.floor(Shortage / Box) * Box
        )

        MaxStock = FinalDaily * 60
        IsOverstock = Stock > MaxStock
        ForceMinBox = Shortage > 0 and (IsTopHotcake or IsPositive or IsNewSKU or IsTop200)

        FinalQty = Box if (IsOverstock and ForceMinBox) else QtyRaw
        FinalQty = min(FinalQty, min(10 * Box, 120))

        if S7 == S15 == S30 == S45 == S60 == 0:
            return 0

        return int(max(FinalQty, 0))

    df["Manual Required Qty"] = df.apply(auto_po, axis=1)
    return df

# ---------------------------------------------------------
# PROCESS FILE
# ---------------------------------------------------------
if uploaded_file:

    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded successfully")

        result_df = calculate_po(df)

        st.success("✅ Purchase Order Calculated Successfully")
        st.dataframe(result_df, use_container_width=True)

        # -------------------------------------------------
        # EXCEL DOWNLOAD (STREAMLIT CLOUD SAFE)
        # -------------------------------------------------
        output = BytesIO()
        result_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            label="📥 Download PO Excel",
            data=output,
            file_name="Smart_PO_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error while processing file")
        st.exception(e)

# =========================================================
# END OF APP
# =========================================================

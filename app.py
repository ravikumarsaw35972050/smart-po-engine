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

# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
st.subheader("📤 Upload Excel File")

uploaded_file = st.file_uploader(
    "Drag and drop file here",
    type=["xlsx"]
)

if uploaded_file:

    try:
        # -------------------------------------------------
        # READ DATA
        # -------------------------------------------------
        df = pd.read_excel(uploaded_file)

        # -------------------------------------------------
        # SAFE COLUMN HANDLING
        # -------------------------------------------------
        sales_cols = {
            "7 Days Sales": "7_days",
            "15 Days Sales": "15_days",
            "30 Days Sales": "30_days",
            "45 Days Sales": "45_days",
            "60 Days Sales": "60_days",
            "Current Stock": "current_stock",
            "Hold & Unbilled Stock": "hold_stock"
        }

        for col in sales_cols.keys():
            if col not in df.columns:
                df[col] = 0

        df[list(sales_cols.keys())] = df[list(sales_cols.keys())].fillna(0)

        # -------------------------------------------------
        # TOTAL STOCK
        # -------------------------------------------------
        df["TOTAL_STOCK"] = df["Current Stock"] + df["Hold & Unbilled Stock"]

        # -------------------------------------------------
        # WEIGHTED AVG SALES
        # -------------------------------------------------
        df["WEIGHTED_SALES"] = (
            df["7 Days Sales"]  * w7  +
            df["15 Days Sales"] * w15 +
            df["30 Days Sales"] * w30 +
            df["45 Days Sales"] * w45 +
            df["60 Days Sales"] * w60
        )

        # -------------------------------------------------
        # MOS (MONTHS OF STOCK)
        # -------------------------------------------------
        df["MOS"] = np.where(
            df["WEIGHTED_SALES"] > 0,
            (df["TOTAL_STOCK"] / df["WEIGHTED_SALES"]).round(2),
            9999
        )

        # -------------------------------------------------
        # RANK (LOW MOS = HIGH PRIORITY)
        # -------------------------------------------------
        df["RANK"] = df["MOS"].rank(method="dense", ascending=True).astype(int)

        st.success("✅ Purchase Order Calculated Successfully")

        # -------------------------------------------------
        # DISPLAY DATA (HIDE INTERNAL COLUMNS)
        # -------------------------------------------------
        hide_cols = ["TOTAL_STOCK", "MOS", "RANK", "WEIGHTED_SALES"]
        display_cols = [c for c in df.columns if c not in hide_cols]

        st.dataframe(
            df[display_cols],
            use_container_width=True
        )

        # -------------------------------------------------
        # EXCEL DOWNLOAD (STREAMLIT CLOUD SAFE)
        # -------------------------------------------------
        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
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

else:
    st.info("⬆️ Please upload an Excel file to continue")

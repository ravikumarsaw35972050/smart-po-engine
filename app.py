# ============================================================
# SMART PURCHASE ORDER ENGINE – STREAMLIT APP
# FINAL CLEAN VERSION | STREAMLIT CLOUD SAFE
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Smart Purchase Order Engine",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Smart Purchase Order Engine")

# ------------------------------------------------------------
# SIDEBAR – CONFIGURATION
# ------------------------------------------------------------
st.sidebar.header("⚙️ Configuration")

st.sidebar.subheader("Weighted Average (Include 7 Days)")
w7  = st.sidebar.slider("7 Days Weight",  0.0, 1.0, 0.35, 0.01)
w15 = st.sidebar.slider("15 Days Weight", 0.0, 1.0, 0.25, 0.01)
w30 = st.sidebar.slider("30 Days Weight", 0.0, 1.0, 0.20, 0.01)
w45 = st.sidebar.slider("45 Days Weight", 0.0, 1.0, 0.12, 0.01)
w60 = st.sidebar.slider("60 Days Weight", 0.0, 1.0, 0.08, 0.01)

# ------------------------------------------------------------
# FILE UPLOAD
# ------------------------------------------------------------
uploaded_file = st.file_uploader(
    "📤 Upload Excel File",
    type=["xlsx"],
    help="Upload your sales & stock Excel file"
)

if uploaded_file is not None:
    try:
        # ----------------------------------------------------
        # READ EXCEL
        # ----------------------------------------------------
        df = pd.read_excel(uploaded_file)

        # ----------------------------------------------------
        # REQUIRED COLUMNS CHECK
        # ----------------------------------------------------
        required_cols = [
            "7 Days Sales", "15 Days Sales",
            "30 Days Sales", "45 Days Sales",
            "60 Days Sales", "Current Stock"
        ]

        for col in required_cols:
            if col not in df.columns:
                st.error(f"❌ Missing required column: {col}")
                st.stop()

        # ----------------------------------------------------
        # CLEAN DATA
        # ----------------------------------------------------
        df[required_cols] = df[required_cols].fillna(0)

        # ----------------------------------------------------
        # WEIGHTED AVERAGE SALES
        # ----------------------------------------------------
        df["Weighted Avg Sales"] = (
            df["7 Days Sales"]  * w7  +
            df["15 Days Sales"] * w15 +
            df["30 Days Sales"] * w30 +
            df["45 Days Sales"] * w45 +
            df["60 Days Sales"] * w60
        )

        # ----------------------------------------------------
        # SYSTEM SUGGESTED ORDER
        # ----------------------------------------------------
        df["Suggested Order Qty"] = np.maximum(
            np.round(df["Weighted Avg Sales"] * 30 - df["Current Stock"]),
            0
        ).astype(int)

        # ----------------------------------------------------
        # MANUAL REQUIRED QTY (USER INPUT)
        # ----------------------------------------------------
        if "Manual Required Qty" not in df.columns:
            df["Manual Required Qty"] = 0

        # ----------------------------------------------------
        # DISPLAY EDITABLE TABLE
        # ----------------------------------------------------
        st.success("✅ Purchase Order Calculated Successfully")

        editable_cols = ["Manual Required Qty"]

        result_df = st.data_editor(
            df,
            use_container_width=True,
            disabled=[c for c in df.columns if c not in editable_cols],
            num_rows="fixed"
        )

        # ----------------------------------------------------
        # FINAL ORDER QTY (MANUAL OVERRIDE)
        # ----------------------------------------------------
        result_df["Final Order Qty"] = np.where(
            result_df["Manual Required Qty"] > 0,
            result_df["Manual Required Qty"],
            result_df["Suggested Order Qty"]
        ).astype(int)

        # ----------------------------------------------------
        # SANITIZE DATA BEFORE EXPORT
        # ----------------------------------------------------
        safe_df = result_df.copy()
        safe_df = safe_df.fillna("")

        for col in safe_df.columns:
            if safe_df[col].apply(lambda x: isinstance(x, (bytes, bytearray))).any():
                safe_df.drop(columns=[col], inplace=True)

        # ----------------------------------------------------
        # EXCEL DOWNLOAD (STREAMLIT CLOUD SAFE)
        # ----------------------------------------------------
        output = BytesIO()
        safe_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            label="⬇️ Download Final PO Excel",
            data=output,
            file_name="Smart_PO_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error while processing file")
        st.text(f"Error Type: {type(e).__name__}")
        st.text(f"Error Message: {str(e)}")

else:
    st.info("👆 Please upload an Excel file to start")

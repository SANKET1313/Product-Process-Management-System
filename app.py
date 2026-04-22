import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import plotly.express as px
import plotly.io as pio
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Product Process Tracker", layout="wide")
st.title("📊 Product Process Management System")

# -------------------------------
# SESSION STATE
# -------------------------------
if "df" not in st.session_state:
    st.session_state.df = None

if "prev_df" not in st.session_state:
    st.session_state.prev_df = None

if "file_loaded" not in st.session_state:
    st.session_state.file_loaded = False

# -------------------------------
# PROCESS STEPS
# -------------------------------
process_steps = [
    "CAD", "Casting", "Filing", "Setting",
    "Polishing", "QC", "Gold Check", "Export"
]

# -------------------------------
# TABS
# -------------------------------
tab1, tab2, tab3 = st.tabs([
    "📥 Import File",
    "✏️ Update Data",
    "📊 Dashboard & Export"
])

# =========================================================
# 📥 TAB 1: IMPORT FILE
# =========================================================
with tab1:

    uploaded_file = st.file_uploader("Upload CSV/Excel", type=["xlsx", "csv"])

    if st.button("🚀 Import File"):

        if uploaded_file is None:
            st.error("❌ Select file first")

        else:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)

            # Required columns
            df["Status"] = df.get("Status", "")
            df["Status Timestamp"] = df.get("Status Timestamp", "")
            df["Process"] = df.get("Process", "")
            df["Diamond"] = df.get("Diamond", "")
            df["Entity Name"] = df.get("Entity Name", "")

            # Process timestamps
            for step in process_steps:
                col = f"{step} Timestamp"
                if col not in df.columns:
                    df[col] = ""

            st.session_state.df = df.copy()
            st.session_state.prev_df = df.copy()
            st.session_state.file_loaded = True

            st.success("✅ File Imported")

    if st.session_state.file_loaded:
        st.dataframe(st.session_state.df, use_container_width=True)

# =========================================================
# ✏️ TAB 2: UPDATE DATA (NO DELAY)
# =========================================================
with tab2:

    if not st.session_state.file_loaded:
        st.warning("⚠️ Import file first")

    else:
        df = st.session_state.df

        st.subheader("✏️ Update Data")

        # -------------------------------
        # COLUMN MANAGEMENT
        # -------------------------------
        colA, colB = st.columns(2)

        with colA:
            new_col = st.text_input("New Column Name")
            default_val = st.text_input("Default Value")

            if st.button("➕ Add Column"):
                if new_col and new_col not in df.columns:
                    df[new_col] = default_val
                    st.session_state.df = df
                    st.session_state.prev_df = df.copy()
                    st.rerun()

        with colB:
            col_del = st.selectbox("Delete Column", df.columns)

            if st.button("🗑️ Delete Column"):
                df = df.drop(columns=[col_del])
                st.session_state.df = df
                st.session_state.prev_df = df.copy()
                st.rerun()

        st.divider()

        # -------------------------------
        # DATA EDITOR
        # -------------------------------
        edited_df = st.data_editor(
            st.session_state.df,
            num_rows="dynamic",
            use_container_width=True,
            key="editor",
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=[
                        "Draft", "Active", "Inactive",
                        "Out of Stock", "Archived",
                        "Pending Approval", "Pre-Order"
                    ]
                ),
                "Process": st.column_config.SelectboxColumn(
                    "Process",
                    options=process_steps
                ),
                "Diamond": st.column_config.SelectboxColumn(
                    "Diamond",
                    options=["Given", "Not Given"]
                ),
                "Entity Name": st.column_config.TextColumn("Entity Name")
            }
        )

        # -------------------------------
        # 🔥 TIMESTAMP LOGIC (NO DELAY)
        # -------------------------------
        prev_df = st.session_state.prev_df

        for i in range(len(edited_df)):

            # STATUS TIMESTAMP
            old_status = prev_df.loc[i, "Status"] if i in prev_df.index else ""
            new_status = edited_df.loc[i, "Status"]

            if old_status != new_status and new_status != "":
                edited_df.loc[i, "Status Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # PROCESS TIMESTAMP
            old_process = prev_df.loc[i, "Process"] if i in prev_df.index else ""
            new_process = edited_df.loc[i, "Process"]

            if old_process != new_process and new_process != "":
                col = f"{new_process} Timestamp"
                edited_df.loc[i, col] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # SAVE STATE
        st.session_state.df = edited_df
        st.session_state.prev_df = edited_df.copy()

        st.success("✅ Updated Instantly (No Delay)")

# =========================================================
# 📊 TAB 3: DASHBOARD
# =========================================================
with tab3:

    df = st.session_state.get("df")

    if df is None or df.empty:
        st.warning("⚠️ No data")

    else:
        status_count = df["Status"].value_counts().reset_index()
        status_count.columns = ["Status", "Count"]

        col1, col2 = st.columns(2)

        # BAR CHART
        with col1:
            fig = px.bar(status_count, x="Status", y="Count",
                         color_discrete_sequence=["#79A9D1"])
            fig.update_layout(template="plotly_dark")
            st.plotly_chart(fig, use_container_width=True)

        # PIE CHART
        with col2:
            fig2 = px.pie(status_count, names="Status", values="Count")
            fig2.update_layout(template="plotly_dark")
            st.plotly_chart(fig2, use_container_width=True)

        # EXPORT
        try:
            img1 = pio.to_image(fig, format="png", scale=2)
            img2 = pio.to_image(fig2, format="png", scale=2)

            def create_excel():
                wb = Workbook()
                ws = wb.active
                ws.add_image(XLImage(BytesIO(img1)), "A1")
                ws.add_image(XLImage(BytesIO(img2)), "A22")
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                return buffer

            st.download_button(
                "⬇️ Download Charts Excel",
                data=create_excel(),
                file_name="charts.xlsx"
            )

        except:
            st.warning("Install kaleido → pip install kaleido")
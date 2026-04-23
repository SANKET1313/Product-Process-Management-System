"""
Product Process Management System
===================================
Install:
    pip install streamlit pandas plotly openpyxl kaleido

Run:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import plotly.express as px
import plotly.io as pio
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# ──────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Product Process Tracker", layout="wide")
st.title("📊 Product Process Management System")

# ──────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────
PROCESS_STEPS = [
    "CAD", "Casting", "Filing", "Setting",
    "Polishing", "QC", "Gold Check", "Export"
]

STATUS_OPTIONS = [
    "Draft", "Active", "Inactive",
    "Out of Stock", "Archived",
    "Pending Approval", "Pre-Order"
]

DIAMOND_OPTIONS = ["Given", "Not Given"]

REQUIRED_COLS  = ["Entity Name", "Status", "Status Timestamp",
                   "Process", "Diamond"]
TIMESTAMP_COLS = ["Status Timestamp"] + [f"{s} Timestamp" for s in PROCESS_STEPS]

# ──────────────────────────────────────────────────────────────
# SESSION STATE — initialise once
# ──────────────────────────────────────────────────────────────
for _k, _v in {
    "df":          None,
    "prev_df":     None,
    "file_loaded": False,
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def safe(val: object) -> str:
    """Return a clean string; NaN / None → ''."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val).strip()


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """Add every required column that is missing; fill NaN → ''."""
    df = df.copy()
    for col in REQUIRED_COLS + TIMESTAMP_COLS:
        if col not in df.columns:
            df[col] = ""
    return df.fillna("").reset_index(drop=True)


def stamp_changes(new_df: pd.DataFrame, old_df: pd.DataFrame) -> pd.DataFrame:
    """
    Walk through every row.  Where Status or Process changed,
    write the current timestamp into the matching column.
    """
    now    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result = new_df.copy()

    for i in result.index:
        if i not in old_df.index:
            continue

        # Status
        if safe(result.loc[i, "Status"]) != safe(old_df.loc[i, "Status"]) \
                and safe(result.loc[i, "Status"]):
            result.loc[i, "Status Timestamp"] = now

        # Process
        new_proc = safe(result.loc[i, "Process"])
        if new_proc and new_proc != safe(old_df.loc[i, "Process"]):
            ts_col = f"{new_proc} Timestamp"
            if ts_col in result.columns:
                result.loc[i, ts_col] = now

    return result


# ──────────────────────────────────────────────────────────────
# on_change CALLBACK
# Fires synchronously BEFORE Streamlit re-renders the page.
# At this moment st.session_state["_editor"] holds the diff
# (edited_rows / added_rows / deleted_rows) and prev_df is
# still the snapshot from the previous render — perfect for
# detecting what actually changed.
# ──────────────────────────────────────────────────────────────
def _on_change():
    diff       = st.session_state.get("_editor", {})
    base       = st.session_state.df.copy()      # last saved state
    prev       = st.session_state.prev_df.copy() # snapshot before this edit

    edited_rows  = diff.get("edited_rows",  {})
    added_rows   = diff.get("added_rows",   [])
    deleted_rows = diff.get("deleted_rows", [])

    # 1️⃣  Apply cell edits
    for idx_str, changes in edited_rows.items():
        idx = int(idx_str)
        for col, val in changes.items():
            base.loc[idx, col] = val if val is not None else ""

    # 2️⃣  Apply new rows
    if added_rows:
        extras = pd.DataFrame(added_rows)
        for col in base.columns:
            if col not in extras.columns:
                extras[col] = ""
        extras = extras.reindex(columns=base.columns, fill_value="")
        base   = pd.concat([base, extras], ignore_index=True)

    # 3️⃣  Apply deletions
    if deleted_rows:
        base = base.drop(index=[r for r in deleted_rows if r in base.index])
        prev = prev.drop(index=[r for r in deleted_rows if r in prev.index])
        base = base.reset_index(drop=True)
        prev = prev.reset_index(drop=True)

    # 4️⃣  Stamp timestamps where dropdowns changed
    base = stamp_changes(base, prev)

    # 5️⃣  Persist — prev_df updated AFTER diff so next edit sees correct old values
    st.session_state.df      = base.copy()
    st.session_state.prev_df = base.copy()


# ──────────────────────────────────────────────────────────────
# COLUMN CONFIG for st.data_editor
# ──────────────────────────────────────────────────────────────
def build_col_config(df: pd.DataFrame) -> dict:
    cfg = {
        "Entity Name": st.column_config.TextColumn(
            "Entity Name", width="medium"),

        "Status": st.column_config.SelectboxColumn(
            "Status ▼",
            options=STATUS_OPTIONS,
            required=False,
            width="medium"),

        "Status Timestamp": st.column_config.TextColumn(
            "Status TS 🕐",
            disabled=True,
            width="medium"),

        "Process": st.column_config.SelectboxColumn(
            "Process ▼",
            options=PROCESS_STEPS,
            required=False,
            width="medium"),

        "Diamond": st.column_config.SelectboxColumn(
            "Diamond ▼",
            options=DIAMOND_OPTIONS,
            required=False,
            width="small"),
    }

    for step in PROCESS_STEPS:
        col = f"{step} Timestamp"
        if col in df.columns:
            cfg[col] = st.column_config.TextColumn(
                f"{step} TS 🕐",
                disabled=True,
                width="medium")

    return cfg


# ──────────────────────────────────────────────────────────────
# TABS
# ──────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "📥 Import File",
    "✏️ Update Data",
    "📊 Dashboard & Export",
])


# ══════════════════════════════════════════════════════════════
# TAB 1 — IMPORT
# ══════════════════════════════════════════════════════════════
with tab1:
    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

    if st.button("🚀 Import File", type="primary"):
        if uploaded is None:
            st.error("❌ Choose a file first.")
        else:
            try:
                raw = (pd.read_csv(uploaded)
                       if uploaded.name.endswith(".csv")
                       else pd.read_excel(uploaded))
                df  = prepare(raw)
                st.session_state.df          = df.copy()
                st.session_state.prev_df     = df.copy()
                st.session_state.file_loaded = True
                # Clear the editor widget so it reloads fresh data
                if "_editor" in st.session_state:
                    del st.session_state["_editor"]
                st.success(f"✅ Loaded {len(df)} rows, {len(df.columns)} columns.")
            except Exception as exc:
                st.error(f"❌ Could not read file: {exc}")

    if st.session_state.file_loaded:
        st.subheader("Preview")
        st.dataframe(st.session_state.df, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# TAB 2 — EDIT
# ══════════════════════════════════════════════════════════════
with tab2:

    if not st.session_state.file_loaded:
        st.warning("⚠️  Import a file in Tab 1 first.")

    else:
        # ── Column management ────────────────────────────────
        with st.expander("🗂  Add / Delete Columns", expanded=False):
            c1, c2 = st.columns(2)

            with c1:
                st.markdown("**Add a column**")
                new_col = st.text_input("Column name",    key="t2_new_col")
                def_val = st.text_input("Default value",  key="t2_def_val")
                if st.button("➕ Add"):
                    nc = new_col.strip()
                    if not nc:
                        st.warning("Enter a column name.")
                    elif nc in st.session_state.df.columns:
                        st.warning(f"'{nc}' already exists.")
                    else:
                        st.session_state.df[nc]      = def_val
                        st.session_state.prev_df[nc] = def_val
                        if "_editor" in st.session_state:
                            del st.session_state["_editor"]
                        st.rerun()

            with c2:
                st.markdown("**Delete a column**")
                protected = set(TIMESTAMP_COLS)
                deletable = [c for c in st.session_state.df.columns
                             if c not in protected]
                col_del   = st.selectbox("Column to delete", deletable, key="t2_del")
                if st.button("🗑️  Delete"):
                    st.session_state.df      = st.session_state.df.drop(columns=[col_del])
                    st.session_state.prev_df = st.session_state.prev_df.drop(columns=[col_del])
                    if "_editor" in st.session_state:
                        del st.session_state["_editor"]
                    st.rerun()

        st.divider()

        # ── Instructions ─────────────────────────────────────
        st.caption(
            "✏️  Click a **Status** or **Process** cell → choose from the dropdown.  "
            "The matching timestamp fills automatically.  "
            "All other cells are freely editable.  "
            "Use the ＋ button at the bottom of the table to add rows."
        )

        # ── Data editor ──────────────────────────────────────
        # Key insight:
        #   • key="_editor"  → Streamlit stores the diff (edited/added/deleted rows)
        #                       in st.session_state["_editor"] automatically.
        #   • on_change      → our callback reads that diff BEFORE the page rerenders,
        #                       so prev_df is still the previous snapshot → accurate diff.
        #   • We pass st.session_state.df as the base data.  The editor shows the
        #                       latest saved state on every render.
        # ─────────────────────────────────────────────────────
        st.data_editor(
            st.session_state.df,
            column_config=build_col_config(st.session_state.df),
            num_rows="dynamic",
            use_container_width=True,
            key="_editor",
            on_change=_on_change,
            height=500,
        )

        # ── Timestamp review table ───────────────────────────
        st.divider()
        st.subheader("🕐 Timestamps")
        show_cols = [c for c in
                     ["Entity Name", "Status", "Status Timestamp", "Process"]
                     + [f"{s} Timestamp" for s in PROCESS_STEPS]
                     if c in st.session_state.df.columns]
        st.dataframe(st.session_state.df[show_cols], use_container_width=True)

        # ── Quick CSV download ───────────────────────────────
        st.divider()
        st.download_button(
            "⬇️  Download current data (CSV)",
            data=st.session_state.df.to_csv(index=False).encode(),
            file_name="updated_data.csv",
            mime="text/csv",
        )


# ══════════════════════════════════════════════════════════════
# TAB 3 — DASHBOARD & EXPORT
# ══════════════════════════════════════════════════════════════
with tab3:

    _df = st.session_state.df

    if _df is None or _df.empty:
        st.warning("⚠️  No data — import a file first.")

    else:
        # ── Status charts ────────────────────────────────────
        _status_df = _df[_df["Status"].str.strip() != ""]

        fig_bar = fig_pie = None          # so export block can check safely

        if _status_df.empty:
            st.info("Assign statuses in Tab 2 to see charts.")
        else:
            s_cnt = _status_df["Status"].value_counts().reset_index()
            s_cnt.columns = ["Status", "Count"]

            c1, c2 = st.columns(2)
            with c1:
                fig_bar = px.bar(s_cnt, x="Status", y="Count",
                                 color_discrete_sequence=["#79A9D1"],
                                 title="Status Distribution")
                fig_bar.update_layout(template="plotly_dark")
                st.plotly_chart(fig_bar, use_container_width=True)

            with c2:
                fig_pie = px.pie(s_cnt, names="Status", values="Count",
                                 title="Status Share")
                fig_pie.update_layout(template="plotly_dark")
                st.plotly_chart(fig_pie, use_container_width=True)

        # ── Process chart ────────────────────────────────────
        st.subheader("⚙️  Process Stage Overview")
        _proc_df = _df[_df["Process"].str.strip() != ""]

        fig_proc = None
        if _proc_df.empty:
            st.info("Assign processes in Tab 2 to see this chart.")
        else:
            p_cnt = _proc_df["Process"].value_counts().reset_index()
            p_cnt.columns = ["Process", "Count"]
            fig_proc = px.bar(p_cnt, x="Process", y="Count",
                              color_discrete_sequence=["#F4A261"],
                              title="Items per Process Stage")
            fig_proc.update_layout(template="plotly_dark")
            st.plotly_chart(fig_proc, use_container_width=True)

        # ── Exports ──────────────────────────────────────────
        st.divider()
        st.subheader("⬇️  Export")

        # CSV — always available
        st.download_button(
            "⬇️  Download Data (CSV)",
            data=_df.to_csv(index=False).encode(),
            file_name="product_tracker.csv",
            mime="text/csv",
        )

        # Excel with embedded charts — needs kaleido
        if fig_bar and fig_pie:
            try:
                _img1 = pio.to_image(fig_bar, format="png", scale=2)
                _img2 = pio.to_image(fig_pie, format="png", scale=2)

                def _make_excel(_df=_df, _img1=_img1, _img2=_img2):
                    wb  = Workbook()
                    wsc = wb.active
                    wsc.title = "Charts"
                    wsc.add_image(XLImage(BytesIO(_img1)), "A1")
                    wsc.add_image(XLImage(BytesIO(_img2)), "A22")
                    wsd = wb.create_sheet("Data")
                    wsd.append(list(_df.columns))
                    for _row in _df.itertuples(index=False):
                        wsd.append(list(_row))
                    _buf = BytesIO()
                    wb.save(_buf)
                    _buf.seek(0)
                    return _buf

                st.download_button(
                    "⬇️  Download Charts + Data (Excel)",
                    data=_make_excel(),
                    file_name="product_tracker.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception:
                st.info("💡 Chart export needs kaleido →  `pip install kaleido`")
        else:
            st.info("Add status data to enable Excel export with charts.")
"""
Home Depot Planogram Generator — Streamlit App
"""
from __future__ import annotations

import traceback

import pandas as pd
import streamlit as st

# ─── Page config (must be first Streamlit call) ───────────────────────────────
st.set_page_config(
    page_title="HD Planogram Generator",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
/* ── Globals ──────────────────────────────────────────────────────────── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0f0f0f 0%, #1a1a1a 50%, #111 100%);
    color: #f0f0f0;
}
[data-testid="stHeader"]          { background: transparent; }
[data-testid="stSidebar"]         { background: #1a1a1a; }

/* ── Hero banner ──────────────────────────────────────────────────────── */
.hero-banner {
    background: linear-gradient(135deg, #F96302 0%, #cc4f00 60%, #1a1a1a 100%);
    border-radius: 16px;
    padding: 2.5rem 2.5rem 2rem;
    margin-bottom: 2rem;
}
.hero-title {
    font-size: 2.6rem;
    font-weight: 800;
    color: #ffffff;
    margin: 0 0 0.4rem 0;
    line-height: 1.1;
}
.hero-sub {
    font-size: 1.05rem;
    color: rgba(255,255,255,0.78);
    margin: 0;
}

/* ── Cards ────────────────────────────────────────────────────────────── */
.info-card {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px;
    padding: 1.1rem 1.4rem;
    margin-bottom: 1.2rem;
    font-size: 0.9rem;
    color: #aaa;
    line-height: 1.7;
}
.info-card code {
    background: rgba(249,99,2,0.18);
    color: #F96302;
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 0.85rem;
}
.stat-card {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    border-left: 4px solid #F96302;
    text-align: center;
}
.stat-value  { font-size: 2.2rem; font-weight: 800; color: #F96302; line-height:1; }
.stat-label  { font-size: 0.8rem; color: #888; margin-top: 0.4rem; }

/* ── Section headers ─────────────────────────────────────────────────── */
.section-hdr {
    font-size: 1.05rem;
    font-weight: 700;
    color: #f0f0f0;
    border-bottom: 2px solid #F96302;
    padding-bottom: 0.45rem;
    margin: 1.8rem 0 1rem;
}

/* ── Upload zone ─────────────────────────────────────────────────────── */
[data-testid="stFileUploader"] section {
    border: 2px dashed #F96302 !important;
    border-radius: 12px !important;
    background: rgba(249,99,2,0.06) !important;
}

/* ── Primary button ──────────────────────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #F96302, #cc4f00) !important;
    color: #fff !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.65rem 2.2rem !important;
    font-size: 1rem !important;
    transition: opacity .2s !important;
    box-shadow: 0 4px 15px rgba(249,99,2,.35) !important;
}
.stButton > button[kind="primary"]:hover { opacity: .85 !important; }

/* ── Download button ─────────────────────────────────────────────────── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #28a745, #1e7e34) !important;
    color: #fff !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 10px !important;
    box-shadow: 0 4px 15px rgba(40,167,69,.3) !important;
}

/* ── Dataframe ────────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] > div { border-radius: 10px; }
</style>
""",
    unsafe_allow_html=True,
)


# ─── Lazy module imports (keeps startup instant) ───────────────────────────────
@st.cache_resource
def _imports():
    from planogram.loader import load_workbook_from_bytes
    from planogram.logger import PlanogramLogger
    from planogram.allocator import generate_planogram
    from planogram.writer import write_output_bytes

    return load_workbook_from_bytes, PlanogramLogger, generate_planogram, write_output_bytes


# ─── Hero banner ──────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="hero-banner">
  <div class="hero-title">🏗️ Home Depot Planogram Generator</div>
  <p class="hero-sub">
    Upload your Excel workbook to automatically generate a complete
    bay-level planogram layout for every store.
  </p>
</div>
""",
    unsafe_allow_html=True,
)

# ─── Info card ────────────────────────────────────────────────────────────────
st.markdown(
    """
<div class="info-card">
  <strong>Required sheets in workbook:</strong><br>
  &nbsp;&nbsp;• <code>Store List</code> — Store | Current Store POG | Current LFT | Notes<br>
  &nbsp;&nbsp;• <code>Stock SKUs and Displays</code> — Store | Stock SKU | Stock Desc | ... | Facings | Display SKU | Display Desc | ... | Facings | CF<br>
  &nbsp;&nbsp;• <code>Special Order Boards</code> — Store | SO SKU | Description | ... | Facings | ... | CF<br><br>
  <strong>Bay order:</strong> LFT bays are placed <em>first</em>, then POG bays follow.<br>
  <strong>Facings:</strong> Only values 1 or 2 are valid — anything else is automatically clamped to 1.
</div>
""",
    unsafe_allow_html=True,
)

# ─── Upload ───────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">📁 Upload Workbook</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Choose your Excel workbook (.xlsx)",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

st.markdown("")
run_clicked = st.button("▶  Generate Planogram", type="primary")

# ─── Processing ───────────────────────────────────────────────────────────────
if run_clicked:
    if uploaded_file is None:
        st.error("⚠️  Please upload an Excel workbook first.")
        st.stop()

    load_wb, Logger, gen_pog, write_bytes = _imports()
    file_bytes = uploaded_file.getvalue()
    logger = Logger()

    try:
        with st.spinner("📂  Loading workbook…"):
            wb_data = load_wb(file_bytes, logger)

        # ── Column Detection Report (always shown, collapsed by default) ───
        st.markdown("---")
        with st.expander("🔍 Column Detection Report — verify these before generating", expanded=False):
            st.caption(
                "This table shows which actual Excel column header was mapped to each "
                "logical field. If a Facing column looks wrong (e.g. mapped to CF), "
                "that explains incorrect Facing values and SKU pool sizes."
            )
            col_report_rows = [
                # ── Store List ────────────────────────────────────────────
                ("Store List",      "Store ID",          wb_data.cols_sl.get("store", "?")),
                ("Store List",      "Current Store POG", wb_data.cols_sl.get("pog",   "?")),
                ("Store List",      "Current LFT",       wb_data.cols_sl.get("lft",   "?")),
                ("Store List",      "Notes",             wb_data.cols_sl.get("notes", "?")),
                # ── Stock SKUs and Displays ───────────────────────────────
                ("Stock & Display", "Store ID",          wb_data.cols_sd.get("store",      "?")),
                ("Stock & Display", "Stock SKU",         wb_data.cols_sd.get("stock_sku",  "?")),
                ("Stock & Display", "Stock Description", wb_data.cols_sd.get("stock_desc", "?")),
                ("Stock & Display", "⭐ Stock Facing",   wb_data.cols_sd.get("stock_face", "?")),
                ("Stock & Display", "Display SKU",       wb_data.cols_sd.get("disp_sku",   "?")),
                ("Stock & Display", "Display Description", wb_data.cols_sd.get("disp_desc","?")),
                ("Stock & Display", "⭐ Display Facing", wb_data.cols_sd.get("disp_face",  "?")),
                ("Stock & Display", "CF",                wb_data.cols_sd.get("cf",         "?")),
                # ── Special Order Boards ──────────────────────────────────
                ("Special Orders",  "Store ID",          wb_data.cols_so.get("store",   "?")),
                ("Special Orders",  "SO SKU",            wb_data.cols_so.get("so_sku",  "?")),
                ("Special Orders",  "SO Description",    wb_data.cols_so.get("so_desc", "?")),
                ("Special Orders",  "⭐ SO Facing",      wb_data.cols_so.get("so_face", "?")),
                ("Special Orders",  "CF",                wb_data.cols_so.get("cf",      "?")),
            ]
            import pandas as _pd
            _col_df = _pd.DataFrame(col_report_rows, columns=["Sheet", "Logical Field", "→ Actual Excel Column Header Detected"])
            st.dataframe(_col_df, use_container_width=True, hide_index=True)
            st.caption(
                "⭐ Facing columns are the most critical. If 'Stock Facing' and 'Display Facing' "
                "both show the SAME column header, the duplicate Facings column was not detected. "
                "Similarly, if a Facing column shows 'CF', it's reading the wrong column — "
                "this causes facing values to be clamped to 1."
            )

        with st.spinner("⚙️  Generating planogram…"):
            planogram_df, validation_df = gen_pog(wb_data, logger)

        with st.spinner("💾  Writing Excel output…"):
            output_bytes = write_bytes(file_bytes, planogram_df, validation_df)

    except ValueError as exc:
        st.error(f"❌ Data Error: {exc}")
        st.stop()
    except RuntimeError as exc:
        st.error(f"❌ File Error: {exc}")
        st.stop()
    except Exception as exc:
        st.error(f"❌ Unexpected error: {exc}")
        with st.expander("Technical details"):
            st.code(traceback.format_exc())
        st.stop()


    # ─── Results ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-hdr">✅ Generation Complete</div>', unsafe_allow_html=True)

    n_stores = planogram_df["Store"].nunique() if not planogram_df.empty else 0
    n_rows   = len(planogram_df)
    n_issues = len(validation_df)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f'<div class="stat-card"><div class="stat-value">{n_stores:,}</div>'
            f'<div class="stat-label">Stores processed</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="stat-card"><div class="stat-value">{n_rows:,}</div>'
            f'<div class="stat-label">Planogram rows</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        color = "#e74c3c" if n_issues > 0 else "#2ecc71"
        st.markdown(
            f'<div class="stat-card" style="border-left-color:{color};">'
            f'<div class="stat-value" style="color:{color};">{n_issues:,}</div>'
            f'<div class="stat-label">Validation issues</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("")

    # Download button
    st.download_button(
        label="📥  Download Planogram_Output.xlsx",
        data=output_bytes,
        file_name="Planogram_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Validation issues table
    if n_issues > 0:
        st.markdown(
            '<div class="section-hdr">⚠️ Validation Issues</div>',
            unsafe_allow_html=True,
        )
        with st.expander(f"Show {n_issues} issue(s)", expanded=True):
            st.dataframe(
                validation_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Level":   st.column_config.TextColumn("Level",   width="small"),
                    "Store":   st.column_config.TextColumn("Store",   width="small"),
                    "Bay#":    st.column_config.TextColumn("Bay#",    width="small"),
                    "Message": st.column_config.TextColumn("Message", width="large"),
                },
            )
    else:
        st.success("✅ No validation issues found.")

    # Planogram preview
    if not planogram_df.empty:
        st.markdown(
            '<div class="section-hdr">👁️ Planogram Preview (first 200 rows)</div>',
            unsafe_allow_html=True,
        )
        with st.expander("Show preview", expanded=False):
            st.dataframe(
                planogram_df.head(200),
                use_container_width=True,
                hide_index=True,
            )

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p style="text-align:center;color:#444;font-size:.8rem;">'
    "Home Depot Planogram Generator</p>",
    unsafe_allow_html=True,
)

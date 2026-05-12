"""
Preview Page — inspect uploaded export files before running the pipeline.
"""
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Preview · Mobile.de", page_icon="📊", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
    html, body, [class*="css"] { font-family:'IBM Plex Sans',sans-serif; }
    .page-title { font-family:'IBM Plex Mono',monospace; font-size:2rem; font-weight:600; color:#fff; margin-bottom:.2rem; }
    .page-sub   { font-size:.9rem; color:#666; margin-bottom:2rem; }
    .accent     { color:#e8ff4d; }
    .stat-val   { font-family:'IBM Plex Mono',monospace; font-size:1.6rem; font-weight:600; color:#e8ff4d; }
    .stat-lbl   { font-size:.75rem; color:#666; text-transform:uppercase; letter-spacing:.1em; }
    .stButton>button { background:#e8ff4d; color:#0d0d0d; border:none;
        font-family:'IBM Plex Mono',monospace; font-weight:600; font-size:.85rem; padding:.5rem 1.2rem; border-radius:4px; }
    [data-testid="stSidebar"] { background:#111; border-right:1px solid #2a2a2a; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="page-title">Preview <span class="accent">Data</span></div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">Inspect your export files before running the pipeline</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Upload an export or merged file to preview",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if not uploaded:
    st.info("Upload one or more Excel files above to preview their contents.")
    st.stop()

for uf in uploaded:
    uf.seek(0)
    try:
        df = pd.read_excel(uf)
    except Exception as e:
        st.error(f"Could not read {uf.name}: {e}")
        continue

    st.markdown(f"---\n### `{uf.name}`")

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-val">{len(df)}</div><div class="stat-lbl">Rows</div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-val">{len(df.columns)}</div><div class="stat-lbl">Columns</div>', unsafe_allow_html=True)

    if "Website" in df.columns:
        n_dealers = df["Website"].nunique()
        c3.markdown(f'<div class="stat-val">{n_dealers}</div><div class="stat-lbl">Unique Dealers</div>', unsafe_allow_html=True)

    if "Name Of Detail Pige" in df.columns:
        n_kpis = df["Name Of Detail Pige"].nunique()
        c4.markdown(f'<div class="stat-val">{n_kpis}</div><div class="stat-lbl">Unique KPIs</div>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📋 Data", "📈 KPI Summary", "🏢 Dealer List"])

    with tab1:
        st.dataframe(df, use_container_width=True, height=400)

    with tab2:
        if "Name Of Detail Pige" in df.columns and "Detail Pige" in df.columns:
            summary = (
                df.groupby("Name Of Detail Pige")["Detail Pige"]
                .agg(["count", "sum", "mean"])
                .rename(columns={"count": "Count", "sum": "Total", "mean": "Average"})
                .reset_index()
            )
            st.dataframe(summary, use_container_width=True)
        else:
            st.info("No KPI columns found (expected 'Name Of Detail Pige' and 'Detail Pige').")

    with tab3:
        if "Website" in df.columns:
            dealer_counts = df["Website"].value_counts().reset_index()
            dealer_counts.columns = ["Dealer", "Row Count"]
            st.dataframe(dealer_counts, use_container_width=True, height=400)
        else:
            st.info("No 'Website' column found.")

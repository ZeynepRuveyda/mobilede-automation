import streamlit as st

st.set_page_config(
    page_title="Mobile.de Automation",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans', sans-serif;
    }

    .main { background-color: #0d0d0d; }

    [data-testid="stSidebar"] {
        background: #111111;
        border-right: 1px solid #2a2a2a;
    }

    .sidebar-logo {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.1rem;
        font-weight: 600;
        color: #e8ff4d;
        letter-spacing: 0.05em;
        padding: 1rem 0 0.5rem 0;
    }

    .sidebar-sub {
        font-size: 0.7rem;
        color: #555;
        letter-spacing: 0.15em;
        text-transform: uppercase;
        margin-bottom: 1.5rem;
    }

    .hero-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 2.8rem;
        font-weight: 600;
        color: #ffffff;
        line-height: 1.1;
        margin-bottom: 0.3rem;
    }

    .hero-accent {
        color: #e8ff4d;
    }

    .hero-sub {
        font-size: 1rem;
        color: #666;
        margin-bottom: 2.5rem;
        letter-spacing: 0.02em;
    }

    .step-card {
        background: #161616;
        border: 1px solid #242424;
        border-radius: 8px;
        padding: 1.4rem;
        margin-bottom: 1rem;
        transition: border-color 0.2s;
    }

    .step-card:hover { border-color: #e8ff4d44; }

    .step-number {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.7rem;
        color: #e8ff4d;
        letter-spacing: 0.2em;
        text-transform: uppercase;
        margin-bottom: 0.4rem;
    }

    .step-title {
        font-size: 1rem;
        font-weight: 600;
        color: #e0e0e0;
        margin-bottom: 0.3rem;
    }

    .step-desc {
        font-size: 0.85rem;
        color: #777;
        line-height: 1.5;
    }

    .status-badge {
        display: inline-block;
        padding: 0.2rem 0.6rem;
        border-radius: 4px;
        font-size: 0.72rem;
        font-family: 'IBM Plex Mono', monospace;
        font-weight: 600;
        letter-spacing: 0.05em;
    }

    .badge-ready { background: #e8ff4d22; color: #e8ff4d; border: 1px solid #e8ff4d44; }
    .badge-pending { background: #ff884422; color: #ff8844; border: 1px solid #ff884444; }

    .stButton > button {
        background: #e8ff4d;
        color: #0d0d0d;
        border: none;
        font-family: 'IBM Plex Mono', monospace;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 0.05em;
        padding: 0.6rem 1.4rem;
        border-radius: 4px;
        transition: all 0.15s;
    }

    .stButton > button:hover {
        background: #ffffff;
        transform: translateY(-1px);
    }

    div[data-testid="metric-container"] {
        background: #161616;
        border: 1px solid #242424;
        border-radius: 8px;
        padding: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown('<div class="sidebar-logo">MOBILE.DE</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Automation Suite</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**Navigation**")
    st.page_link("Home.py", label="🏠 Ana Sayfa")
    st.page_link("pages/1_Generate_Delivery_File.py", label="⚙️ Template Oluştur")
    st.page_link("pages/2_Preview_Optional.py", label="📊 Dosya Önizleme (Opsiyonel)")
    st.markdown("---")
    st.markdown(
        '<span style="font-size:0.72rem; color:#444; font-family:monospace;">'
        'v1.0.0 · Mobile.de Automation</span>',
        unsafe_allow_html=True,
    )

# ── Main Content ──────────────────────────────────────────────────────────────
st.markdown(
    '<div class="hero-title">Mobile<span class="hero-accent">.de</span><br>Automation</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="hero-sub">Weekly KPI template generator · Export → Merge → Reshape → Deliver</div>',
    unsafe_allow_html=True,
)

# Quick nav buttons
col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    if st.button("▶  Run Pipeline", use_container_width=True):
        st.switch_page("pages/1_Generate_Delivery_File.py")
with col2:
    if st.button("📋  Preview Data", use_container_width=True):
        st.switch_page("pages/2_Preview_Optional.py")

st.markdown("<br>", unsafe_allow_html=True)

# Pipeline overview cards
st.markdown("### Pipeline Steps")

steps = [
    ("STEP 01", "Merge Raw Exports", "Scans the exports folder and concatenates all raw .xlsx/.xls detail export files into a single merged_original_export file.", "merge_original_exports.py"),
    ("STEP 02", "Reshape for KPI Order", "Normalises KPI names via kpi_map, aligns dealer order from the Template, and fills N/A for missing combinations.", "reshape_merged_for_kpi_order.py"),
    ("STEP 03", "Build Final Template", "Reads the reshaped export, carries previous week Values → Last week Value, matches Dealer+KPI and writes the deliverable Template.", "build_from_exports_generic.py"),
    ("STEP 04", "Crawl Update (Optional)", "Applies a crawl_export file on top of the Template to overwrite Value cells matched by Dealer+KPI.", "update_template_from_crawl_export.py"),
]

for num, title, desc, script in steps:
    st.markdown(f"""
    <div class="step-card">
        <div class="step-number">{num}</div>
        <div class="step-title">{title}</div>
        <div class="step-desc">{desc}</div>
        <br>
        <span class="status-badge badge-ready" style="font-size:0.7rem; font-family:monospace; background:#1a1a1a; color:#555; border:1px solid #2a2a2a; padding:0.2rem 0.5rem; border-radius:3px;">{script}</span>
    </div>
    """, unsafe_allow_html=True)

# How-to
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### Quick Start")
with st.expander("How to use this app", expanded=False):
    st.markdown("""
1. **Go to Run Pipeline** page
2. Enter the **week date** (DD.MM.YYYY format, e.g. `28.04.2026`)
3. Upload your **Detail Export files** (one or more `.xlsx`)
4. Upload the **previous week's Template** file
5. *(Optional)* Upload a **Crawl Export** file to overwrite Values
6. Click **Run** — download the finished `Template_<date>.xlsx`
    """)

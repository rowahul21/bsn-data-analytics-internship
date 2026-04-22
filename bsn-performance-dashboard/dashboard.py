"""
2025 Realization Dashboard
==========================
A Streamlit dashboard for analyzing branch and sub-branch performance
across three loan products: Non KPR (S), KPR Subsidi (NS), KPR Non Subsidi (NK).

Project folder structure (required for Streamlit Cloud deployment):
──────────────────────────────────────────────────────────────────
  bsn-performance-dashboard/
  ├── dashboard.py          ← this script
  ├── requirements.txt      ← pip dependencies
  └── data/
      └── DUMMY_Database.xlsx   ← data file lives here

WHY this structure?
  On Streamlit Cloud the working directory is the repo root, NOT the
  folder that contains dashboard.py.  Using __file__ to build an
  absolute path makes the code work identically locally AND in the cloud.

Run locally:
    cd bsn-performance-dashboard
    streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path   # ← used for deployment-safe file paths

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="2025 Realization Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL STYLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .main { background-color: #f7f9fc; }

    /* Metric cards */
    [data-testid="metric-container"] {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    [data-testid="metric-container"] label {
        font-size: 13px !important;
        color: #64748b !important;
        font-weight: 500;
    }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        font-size: 26px !important;
        font-weight: 700;
        color: #1e293b !important;
    }

    /* Section headers */
    .section-title {
        font-size: 17px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 4px;
        margin-top: 8px;
    }
    .section-sub {
        font-size: 13px;
        color: #64748b;
        margin-bottom: 12px;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #1e293b;
    }
    [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiselect label {
        font-size: 13px;
        font-weight: 500;
    }

    /* Divider */
    hr { border-color: #e2e8f0; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL COLOR THEME
# ─────────────────────────────────────────────────────────────────────────────
#
# HOW TO USE THIS SECTION:
#   All colors in the entire dashboard are defined here and ONLY here.
#   Charts reference these dictionaries instead of hardcoding hex values.
#   To change a color → edit it here once, all charts update automatically.
#   To add a new region → add one line to REGION_COLOR_MAP, nothing else.
#
# ── Brand anchor ──────────────────────────────────────────────────────────
BSN_GREEN   = "#006747"   # BSN primary brand color — used as anchor for theme

# ── Product color map  ────────────────────────────────────────────────────
# Plotly px charts accept color_discrete_map=PRODUCT_COLORS  (dict, key = label)
# Each product always appears in the same color across every chart.
#
#   Non KPR         → Amber  : warm, energetic — largest product by volume
#   KPR Subsidi     → BSN Green : brand color  — government-backed product
#   KPR Non Subsidi → Blue   : cool, stable    — smaller / premium segment
#
PRODUCT_COLORS = {
    "Non KPR":          "#F59E0B",   # amber      H=38°
    "KPR Subsidi":      "#006747",   # BSN green  H=161°
    "KPR Non Subsidi":  "#2563EB",   # blue       H=221°
}

# ── Region color map  ─────────────────────────────────────────────────────
# Plotly px charts accept color_discrete_map=REGION_COLOR_MAP (dict, key = label)
#
# Design rule: hues are spread ~90–130° apart on the color wheel so each
# region is instantly distinguishable even in small legends or on projectors.
#
#   Jakarta Jabar Banten    → BSN Green H=161°  (dominant region → brand color)
#   Jateng DIY Jatim Nusra  → Amber     H=38°   (Δ123° from green)
#   Sumatera                → Blue      H=221°   (Δ60°  from amber)
#   Kalimantan Sulawesi     → Red/Coral H=0°     (Δ139° from blue)
#
# Adding a new region? Just add a new line here:
#   "Papua": "#06B6D4",   # teal H=189°
#
REGION_COLOR_MAP = {
    "Jakarta Jabar Banten":    "#006747",   # BSN green
    "Jateng DIY Jatim Nusra":  "#F59E0B",   # amber
    "Sumatera":                "#2563EB",   # blue
    "Kalimantan Sulawesi":     "#DC2626",   # red
}

# Convenience: ordered list of region colors (same order as the dict above).
# Use this ONLY when a Plotly chart doesn't support color_discrete_map,
# e.g., px.pie with color_discrete_sequence=REGION_COLOR_SEQ.
# ⚠️  Do NOT pass the dict REGION_COLOR_MAP to color_discrete_sequence —
#     Plotly expects a list there, not a dict, and will silently ignore it.
# ⚠️  Do NOT use this for px.pie — use color="WILAYAH" + color_discrete_map
#     instead, so colors are bound by name and not by row position.
REGION_COLOR_SEQ = list(REGION_COLOR_MAP.values())

# ── Semantic / status colors  ─────────────────────────────────────────────
# Used for RKAP hit/miss coloring and MoM trend line.
# Keeping these in the theme means they can be updated consistently too.
COLOR_POSITIVE  = "#006747"   # BSN green — target met     (≥ 100 % RKAP)
COLOR_NEGATIVE  = "#EF4444"   # red       — target missed  (< 100 % RKAP)
COLOR_MOM_LINE  = "#EF4444"   # red line  — MoM trend overlay on bar chart
COLOR_MOM_BAR   = "#006747"   # BSN green — monthly total bars

# ── Continuous (gradient) color scale  ────────────────────────────────────
# Used for single-metric bar charts where bar height/length = the color value.
# Low → High maps light green tint → BSN dark green.
GRADIENT_SCALE  = ["#A7F3D0", BSN_GREEN]   # light mint → BSN green

# ─────────────────────────────────────────────────────────────────────────────
# FILE PATH  –  deployment-safe approach using __file__
# ─────────────────────────────────────────────────────────────────────────────
# WHY __file__ instead of a plain filename like "DUMMY_Database.xlsx"?
#
#   When you run a script locally, Python's working directory is wherever
#   you opened your terminal — which usually matches the script's folder.
#   So a plain filename happens to work.
#
#   On Streamlit Cloud, the server clones your entire GitHub repo, then
#   sets the working directory to the REPO ROOT, not to the subfolder
#   that contains dashboard.py.  So if your script is at:
#       bsn-performance-dashboard/dashboard.py
#   and your data is at:
#       bsn-performance-dashboard/data/DUMMY_Database.xlsx
#   then pd.read_excel("DUMMY_Database.xlsx") looks for the file at
#   the repo root and raises FileNotFoundError.
#
#   The fix:  build the path relative to THIS script file (__file__),
#   not relative to wherever Python happens to be running from.
#
#   Path(__file__).parent  →  the folder containing dashboard.py
#   / "data"               →  step into the data/ subfolder
#   / "DUMMY_Database.xlsx"→  the actual file
#
#   This always resolves to the right location on any machine or cloud server.

# ── Where this script lives (works locally AND on Streamlit Cloud) ──
THIS_DIR = Path(__file__).parent

# ── Path to your data file — put the Excel file in a /data subfolder ──
DATA_FILE = THIS_DIR / "data" / "DUMMY_Database.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# DATA PROCESSING  (separated from loading so it can be reused for uploads)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data
def process_sheets(sheets: dict):
    """
    Accepts a dict of DataFrames (one per sheet) and returns the three
    cleaned DataFrames used by the dashboard.

    Keeping processing separate from file I/O means we can call this
    function whether the data came from disk OR from st.file_uploader.
    """
    # ── Historical: monthly realization per branch & product ──
    hist = sheets["Historical"].copy()
    hist["BULAN"]       = pd.to_datetime(hist["BULAN"])
    hist["MONTH"]       = hist["BULAN"].dt.month
    hist["MONTH_NAME"]  = hist["BULAN"].dt.strftime("%b")       # Jan, Feb …
    hist["MONTH_LABEL"] = hist["BULAN"].dt.strftime("%b %Y")

    # ── ALL_FIKS: branch-level snapshot (with WILAYAH/region) ──
    all_fiks = sheets["ALL_FIKS"].copy()
    all_fiks["total_real_JT"]   = all_fiks["S_DES_25_(JT)"]  + all_fiks["NS_DES_25_(JT)"]  + all_fiks["NK_DES_25_(JT)"]
    all_fiks["total_rkap_JT"]   = all_fiks["S_RKAP_25_(JT)"] + all_fiks["NS_RKAP_25_(JT)"] + all_fiks["NK_RKAP_25_(JT)"]
    all_fiks["total_units"]      = all_fiks["S_DES_25_(U)"]   + all_fiks["NS_DES_25_(U)"]   + all_fiks["NK_DES_25_(U)"]
    all_fiks["rkap_achieve_pct"] = (all_fiks["total_real_JT"] / all_fiks["total_rkap_JT"]).replace([np.inf, -np.inf], np.nan)
    all_fiks["avg_deal_size"]    = (all_fiks["total_real_JT"] / all_fiks["total_units"]).replace([np.inf, -np.inf], np.nan)
    all_fiks["total_real_24_JT"] = all_fiks["S_DES_24_(JT)"]  + all_fiks["NS_DES_24_(JT)"]  + all_fiks["NK_DES_24_(JT)"]
    all_fiks["yoy_growth"]       = ((all_fiks["total_real_JT"] - all_fiks["total_real_24_JT"]) / all_fiks["total_real_24_JT"]).replace([np.inf, -np.inf], np.nan)

    # ── KCPS: sub-branch snapshot ──
    kcps = sheets["KCPS"].copy()
    kcps["total_real_JT"]    = kcps["S_DES_25_(JT)"] + kcps["NS_DES_25_(JT)"] + kcps["NK_DES_25_(JT)"]
    kcps["total_units"]      = kcps["S_DES_25_(U)"]  + kcps["NS_DES_25_(U)"]  + kcps["NK_DES_25_(U)"]
    kcps["rkap_achieve_pct"] = (kcps["total_real_JT"] / (kcps["S_RKAP_25_(JT)"].fillna(0) + kcps["NS_RKAP_25_(JT)"].fillna(0) + kcps["NK_RKAP_25_(JT)"].fillna(0))).replace([np.inf, -np.inf], np.nan)
    kcps["avg_deal_size"]    = (kcps["total_real_JT"] / kcps["total_units"]).replace([np.inf, -np.inf], np.nan)

    # ── Merge region label into KCPS and Historical ──
    branch_region = all_fiks[["KODE_KCS", "WILAYAH"]].drop_duplicates()
    kcps = kcps.merge(branch_region, on="KODE_KCS", how="left")
    hist = hist.merge(branch_region, on="KODE_KCS", how="left")

    return hist, all_fiks, kcps


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING  –  tries bundled file first, falls back to file uploader
# ─────────────────────────────────────────────────────────────────────────────
def load_data(source):
    """
    Load all sheets from either a file path (Path object) or an
    in-memory uploaded file (UploadedFile from st.file_uploader).
    Returns the raw dict of DataFrames so process_sheets() can clean them.

    Using a separate non-cached loader means st.cache_data on
    process_sheets() still works — cache_data can't cache file handles.
    """
    return pd.read_excel(source, sheet_name=None)


# ── Try to load the bundled data file ──────────────────────────────────────
# Priority order:
#   1. Bundled file at data/DUMMY_Database.xlsx  (works on Streamlit Cloud
#      once the file is committed to the GitHub repo inside the data/ folder)
#   2. File uploaded manually via st.file_uploader  (fallback — useful if
#      you don't want to commit sensitive data to GitHub)

if DATA_FILE.exists():
    # ✅ Bundled file found — load silently
    raw_sheets = load_data(DATA_FILE)
    hist, all_fiks, kcps = process_sheets(raw_sheets)

else:
    # ⚠️ Bundled file NOT found — show a friendly uploader instead of crashing
    st.warning(
        "📂 **Data file not found** — `data/DUMMY_Database.xlsx` is missing from the app folder.\n\n"
        "This usually means the file was not committed to the GitHub repository. "
        "You can either:\n"
        "- Add the file to `bsn-performance-dashboard/data/` in your repo, **or**\n"
        "- Upload it manually below to continue.",
        icon="⚠️",
    )

    uploaded = st.file_uploader(
        "Upload DUMMY_Database.xlsx to continue",
        type=["xlsx"],
        help="Upload the same Excel file that was used to build this dashboard.",
    )

    if uploaded is None:
        # Stop rendering — nothing else can work without data
        st.info("👆 Please upload the Excel file to load the dashboard.")
        st.stop()   # Halts execution cleanly; no scary error traceback shown

    # File uploaded — process it exactly the same way as the bundled file
    try:
        raw_sheets = load_data(uploaded)
        hist, all_fiks, kcps = process_sheets(raw_sheets)
        st.success("✅ File uploaded and loaded successfully!")
    except Exception as e:
        st.error(
            f"❌ Could not read the uploaded file.\n\n"
            f"Make sure it is the correct DUMMY_Database.xlsx with sheets: "
            f"**Historical**, **ALL_FIKS**, **KCPS**, **2025**.\n\n"
            f"Technical detail: `{e}`"
        )
        st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Dashboard Filters")
    st.markdown("---")

    # Region filter
    regions = ["All Regions"] + sorted(all_fiks["WILAYAH"].dropna().unique().tolist())
    sel_region = st.selectbox("Region (WILAYAH)", regions)

    # Branch filter (cascades from region)
    if sel_region == "All Regions":
        branch_options = sorted(all_fiks["CABANG"].unique().tolist())
    else:
        branch_options = sorted(all_fiks[all_fiks["WILAYAH"] == sel_region]["CABANG"].unique().tolist())

    sel_branch = st.selectbox("Branch (KCS)", ["All Branches"] + branch_options)

    # Sub-branch filter (cascades from branch)
    if sel_branch == "All Branches":
        branch_kcs_codes = all_fiks["KODE_KCS"].unique() if sel_region == "All Regions" else all_fiks[all_fiks["WILAYAH"] == sel_region]["KODE_KCS"].unique()
        sub_options = sorted(kcps[kcps["KODE_KCS"].isin(branch_kcs_codes)]["CABANG"].unique().tolist())
    else:
        kcs_code = all_fiks[all_fiks["CABANG"] == sel_branch]["KODE_KCS"].values[0]
        sub_options = sorted(kcps[kcps["KODE_KCS"] == kcs_code]["CABANG"].unique().tolist())

    sel_sub = st.selectbox("Sub-Branch (KCPS)", ["All Sub-Branches"] + sub_options)

    # Product filter
    st.markdown("---")
    product_options = ["Non KPR", "KPR Subsidi", "KPR Non Subsidi"]
    sel_products = st.multiselect("Product", product_options, default=product_options)

    # Month filter
    month_labels = sorted(hist["MONTH_LABEL"].unique().tolist(), key=lambda x: pd.to_datetime(x, format="%b %Y"))
    sel_months = st.multiselect("Months", month_labels, default=month_labels)

    st.markdown("---")
    st.markdown("<small style='color:#94a3b8'>Data: 2025 Realization<br>Branches: 35 | Sub-Branches: 118</small>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────────────────────────────────────
# Filter historical data
hist_f = hist[hist["PRODUK"].isin(sel_products) & hist["MONTH_LABEL"].isin(sel_months)].copy()

if sel_region != "All Regions":
    hist_f = hist_f[hist_f["WILAYAH"] == sel_region]
if sel_branch != "All Branches":
    hist_f = hist_f[hist_f["CABANG"] == sel_branch]

# Filter ALL_FIKS
all_fiks_f = all_fiks.copy()
if sel_region != "All Regions":
    all_fiks_f = all_fiks_f[all_fiks_f["WILAYAH"] == sel_region]
if sel_branch != "All Branches":
    all_fiks_f = all_fiks_f[all_fiks_f["CABANG"] == sel_branch]

# Filter KCPS
kcps_f = kcps.copy()
if sel_region != "All Regions":
    kcps_f = kcps_f[kcps_f["WILAYAH"] == sel_region]
if sel_branch != "All Branches":
    kcps_f = kcps_f[kcps_f["KODE_KCS"].isin(all_fiks_f["KODE_KCS"].values)]
if sel_sub != "All Sub-Branches":
    kcps_f = kcps_f[kcps_f["CABANG"] == sel_sub]

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: format numbers
# ─────────────────────────────────────────────────────────────────────────────
def fmt_jt(val):
    """Format Juta (millions) value into readable string."""
    if val >= 1_000_000:
        return f"Rp {val/1_000_000:.2f} T"
    elif val >= 1_000:
        return f"Rp {val/1_000:.1f} M"
    else:
        return f"Rp {val:,.0f} Jt"

def fmt_pct(val):
    return f"{val*100:.1f}%"

def safe_region_color_map(data_regions, color_map=REGION_COLOR_MAP,
                          fallback="#94a3b8"):
    """
    Build a complete {region: color} dict that covers every region actually
    present in the currently-filtered data.

    WHY this exists:
      When the user applies a sidebar filter, hist_f may contain only 1 or 2
      regions instead of all 4.  Passing REGION_COLOR_MAP directly still works
      for those regions — Plotly ignores unused keys — so this function is
      mainly a safety net for a different problem: if a region name in the data
      does NOT exist in REGION_COLOR_MAP (e.g. a new branch was added to the
      dataset but not yet to the theme dict), Plotly silently falls back to its
      default D3 palette and the color becomes inconsistent.

      This function catches that case and applies a neutral fallback color
      (#94a3b8 slate-gray) instead of letting Plotly pick something random.

    Args:
        data_regions : iterable of region name strings from the current df
        color_map    : the master dict (defaults to REGION_COLOR_MAP)
        fallback     : hex color for any region not in color_map

    Returns:
        dict  {region_name: hex_color}  — one entry per region in data_regions

    Usage:
        # Instead of:  color_discrete_map=REGION_COLOR_MAP
        # Write:       color_discrete_map=safe_region_color_map(df["WILAYAH"].unique())
    """
    result = {}
    for region in data_regions:
        result[region] = color_map.get(region, fallback)
    return result

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("# 2025 Realization Dashboard")
st.markdown(
    f"<p style='color:#64748b;margin-top:-8px;font-size:14px;'>"
    f"Showing: <b>{sel_region}</b> · <b>{sel_branch}</b> · <b>{sel_sub}</b> · "
    f"Products: <b>{', '.join(sel_products) if sel_products else 'None'}</b></p>",
    unsafe_allow_html=True,
)
st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 – KPI SCORECARDS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Key Performance Indicators</p>', unsafe_allow_html=True)
st.markdown('<p class="section-sub">Cumulative 2025 figures based on current filters</p>', unsafe_allow_html=True)

total_real   = hist_f["REALISASI_JT"].sum()
total_units  = int(all_fiks_f["total_units"].sum())
avg_rkap_pct = all_fiks_f["rkap_achieve_pct"].mean()
avg_yoy      = all_fiks_f["yoy_growth"].mean()
avg_deal     = all_fiks_f["avg_deal_size"].mean()
n_branches   = all_fiks_f["CABANG"].nunique()

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Total Realization",  fmt_jt(total_real))
k2.metric("Total Units",        f"{total_units:,}")
k3.metric("Active Branches",    str(n_branches))
k4.metric("Avg RKAP Achieve",   fmt_pct(avg_rkap_pct) if not np.isnan(avg_rkap_pct) else "–")
k5.metric("YoY Growth (Dec)",   fmt_pct(avg_yoy) if not np.isnan(avg_yoy) else "–")
k6.metric("Avg Deal Size",      f"Rp {avg_deal:,.0f} Jt" if not np.isnan(avg_deal) else "–")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 – MONTHLY TREND
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Monthly Realization Trend</p>', unsafe_allow_html=True)
st.markdown('<p class="section-sub">Total realization (in Juta IDR) per month, split by product</p>', unsafe_allow_html=True)

monthly = (
    hist_f.groupby(["MONTH", "MONTH_LABEL", "PRODUK"])["REALISASI_JT"]
    .sum()
    .reset_index()
    .sort_values("MONTH")
)

fig_trend = px.area(
    monthly,
    x="MONTH_LABEL",
    y="REALISASI_JT",
    color="PRODUK",
    color_discrete_map=PRODUCT_COLORS,
    labels={"REALISASI_JT": "Realization (Jt)", "MONTH_LABEL": "Month", "PRODUK": "Product"},
    category_orders={"MONTH_LABEL": month_labels},
)
fig_trend.update_layout(
    height=340,
    plot_bgcolor="white",
    paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    xaxis=dict(showgrid=False),
    yaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
)
st.plotly_chart(fig_trend, use_container_width=True)

# Monthly MoM growth line
monthly_total = monthly.groupby(["MONTH", "MONTH_LABEL"])["REALISASI_JT"].sum().reset_index().sort_values("MONTH")
monthly_total["MoM_%"] = monthly_total["REALISASI_JT"].pct_change() * 100

fig_mom = go.Figure()
fig_mom.add_bar(
    x=monthly_total["MONTH_LABEL"],
    y=monthly_total["REALISASI_JT"],
    name="Total Realization",
    marker_color=COLOR_MOM_BAR,   # ← was hardcoded "#3b82f6" (blue, off-theme)
    opacity=0.7,
)
fig_mom.add_scatter(
    x=monthly_total["MONTH_LABEL"],
    y=monthly_total["MoM_%"],
    name="MoM Growth (%)",
    yaxis="y2",
    mode="lines+markers",
    line=dict(color=COLOR_MOM_LINE, width=2),   # ← was hardcoded "#ef4444"
    marker=dict(size=7),
)
fig_mom.update_layout(
    height=280,
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=60, t=10, b=10),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    xaxis=dict(showgrid=False),
    yaxis=dict(title="Realization (Jt)", showgrid=True, gridcolor="#f1f5f9"),
    yaxis2=dict(title="MoM Growth (%)", overlaying="y", side="right", showgrid=False, ticksuffix="%"),
)
st.plotly_chart(fig_mom, use_container_width=True)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 – BRANCH COMPARISON
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Branch Performance Comparison</p>', unsafe_allow_html=True)
st.markdown('<p class="section-sub">Full-year realization vs RKAP target per branch</p>', unsafe_allow_html=True)

col_a, col_b = st.columns(2)

# Chart A: Total realization per branch (horizontal bar)
branch_real = (
    hist_f.groupby("CABANG")["REALISASI_JT"].sum()
    .reset_index()
    .sort_values("REALISASI_JT", ascending=True)
    .tail(20)  # top 20 to keep readable
)

fig_branch_bar = px.bar(
    branch_real,
    x="REALISASI_JT", y="CABANG",
    orientation="h",
    labels={"REALISASI_JT": "Realization (Jt)", "CABANG": ""},
    color="REALISASI_JT",
    color_continuous_scale=GRADIENT_SCALE,   # ← was hardcoded ["#bfdbfe","#1d4ed8"] (blue, off-theme)
)
fig_branch_bar.update_coloraxes(showscale=False)
fig_branch_bar.update_layout(
    height=460, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    yaxis=dict(showgrid=False),
)
col_a.plotly_chart(fig_branch_bar, use_container_width=True)

# Chart B: RKAP achievement % per branch
rkap_data = all_fiks_f[["CABANG", "rkap_achieve_pct"]].dropna().sort_values("rkap_achieve_pct", ascending=True)
# Color each bar green (target met) or red (target missed) using theme constants.
# ← was hardcoded "#10b981" / "#ef4444" inline
rkap_data["color"] = rkap_data["rkap_achieve_pct"].apply(
    lambda x: COLOR_POSITIVE if x >= 1.0 else COLOR_NEGATIVE
)

fig_rkap = go.Figure(go.Bar(
    x=rkap_data["rkap_achieve_pct"] * 100,
    y=rkap_data["CABANG"],
    orientation="h",
    marker_color=rkap_data["color"],
    text=[f"{v*100:.1f}%" for v in rkap_data["rkap_achieve_pct"]],
    textposition="outside",
))
fig_rkap.add_vline(x=100, line_width=2, line_dash="dash", line_color="#64748b", annotation_text="100% target")
fig_rkap.update_layout(
    height=460, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=60, t=10, b=10),
    xaxis=dict(title="RKAP Achievement (%)", showgrid=True, gridcolor="#f1f5f9"),
    yaxis=dict(showgrid=False),
)
col_b.plotly_chart(fig_rkap, use_container_width=True)

# Branch product mix (stacked bar)
st.markdown('<p class="section-sub">Product mix per branch</p>', unsafe_allow_html=True)
branch_mix = (
    hist_f.groupby(["CABANG", "PRODUK"])["REALISASI_JT"]
    .sum().reset_index()
)
branch_order = hist_f.groupby("CABANG")["REALISASI_JT"].sum().sort_values(ascending=False).index.tolist()

fig_mix = px.bar(
    branch_mix,
    x="CABANG", y="REALISASI_JT",
    color="PRODUK",
    color_discrete_map=PRODUCT_COLORS,
    labels={"REALISASI_JT": "Realization (Jt)", "CABANG": "Branch"},
    category_orders={"CABANG": branch_order},
)
fig_mix.update_layout(
    height=320, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    xaxis=dict(showgrid=False, tickangle=45),
    yaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    barmode="stack",
)
st.plotly_chart(fig_mix, use_container_width=True)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 – SUB-BRANCH COMPARISON
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Sub-Branch Performance Comparison</p>', unsafe_allow_html=True)
st.markdown('<p class="section-sub">Realization and RKAP achievement across sub-branches (KCPS)</p>', unsafe_allow_html=True)

col_c, col_d = st.columns(2)

# Sub-branch realization (top 25)
sub_real = kcps_f[["CABANG", "total_real_JT", "WILAYAH"]].dropna(subset=["total_real_JT"])
sub_real = sub_real.sort_values("total_real_JT", ascending=True).tail(25)

fig_sub_bar = px.bar(
    sub_real,
    x="total_real_JT", y="CABANG",
    orientation="h",
    color="WILAYAH",
    # color_discrete_map ensures each region name maps to its fixed color.
    # ← was color_discrete_sequence=REGION_COLORS (Set2 palette — random order,
    #   no name binding, and a dict was incorrectly passed as a sequence which
    #   Plotly silently ignores, producing wrong or default colors).
    color_discrete_map=REGION_COLOR_MAP,
    labels={"total_real_JT": "Realization (Jt)", "CABANG": "", "WILAYAH": "Region"},
)
fig_sub_bar.update_layout(
    height=520, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    yaxis=dict(showgrid=False),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
)
col_c.plotly_chart(fig_sub_bar, use_container_width=True)

# Sub-branch RKAP achievement
sub_rkap = kcps_f[["CABANG", "rkap_achieve_pct"]].dropna().sort_values("rkap_achieve_pct", ascending=True).tail(25)
# ← was hardcoded "#10b981" / "#ef4444" inline
sub_rkap["color"] = sub_rkap["rkap_achieve_pct"].apply(
    lambda x: COLOR_POSITIVE if x >= 1.0 else COLOR_NEGATIVE
)

fig_sub_rkap = go.Figure(go.Bar(
    x=sub_rkap["rkap_achieve_pct"] * 100,
    y=sub_rkap["CABANG"],
    orientation="h",
    marker_color=sub_rkap["color"],
    text=[f"{v*100:.1f}%" for v in sub_rkap["rkap_achieve_pct"]],
    textposition="outside",
))
fig_sub_rkap.add_vline(x=100, line_width=2, line_dash="dash", line_color="#64748b")
fig_sub_rkap.update_layout(
    height=520, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=70, t=10, b=10),
    xaxis=dict(title="RKAP Achievement (%)", showgrid=True, gridcolor="#f1f5f9"),
    yaxis=dict(showgrid=False),
)
col_d.plotly_chart(fig_sub_rkap, use_container_width=True)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 – ACQUISITION EFFICIENCY
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Acquisition Efficiency</p>', unsafe_allow_html=True)
st.markdown('<p class="section-sub">Units acquired vs realization achieved — which branches do more with fewer units?</p>', unsafe_allow_html=True)

eff_data = all_fiks_f[["CABANG", "WILAYAH", "total_units", "total_real_JT", "avg_deal_size", "rkap_achieve_pct"]].dropna()

# Bubble chart: x=units, y=realization, size=avg deal size, color=region
fig_bubble = px.scatter(
    eff_data,
    x="total_units",
    y="total_real_JT",
    size="avg_deal_size",
    color="WILAYAH",
    hover_name="CABANG",
    # color_discrete_map binds each region name to its fixed color.
    # ← was color_discrete_sequence=REGION_COLORS (same bug as fig_sub_bar)
    color_discrete_map=REGION_COLOR_MAP,
    labels={
        "total_units":    "Total Units",
        "total_real_JT":  "Total Realization (Jt)",
        "avg_deal_size":  "Avg Deal Size (Jt)",
        "WILAYAH":        "Region",
    },
    size_max=45,
)
# Add a diagonal reference line (trend)
max_units = eff_data["total_units"].max()
fig_bubble.add_shape(
    type="line",
    x0=0, y0=0,
    x1=max_units, y1=eff_data["total_real_JT"].max(),
    line=dict(color="#94a3b8", width=1, dash="dot"),
)
fig_bubble.update_layout(
    height=420, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    yaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
)
st.plotly_chart(fig_bubble, use_container_width=True)

# Average deal size ranking (bar chart)
deal_df = eff_data.sort_values("avg_deal_size", ascending=True)

fig_deal = px.bar(
    deal_df,
    x="avg_deal_size", y="CABANG",
    orientation="h",
    color="avg_deal_size",
    color_continuous_scale=GRADIENT_SCALE,   # ← was ["#bfdbfe","#1d4ed8"] (blue, off-theme)
    labels={"avg_deal_size": "Avg Deal Size (Jt)", "CABANG": ""},
    text=deal_df["avg_deal_size"].apply(lambda x: f"Rp {x:,.0f} Jt"),
)
fig_deal.update_coloraxes(showscale=False)
fig_deal.update_layout(
    height=max(300, len(deal_df) * 22),
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=120, t=10, b=10),
    xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    yaxis=dict(showgrid=False),
)
st.plotly_chart(fig_deal, use_container_width=True)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 – REGION SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Performance by Region (WILAYAH)</p>', unsafe_allow_html=True)
st.markdown('<p class="section-sub">Aggregated realization across the four operating regions</p>', unsafe_allow_html=True)

reg_data = (
    hist_f.groupby(["WILAYAH", "PRODUK"])["REALISASI_JT"]
    .sum().reset_index()
)

col_e, col_f = st.columns(2)

fig_reg_bar = px.bar(
    reg_data,
    x="WILAYAH", y="REALISASI_JT",
    color="PRODUK",
    color_discrete_map=PRODUCT_COLORS,
    barmode="group",
    labels={"REALISASI_JT": "Realization (Jt)", "WILAYAH": "Region"},
)
fig_reg_bar.update_layout(
    height=320, plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    xaxis=dict(showgrid=False, tickangle=15),
    yaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
)
col_e.plotly_chart(fig_reg_bar, use_container_width=True)

reg_total = hist_f.groupby("WILAYAH")["REALISASI_JT"].sum().reset_index()
fig_pie = px.pie(
    reg_total,
    names="WILAYAH", values="REALISASI_JT",
    # px.pie supports color_discrete_map — use it so each region name
    # is reliably bound to its fixed color regardless of row order.
    # ← was color_discrete_sequence=REGION_COLORS (Set2 variable, unrelated to theme,
    #   AND a dict was previously passed here which Plotly silently ignores)
    # Rule to remember:
    #   color_discrete_map only works when color= points to a column whose
    #   values match the keys of your map dict.
    color="WILAYAH",
    color_discrete_map=safe_region_color_map(reg_total["WILAYAH"].unique()),
    hole=0.45,
)
fig_pie.update_traces(textposition="outside", textinfo="percent+label")
fig_pie.update_layout(
    height=320, paper_bgcolor="white",
    margin=dict(l=10, r=10, t=10, b=10),
    showlegend=False,
)
col_f.plotly_chart(fig_pie, use_container_width=True)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 – DATA TABLE
# ─────────────────────────────────────────────────────────────────────────────
with st.expander("Raw Branch Data Table", expanded=False):
    display_cols = ["CABANG", "WILAYAH", "total_units", "total_real_JT", "total_rkap_JT",
                    "rkap_achieve_pct", "avg_deal_size", "yoy_growth"]
    tbl = all_fiks_f[display_cols].copy()
    tbl.columns = ["Branch", "Region", "Total Units", "Realization (Jt)",
                   "RKAP Target (Jt)", "RKAP Achieve %", "Avg Deal Size (Jt)", "YoY Growth"]
    tbl["RKAP Achieve %"] = (tbl["RKAP Achieve %"] * 100).round(1).astype(str) + "%"
    tbl["YoY Growth"]     = (tbl["YoY Growth"]     * 100).round(1).astype(str) + "%"
    tbl = tbl.sort_values("Realization (Jt)", ascending=False).reset_index(drop=True)
    st.dataframe(tbl, use_container_width=True)

with st.expander("Sub-Branch Data Table", expanded=False):
    sub_tbl = kcps_f[["CABANG", "KODE_KCS", "WILAYAH", "total_units", "total_real_JT", "avg_deal_size", "rkap_achieve_pct"]].copy()
    sub_tbl.columns = ["Sub-Branch", "KCS Code", "Region", "Total Units", "Realization (Jt)", "Avg Deal Size (Jt)", "RKAP Achieve %"]
    sub_tbl["RKAP Achieve %"] = (sub_tbl["RKAP Achieve %"] * 100).round(1).astype(str) + "%"
    sub_tbl = sub_tbl.sort_values("Realization (Jt)", ascending=False).reset_index(drop=True)
    st.dataframe(sub_tbl, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#94a3b8;font-size:12px;'>"
    "2025 Realization Dashboard · Data source: DUMMY_Database.xlsx · "
    "Products: Non KPR (S) · KPR Subsidi (NS) · KPR Non Subsidi (NK)"
    "</p>",
    unsafe_allow_html=True,
)

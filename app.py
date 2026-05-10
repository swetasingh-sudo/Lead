"""
IS Lead Dashboard — Streamlit App
==================================
Run with:
    pip install streamlit plotly pandas openpyxl
    streamlit run is_lead_dashboard.py

Upload:
  • File 1 → New leads created yesterday
  • File 2 → Leads that had activity yesterday
Accepts .xlsx, .xls, or .csv
"""

import io
import warnings
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="IS Lead Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# GLOBAL STYLES
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp { background: #0d1117; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: #161b22 !important;
    border-right: 1px solid #21262d;
}

/* Metric cards */
[data-testid="metric-container"] {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 10px;
    padding: 14px 18px;
}
[data-testid="metric-container"] label { color: #8b949e !important; font-size: 11px !important; text-transform: uppercase; letter-spacing: .06em; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #e6edf3 !important; font-size: 26px !important; font-weight: 700 !important; }
[data-testid="metric-container"] [data-testid="stMetricDelta"] { font-size: 12px !important; }

/* Tabs */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: #161b22;
    border-radius: 10px;
    gap: 4px;
    padding: 6px;
    border: 1px solid #21262d;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    background: transparent;
    border-radius: 8px;
    color: #8b949e;
    font-size: 13px;
    font-weight: 500;
    padding: 8px 16px;
}
[data-testid="stTabs"] [data-baseweb="tab"]:hover { color: #e6edf3; background: #21262d; }
[data-testid="stTabs"] [aria-selected="true"] {
    background: #1f6feb !important;
    color: #fff !important;
}

/* Expanders */
[data-testid="stExpander"] {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 10px;
}

/* DataFrames */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
.dataframe-container { background: #161b22; border-radius: 8px; }

/* File uploader */
[data-testid="stFileUploader"] {
    background: #161b22;
    border: 2px dashed #30363d;
    border-radius: 10px;
    padding: 8px;
}
[data-testid="stFileUploader"]:hover { border-color: #1f6feb; }

/* Selectbox */
[data-testid="stSelectbox"] > div > div {
    background: #161b22 !important;
    border-color: #30363d !important;
    color: #e6edf3 !important;
}

/* General text */
h1, h2, h3, h4 { color: #e6edf3 !important; }
p, li { color: #8b949e; }

/* Alert / info boxes */
.health-box {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 14px;
}
.health-ok   { border-left: 4px solid #3fb950; }
.health-warn { border-left: 4px solid #d29922; }
.health-err  { border-left: 4px solid #f85149; }
.health-title { font-weight: 600; font-size: 14px; margin-bottom: 6px; }
.health-ok .health-title   { color: #3fb950; }
.health-warn .health-title { color: #d29922; }
.health-err .health-title  { color: #f85149; }

.kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 20px; }
.kpi { background:#161b22; border:1px solid #21262d; border-radius:10px;
       padding:14px 18px; min-width:140px; flex:1; }
.kpi .kv { font-size:28px; font-weight:700; color:#e6edf3; }
.kpi .kl { font-size:11px; color:#8b949e; text-transform:uppercase; letter-spacing:.06em; }
.kpi .ks { font-size:12px; color:#484f58; margin-top:2px; }
.kpi.green .kv { color:#3fb950; }
.kpi.red   .kv { color:#f85149; }
.kpi.blue  .kv { color:#58a6ff; }
.kpi.yellow .kv { color:#d29922; }
.kpi.purple .kv { color:#bc8cff; }

.section-header {
    font-size: 14px;
    font-weight: 600;
    color: #e6edf3;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 1px solid #21262d;
}

div[data-testid="stVerticalBlock"] > div[style] { background: transparent; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CONSTANTS — expected column names
# ─────────────────────────────────────────────
COL_MAP = {
    "record_id":            "Record ID",
    "first_name":           "First Name",
    "last_name":            "Last Name",
    "owner":                "Contact owner",
    "create_date":          "Create Date",
    "last_activity":        "Last Activity Date",
    "last_contacted":       "Last Contacted",
    "call_response_status": "Call Response Status",
    "contact_stage":        "Contact Stage",
    "need_another_call":    "Need Another Call?",
    "lead_category":        "Lead Category",
    "disq_reason":          "Disqualification/Lost reason",
    "preferred_language":   "Preferred Language",
    "response_time":        "Lead response time (HH:mm:ss)",
    "product_interest":     "Product Interest",
    "primary_reason":       "Primary reason for enquiry",
    "purchase_timeline":    "Purchase Timeline",
    "ecg_volume":           "Estimated daily ECG Volume",
    "echo_volume":          "Estimated Daily ECHO Volumes",
    "decision_maker":       "Is Decision maker?",
    "current_workflow":     "Current Workflow",
    "num_locations":        "Number of Locations",
    "customer_segment":     "Customer Segment",
    "call_outcome":         "Call Outcome",
    "next_activity_date":   "Next Activity Date",
    "city":                 "City",
    "state":                "State/Region",
    "contact_source":       "Contact Source",
    "lifecycle_stage":      "Lifecycle Stage",
}

BASIC_FIELDS     = ["call_response_status", "contact_stage", "need_another_call", "lead_category", "disq_reason"]
EXTENDED_FIELDS  = ["product_interest", "primary_reason", "purchase_timeline", "ecg_volume",
                    "echo_volume", "decision_maker", "current_workflow", "num_locations",
                    "customer_segment", "call_outcome"]
BASIC_LABELS     = ["Call Response Status", "Contact Stage", "Need Another Call", "Lead Category", "Disq/Lost Reason"]
EXTENDED_LABELS  = ["Product Interest", "Primary Reason", "Purchase Timeline", "ECG Volume",
                    "ECHO Volume", "Decision Maker", "Current Workflow", "No. Locations",
                    "Customer Segment", "Call Outcome"]


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def read_file(uploaded) -> pd.DataFrame:
    """Read uploaded file (xlsx / xls / csv) into DataFrame."""
    if uploaded is None:
        return pd.DataFrame()
    name = uploaded.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(uploaded, low_memory=False)
        else:
            return pd.read_excel(uploaded, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read {uploaded.name}: {e}")
        return pd.DataFrame()


def safe_col(df: pd.DataFrame, key: str) -> pd.Series:
    """Return column by internal key, or empty Series if missing."""
    col = COL_MAP.get(key, key)
    if col in df.columns:
        return df[col]
    # fuzzy fallback — case-insensitive substring match
    for c in df.columns:
        if key.replace("_", " ").lower() in c.lower():
            return df[c]
    return pd.Series([None] * len(df), index=df.index)


def parse_rt(s) -> float | None:
    """Parse HH:MM:SS string → float minutes."""
    try:
        parts = str(s).strip().split(":")
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
        return None
    except Exception:
        return None


def median_minutes(series: pd.Series) -> float:
    vals = series.dropna().tolist()
    if not vals:
        return 0.0
    return float(np.median(vals))


def fmt_mins(m) -> str:
    if m is None or (isinstance(m, float) and np.isnan(m)):
        return "—"
    m = float(m)
    h, rem = divmod(m, 60)
    mins = int(rem)
    secs = int((rem % 1) * 60)
    if h >= 1:
        return f"{int(h)}h {mins}m"
    if mins > 0:
        return f"{mins}m {secs}s"
    return f"{secs}s"


def pct(num, den, decimals=1):
    if den == 0:
        return 0.0
    return round(num / den * 100, decimals)


def normalise(df: pd.DataFrame) -> pd.DataFrame:
    """Map raw columns → internal short names; add computed columns."""
    if df.empty:
        return df

    out = pd.DataFrame(index=df.index)

    for key, raw in COL_MAP.items():
        if raw in df.columns:
            out[key] = df[raw]
        else:
            out[key] = None

    # Also keep all original columns for pass-through in the full table
    for col in df.columns:
        if col not in out.columns:
            out["_raw_" + col] = df[col]

    # Dates
    for dc in ["create_date", "last_activity", "last_contacted", "next_activity_date"]:
        out[dc] = pd.to_datetime(out[dc], errors="coerce")

    # Full name
    fn = out["first_name"].fillna("").astype(str)
    ln = out["last_name"].fillna("").astype(str)
    out["name"] = (fn + " " + ln).str.strip().replace("nan", "").replace("nan nan", "")

    # Preferred language — merge two possible columns
    if "Preferred language" in df.columns and out["preferred_language"].isna().all():
        out["preferred_language"] = df["Preferred language"]

    # Response time in minutes
    out["rt_mins"] = out["response_time"].apply(parse_rt)

    # Missing counts
    out["missing_basic"] = out[BASIC_FIELDS].isnull().sum(axis=1)

    # MQL + responded flag
    out["is_mql_responded"] = (
        (out["lead_category"].fillna("").str.upper().eq("MQL")) &
        (out["call_response_status"].fillna("").str.strip().eq("Call Responded"))
    )
    out["missing_ext"] = out[EXTENDED_FIELDS].apply(
        lambda row: row.isnull().sum() if out.loc[row.name, "is_mql_responded"] else 0, axis=1
    )

    # Task flags
    out["has_task"] = out["next_activity_date"].notna()
    out["needs_task"] = (
        out["call_response_status"].fillna("").str.strip().eq("Call Not Responded") |
        out["need_another_call"].fillna("").str.strip().eq("Yes")
    )
    out["task_missing"] = out["needs_task"] & ~out["has_task"]

    return out


def business_hours_filter(df: pd.DataFrame, col: str = "create_date") -> pd.DataFrame:
    """Keep rows where col is between 09:00 and 21:00."""
    if df.empty or col not in df.columns:
        return df
    dt = pd.to_datetime(df[col], errors="coerce")
    mask = (dt.dt.hour >= 9) & (dt.dt.hour < 21)
    return df[mask.fillna(False)]


# ─────────────────────────────────────────────
# DATA HEALTH REPORT
# ─────────────────────────────────────────────
def data_health(df: pd.DataFrame, label: str):
    if df.empty:
        st.markdown(f'<div class="health-box health-err"><div class="health-title">❌ {label} — No data loaded</div></div>', unsafe_allow_html=True)
        return

    total = len(df)
    key_cols = list(COL_MAP.values())
    present  = [c for c in key_cols if c in df.columns]
    missing_cols = [c for c in key_cols if c not in df.columns]
    null_rates = {c: int(df[c].isna().sum()) for c in present}
    high_null  = {c: v for c, v in null_rates.items() if v > 0 and v / total > 0.5}

    score = 100
    score -= len(missing_cols) * 2
    score -= len(high_null) * 3
    score = max(0, score)

    level = "ok" if score >= 80 else "warn" if score >= 60 else "err"
    icon  = "✅" if level == "ok" else "⚠️" if level == "warn" else "❌"

    st.markdown(f"""
    <div class="health-box health-{level}">
      <div class="health-title">{icon} {label} — Health Score: {score}/100</div>
      <p style="font-size:12px;color:#8b949e;margin:0">
        {total} rows &nbsp;·&nbsp; {len(present)}/{len(key_cols)} key columns found
        {f" &nbsp;·&nbsp; <b style='color:#f85149'>{len(missing_cols)} columns missing</b>" if missing_cols else ""}
      </p>
    </div>
    """, unsafe_allow_html=True)

    if missing_cols or high_null:
        with st.expander(f"🔍 Details for {label}"):
            if missing_cols:
                st.markdown("**Missing columns** (dashboard will show '—' for these fields):")
                st.code(", ".join(missing_cols))
            if high_null:
                st.markdown("**Columns with >50% missing values:**")
                for c, v in high_null.items():
                    st.markdown(f"- `{c}`: {v}/{total} rows null ({pct(v, total):.1f}%)")
            st.markdown("**All columns found in file:**")
            st.code(", ".join(df.columns.tolist()))


# ─────────────────────────────────────────────
# PLOT HELPERS
# ─────────────────────────────────────────────
PLOTLY_THEME = dict(
    plot_bgcolor="#161b22",
    paper_bgcolor="#161b22",
    font=dict(color="#8b949e", family="DM Sans"),
    margin=dict(l=0, r=0, t=30, b=0),
)

def bar_h(labels, values, colors=None, title="", pct_of=None):
    if not labels:
        return go.Figure()
    if colors is None:
        colors = ["#58a6ff"] * len(labels)
    text = [f"{v}" if pct_of is None else f"{v} ({pct(v, pct_of):.1f}%)" for v in values]
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=colors,
        text=text, textposition="outside",
        textfont=dict(color="#e6edf3", size=11),
        hovertemplate="%{y}: %{x}<extra></extra>",
    ))
    fig.update_layout(**PLOTLY_THEME, title=dict(text=title, font=dict(color="#e6edf3", size=13)),
                      xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                      yaxis=dict(tickfont=dict(color="#e6edf3", size=11), autorange="reversed"),
                      height=max(180, len(labels) * 42))
    return fig


def donut(labels, values, colors=None, title=""):
    if not labels:
        return go.Figure()
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.55,
        marker_colors=colors or px.colors.qualitative.Set2,
        textinfo="percent+label",
        textfont=dict(color="#e6edf3", size=11),
        hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
    ))
    fig.update_layout(**PLOTLY_THEME, title=dict(text=title, font=dict(color="#e6edf3", size=13)),
                      showlegend=False, height=260)
    return fig


def progress_table(labels, filled, totals, extra_cols=None):
    """Render an HTML progress-bar table."""
    rows = ""
    for i, (lbl, f, tot) in enumerate(zip(labels, filled, totals)):
        p = pct(f, tot) if tot else 0
        color = "#3fb950" if p >= 90 else "#d29922" if p >= 70 else "#f85149"
        bar = f'<div style="background:#21262d;border-radius:4px;height:7px;width:100%"><div style="width:{p:.1f}%;background:{color};height:7px;border-radius:4px"></div></div>'
        extra = ""
        if extra_cols:
            extra = "".join(f'<td style="padding:8px 12px;color:#8b949e;font-size:12px">{extra_cols[j][i]}</td>' for j in range(len(extra_cols)))
        rows += f"""<tr>
            <td style="padding:8px 12px;color:#e6edf3;font-size:12px;white-space:nowrap">{lbl}</td>
            <td style="padding:8px 12px;width:100%">{bar}</td>
            <td style="padding:8px 12px;color:{color};font-size:12px;font-weight:600;white-space:nowrap">{f}/{tot} ({p:.1f}%)</td>
            {extra}
        </tr>"""
    return f"""
    <table style="width:100%;border-collapse:collapse;background:#161b22;border-radius:8px;overflow:hidden">
    <tbody>{rows}</tbody></table>"""


def style_df(df: pd.DataFrame) -> pd.DataFrame:
    """Return display-ready dataframe with nulls replaced."""
    return df.fillna("—")


# ─────────────────────────────────────────────
# SIDEBAR — FILE UPLOAD
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 IS Lead Dashboard")
    st.markdown("---")
    st.markdown("### 📂 Upload Files")

    file1 = st.file_uploader(
        "**File 1 — New Leads** (created yesterday)",
        type=["xlsx", "xls", "csv"],
        key="f1",
        help="Leads created yesterday — used for response time, raw stage, new lead metrics",
    )
    file2 = st.file_uploader(
        "**File 2 — Activity Leads** (worked yesterday)",
        type=["xlsx", "xls", "csv"],
        key="f2",
        help="Leads that had activity done on them yesterday",
    )

    st.markdown("---")
    st.markdown("### ⚙️ Filters")
    biz_start = st.slider("Business Hours Start", 0, 23, 9, key="biz_start")
    biz_end   = st.slider("Business Hours End",   0, 23, 21, key="biz_end")
    st.caption("Used for response time & raw stage calculations")

    st.markdown("---")
    st.markdown("### 📋 About")
    st.caption(
        "Drag & drop any Excel or CSV file. "
        "The dashboard handles missing columns and null values automatically. "
        "Metrics update instantly on upload."
    )


# ─────────────────────────────────────────────
# LOAD & PROCESS DATA
# ─────────────────────────────────────────────
raw1 = read_file(file1)
raw2 = read_file(file2)

df_new = normalise(raw1) if not raw1.empty else pd.DataFrame()
df_act = normalise(raw2) if not raw2.empty else pd.DataFrame()

# Business-hours subset of new leads
df_new_biz = pd.DataFrame()
if not df_new.empty:
    mask_biz = (
        df_new["create_date"].dt.hour.between(biz_start, biz_end - 1, inclusive="both")
    )
    df_new_biz = df_new[mask_biz.fillna(False)].copy()

# Combined
frames = [df for df in [df_new, df_act] if not df.empty]
df_all = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

no_data = df_all.empty


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col_h1, col_h2 = st.columns([4, 1])
with col_h1:
    st.markdown("# 📊 IS Lead Dashboard")
    if not df_new.empty or not df_act.empty:
        dates = []
        for df, lbl in [(df_new, "New Leads"), (df_act, "Activity Leads")]:
            if not df.empty and df["create_date"].notna().any():
                d = df["create_date"].dropna().dt.date
                dates.append(f"**{lbl}**: {d.min()} → {d.max()}")
        st.caption("  |  ".join(dates) if dates else "Data loaded")
    else:
        st.caption("Upload files from the sidebar to begin →")

with col_h2:
    if not df_all.empty:
        st.metric("Total Leads", len(df_all))


# ─────────────────────────────────────────────
# DATA HEALTH SECTION (always visible)
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("### 🩺 Data Health")
hcol1, hcol2 = st.columns(2)
with hcol1:
    data_health(raw1, "File 1 — New Leads")
with hcol2:
    data_health(raw2, "File 2 — Activity Leads")


if no_data:
    st.info("⬆️ Upload at least one file from the sidebar to see the full dashboard.")
    st.stop()


# ─────────────────────────────────────────────
# TAB LAYOUT
# ─────────────────────────────────────────────
tab_labels = [
    "🏠 Overview",
    "⏱️ Response Time",
    "📞 Call Status",
    "📋 Field Completion",
    "✅ Task Audit",
    "🏷️ Lead Qualification",
    "👤 By Agent",
    "📑 All Leads",
]
tabs = st.tabs(tab_labels)


# ═════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ═════════════════════════════════════════════
with tabs[0]:
    nl = df_new
    al = df_act
    nlb = df_new_biz
    comb = df_all

    n_new     = len(nl)
    n_act     = len(al)
    n_biz     = len(nlb)
    n_resp_nl = int((nl["call_response_status"].fillna("").str.strip() == "Call Responded").sum()) if not nl.empty else 0
    n_notresp_nl = int((nl["call_response_status"].fillna("").str.strip() == "Call Not Responded").sum()) if not nl.empty else 0
    n_invalid_nl  = int((nl["call_response_status"].fillna("").str.strip() == "Invalid Number").sum()) if not nl.empty else 0

    rt_vals = nlb["rt_mins"].dropna().tolist() if not nlb.empty else []
    avg_rt  = float(np.mean(rt_vals)) if rt_vals else None
    med_rt  = float(np.median(rt_vals)) if rt_vals else None

    n_mql = int((comb["lead_category"].fillna("").str.upper() == "MQL").sum()) if not comb.empty else 0
    n_mul = int((comb["lead_category"].fillna("").str.upper() == "MUL").sum()) if not comb.empty else 0
    n_task_miss = int(comb["task_missing"].sum()) if not comb.empty else 0

    # KPI row
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("New Leads", n_new, f"{n_biz} in {biz_start}–{biz_end}h")
    c2.metric("Activity Leads", n_act)
    c3.metric("Responded (New)", n_resp_nl, f"{pct(n_resp_nl, n_new):.1f}%")
    c4.metric("Avg Response Time", fmt_mins(avg_rt), f"Median {fmt_mins(med_rt)}")
    c5.metric("MQL / MUL", f"{n_mql} / {n_mul}")
    c6.metric("Tasks Missing ⚠️", n_task_miss, delta_color="inverse")

    st.markdown("---")

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown('<div class="section-header">📞 Call Response — New Leads</div>', unsafe_allow_html=True)
        if not nl.empty:
            vc = nl["call_response_status"].fillna("").str.strip().value_counts()
            colors_map = {
                "Call Responded": "#3fb950",
                "Call Not Responded": "#f85149",
                "Invalid Number": "#d29922",
            }
            lbls = vc.index.tolist()
            vals = vc.values.tolist()
            clrs = [colors_map.get(l, "#58a6ff") for l in lbls]
            st.plotly_chart(donut(lbls, vals, clrs), use_container_width=True)
        else:
            st.info("No new leads data.")

    with col_b:
        st.markdown('<div class="section-header">🏷️ Lead Category — All Leads</div>', unsafe_allow_html=True)
        if not comb.empty:
            cat = comb["lead_category"].fillna("Not Set").str.strip().value_counts()
            clrs2 = {"MQL": "#58a6ff", "MUL": "#f85149", "Not Set": "#484f58"}
            st.plotly_chart(
                donut(cat.index.tolist(), cat.values.tolist(),
                      [clrs2.get(l, "#8b949e") for l in cat.index]),
                use_container_width=True)
        else:
            st.info("No data.")

    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown('<div class="section-header">📋 Contact Stage — New Leads</div>', unsafe_allow_html=True)
        if not nl.empty:
            stg = nl["contact_stage"].fillna("Not Set").str.strip().value_counts()
            stage_colors = {
                "Lost": "#f85149", "Qualification": "#58a6ff", "Opportunity": "#3fb950",
                "Pass to field": "#bc8cff", "Raw Lead": "#d29922",
                "Zigment Unqualified": "#484f58", "Not Set": "#30363d",
                "Contact in Future": "#d29922", "Cold": "#484f58",
            }
            clrs3 = [stage_colors.get(l, "#8b949e") for l in stg.index]
            st.plotly_chart(bar_h(stg.index.tolist(), stg.values.tolist(), clrs3, pct_of=n_new), use_container_width=True)
        else:
            st.info("No new leads data.")

    with col_d:
        st.markdown('<div class="section-header">⚠️ Key Alerts</div>', unsafe_allow_html=True)
        if not comb.empty:
            nl_resp = nl[nl["call_response_status"].fillna("").str.strip() == "Call Responded"] if not nl.empty else pd.DataFrame()
            lang_miss  = int((nl_resp["preferred_language"].isna()).sum()) if not nl_resp.empty else 0
            cat_miss   = int((nl_resp["lead_category"].isna()).sum()) if not nl_resp.empty else 0
            mul_no_disq = int(((comb["lead_category"].fillna("").str.upper() == "MUL") & comb["disq_reason"].isna()).sum())
            raw_count  = int((nlb["contact_stage"].isna() | (nlb["contact_stage"].fillna("").str.strip() == "Raw Lead")).sum()) if not nlb.empty else 0

            def alert_md(level, icon, msg):
                bg = {"ok": "#0d1117", "warn": "#1a1500", "err": "#1c0a0a"}[level]
                border = {"ok": "#3fb950", "warn": "#d29922", "err": "#f85149"}[level]
                color  = {"ok": "#3fb950", "warn": "#d29922", "err": "#f85149"}[level]
                return f"""<div style="background:{bg};border:1px solid {border};border-left:4px solid {border};
                    border-radius:8px;padding:10px 14px;margin-bottom:8px;font-size:13px;color:{color}">
                    {icon} {msg}</div>"""

            alerts_html = ""
            alerts_html += alert_md("err" if lang_miss > 0 else "ok", "🌐",
                f"{lang_miss} responded calls missing preferred language ({pct(lang_miss, max(len(nl_resp),1)):.1f}%)")
            alerts_html += alert_md("err" if cat_miss > 0 else "ok", "🏷️",
                f"{cat_miss} responded calls missing lead category")
            alerts_html += alert_md("warn" if mul_no_disq > 0 else "ok", "❌",
                f"{mul_no_disq} MUL leads missing disqualification reason")
            alerts_html += alert_md("err" if n_task_miss > 0 else "ok", "📅",
                f"{n_task_miss} leads requiring tasks have NO task created")
            alerts_html += alert_md("warn" if raw_count > 0 else "ok", "🔴",
                f"{raw_count} leads in Raw/unworked stage ({biz_start}–{biz_end}h window)")

            st.markdown(alerts_html, unsafe_allow_html=True)
        else:
            st.info("No data for alerts.")


# ═════════════════════════════════════════════
# TAB 2 — RESPONSE TIME
# ═════════════════════════════════════════════
with tabs[1]:
    st.markdown("### ⏱️ Response Time Analysis")
    st.caption(f"Based on new leads created between {biz_start}:00–{biz_end}:00 only")

    if df_new_biz.empty:
        st.warning("No new leads in the selected business-hours window. Adjust the slider in the sidebar.")
    else:
        rt_series = df_new_biz["rt_mins"].dropna()
        avg_rt_v  = float(np.mean(rt_series)) if len(rt_series) else None
        med_rt_v  = float(np.median(rt_series)) if len(rt_series) else None
        under5    = int((rt_series < 5).sum())
        b5_15     = int(((rt_series >= 5) & (rt_series <= 15)).sum())
        over15    = int((rt_series > 15).sum())

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Leads with RT Data", len(rt_series))
        c2.metric("Average RT", fmt_mins(avg_rt_v))
        c3.metric("Median RT", fmt_mins(med_rt_v))
        c4.metric("Under 5 min ✅", under5, f"{pct(under5, len(rt_series)):.1f}%")
        c5.metric("Over 15 min ⚠️", over15, f"{pct(over15, len(rt_series)):.1f}%", delta_color="inverse")

        st.markdown("---")
        col_rt1, col_rt2 = st.columns(2)

        with col_rt1:
            st.markdown('<div class="section-header">By Agent — Avg & Median RT</div>', unsafe_allow_html=True)
            agent_rt = (
                df_new_biz.dropna(subset=["rt_mins"])
                .groupby("owner")["rt_mins"]
                .agg(["mean", "median", "count"])
                .reset_index()
                .sort_values("mean")
            )
            if not agent_rt.empty:
                agent_rt.columns = ["Agent", "Avg RT (min)", "Median RT (min)", "Count"]
                colors_rt = [
                    "#3fb950" if v < 5 else "#d29922" if v < 10 else "#f85149"
                    for v in agent_rt["Avg RT (min)"]
                ]
                fig_rt = go.Figure()
                fig_rt.add_trace(go.Bar(
                    y=agent_rt["Agent"],
                    x=agent_rt["Avg RT (min)"],
                    name="Avg RT",
                    orientation="h",
                    marker_color=colors_rt,
                    text=[fmt_mins(v) for v in agent_rt["Avg RT (min)"]],
                    textposition="outside",
                    textfont=dict(color="#e6edf3", size=10),
                    hovertemplate="%{y}<br>Avg: %{x:.1f}m<extra></extra>",
                ))
                fig_rt.add_trace(go.Scatter(
                    y=agent_rt["Agent"],
                    x=agent_rt["Median RT (min)"],
                    name="Median RT",
                    mode="markers",
                    marker=dict(color="#bc8cff", size=10, symbol="diamond"),
                    hovertemplate="%{y}<br>Median: %{x:.1f}m<extra></extra>",
                ))
                fig_rt.update_layout(
                    **PLOTLY_THEME,
                    barmode="overlay",
                    xaxis=dict(showgrid=True, gridcolor="#21262d", zeroline=False),
                    yaxis=dict(tickfont=dict(color="#e6edf3"), autorange="reversed"),
                    legend=dict(
    orientation="h",
    y=1.05,
    x=0,
    bgcolor="#161b22",
    bordercolor="#21262d"
),
                    height=max(250, len(agent_rt) * 50),
                )
                st.plotly_chart(fig_rt, use_container_width=True)

                # Table
                display_rt = agent_rt.copy()
                display_rt["Avg RT (min)"] = display_rt["Avg RT (min)"].apply(fmt_mins)
                display_rt["Median RT (min)"] = display_rt["Median RT (min)"].apply(fmt_mins)
                st.dataframe(display_rt, use_container_width=True, hide_index=True)

        with col_rt2:
            st.markdown('<div class="section-header">Distribution Buckets</div>', unsafe_allow_html=True)
            buckets = {
                "< 2 min":   int((rt_series < 2).sum()),
                "2 – 5 min": int(((rt_series >= 2) & (rt_series < 5)).sum()),
                "5 – 10 min":int(((rt_series >= 5) & (rt_series < 10)).sum()),
                "10 – 15 min":int(((rt_series >= 10) & (rt_series < 15)).sum()),
                "15 – 30 min":int(((rt_series >= 15) & (rt_series < 30)).sum()),
                "> 30 min":  int((rt_series >= 30).sum()),
            }
            b_labels = list(buckets.keys())
            b_values = list(buckets.values())
            b_colors = ["#3fb950","#3fb950","#d29922","#d29922","#f85149","#f85149"]
            st.plotly_chart(bar_h(b_labels, b_values, b_colors, pct_of=len(rt_series)), use_container_width=True)

            # Histogram
            st.markdown('<div class="section-header" style="margin-top:12px">RT Histogram (minutes)</div>', unsafe_allow_html=True)
            fig_hist = px.histogram(
                rt_series, nbins=20,
                color_discrete_sequence=["#58a6ff"],
                labels={"value": "Minutes"},
            )
            fig_hist.update_layout(**PLOTLY_THEME, height=200,
                                   xaxis=dict(gridcolor="#21262d"),
                                   yaxis=dict(gridcolor="#21262d"))
            st.plotly_chart(fig_hist, use_container_width=True)

        # Detail table
        st.markdown("---")
        st.markdown('<div class="section-header">Individual Response Times</div>', unsafe_allow_html=True)
        rt_cols = ["name", "owner", "create_date", "call_response_status", "rt_mins", "contact_stage"]
        rt_disp = df_new_biz[[c for c in rt_cols if c in df_new_biz.columns]].copy()
        rt_disp["create_date"] = rt_disp["create_date"].dt.strftime("%H:%M")
        rt_disp["rt_mins"] = rt_disp["rt_mins"].apply(fmt_mins)
        rt_disp.columns = [c.replace("_", " ").title() for c in rt_disp.columns]
        search_rt = st.text_input("Search agent/name", key="rt_search")
        if search_rt:
            mask = rt_disp.apply(lambda r: r.astype(str).str.contains(search_rt, case=False).any(), axis=1)
            rt_disp = rt_disp[mask]
        st.dataframe(style_df(rt_disp), use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════
# TAB 3 — CALL STATUS
# ═════════════════════════════════════════════
with tabs[2]:
    st.markdown("### 📞 Call Status & Preferred Language")

    if df_all.empty:
        st.warning("No data loaded.")
    else:
        nl_r = df_new[df_new["call_response_status"].fillna("").str.strip() == "Call Responded"] if not df_new.empty else pd.DataFrame()
        al_r = df_act[df_act["call_response_status"].fillna("").str.strip() == "Call Responded"] if not df_act.empty else pd.DataFrame()
        all_resp = df_all[df_all["call_response_status"].fillna("").str.strip() == "Call Responded"] if not df_all.empty else pd.DataFrame()
        lang_miss = int(all_resp["preferred_language"].isna().sum()) if not all_resp.empty else 0
        lang_fill = len(all_resp) - lang_miss

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Responded — New",      len(nl_r), f"{pct(len(nl_r), len(df_new)):.1f}%")
        c2.metric("Responded — Activity", len(al_r), f"{pct(len(al_r), len(df_act)):.1f}%")
        c3.metric("Total Responded",      len(all_resp), f"{pct(len(all_resp), len(df_all)):.1f}%")
        c4.metric("Language Filled ✅",   lang_fill, f"{pct(lang_fill, max(len(all_resp),1)):.1f}%")
        c5.metric("Language Missing ⚠️",  lang_miss, f"{pct(lang_miss, max(len(all_resp),1)):.1f}%", delta_color="inverse")

        st.markdown("---")
        col_cs1, col_cs2 = st.columns(2)
        status_order = ["Call Responded", "Call Not Responded", "Invalid Number"]
        s_colors     = ["#3fb950", "#f85149", "#d29922"]

        with col_cs1:
            st.markdown('<div class="section-header">New Leads Call Status</div>', unsafe_allow_html=True)
            if not df_new.empty:
                vc_nl = df_new["call_response_status"].fillna("").str.strip().value_counts().reindex(status_order, fill_value=0)
                st.plotly_chart(bar_h(status_order, vc_nl.values.tolist(), s_colors, pct_of=len(df_new)), use_container_width=True)

        with col_cs2:
            st.markdown('<div class="section-header">Activity Leads Call Status</div>', unsafe_allow_html=True)
            if not df_act.empty:
                vc_al = df_act["call_response_status"].fillna("").str.strip().value_counts().reindex(status_order, fill_value=0)
                st.plotly_chart(bar_h(status_order, vc_al.values.tolist(), s_colors, pct_of=len(df_act)), use_container_width=True)

        # Language by agent
        st.markdown("---")
        st.markdown('<div class="section-header">🌐 Preferred Language Completion — Responded Calls by Agent</div>', unsafe_allow_html=True)
        if not all_resp.empty:
            lang_agent = (
                all_resp.assign(has_lang=all_resp["preferred_language"].notna())
                .groupby("owner")
                .agg(filled=("has_lang", "sum"), total=("has_lang", "count"))
                .reset_index()
                .sort_values("filled", ascending=False)
            )
            filled_list = lang_agent["filled"].astype(int).tolist()
            total_list  = lang_agent["total"].astype(int).tolist()
            st.markdown(
                progress_table(lang_agent["owner"].tolist(), filled_list, total_list),
                unsafe_allow_html=True,
            )

        # Responded detail table
        st.markdown("---")
        st.markdown('<div class="section-header">Responded Calls — Detail Table</div>', unsafe_allow_html=True)
        src_opts = {"All": "all", "New Leads": "new", "Activity Leads": "act"}
        lang_opts = {"All": "all", "Language Filled": "filled", "Language Missing": "missing"}
        fc1, fc2, fc3 = st.columns(3)
        cs_src  = fc1.selectbox("Source", list(src_opts.keys()), key="cs_src")
        cs_lang = fc2.selectbox("Language", list(lang_opts.keys()), key="cs_lang")
        cs_srch = fc3.text_input("Search", key="cs_srch")

        base_cs = {"all": df_all, "new": df_new, "act": df_act}[src_opts[cs_src]]
        if not base_cs.empty:
            base_cs = base_cs[base_cs["call_response_status"].fillna("").str.strip() == "Call Responded"]
            if lang_opts[cs_lang] == "filled":
                base_cs = base_cs[base_cs["preferred_language"].notna()]
            elif lang_opts[cs_lang] == "missing":
                base_cs = base_cs[base_cs["preferred_language"].isna()]
            if cs_srch:
                mask = base_cs.apply(lambda r: r.astype(str).str.contains(cs_srch, case=False).any(), axis=1)
                base_cs = base_cs[mask]
            show_cs = base_cs[["name","owner","call_response_status","preferred_language","lead_category","contact_stage"]].copy()
            st.dataframe(style_df(show_cs), use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════
# TAB 4 — FIELD COMPLETION
# ═════════════════════════════════════════════
with tabs[3]:
    st.markdown("### 📋 Field Completion Audit")

    sub_fc = st.radio("View", ["Basic Fields (All Leads)", "Extended Fields (MQL + Responded)"],
                      horizontal=True, key="fc_sub")

    if df_all.empty:
        st.warning("No data.")
    elif sub_fc.startswith("Basic"):
        total_fc = len(df_all)
        filled_b = [int(df_all[f].notna().sum()) for f in BASIC_FIELDS]
        missing_b = [total_fc - f for f in filled_b]

        # KPIs
        cols_b = st.columns(len(BASIC_FIELDS))
        for i, (lbl, fill) in enumerate(zip(BASIC_LABELS, filled_b)):
            p = pct(fill, total_fc)
            cols_b[i].metric(lbl, f"{p:.1f}%", f"{total_fc-fill} missing")

        st.markdown("---")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.markdown('<div class="section-header">Overall Completion</div>', unsafe_allow_html=True)
            st.markdown(progress_table(BASIC_LABELS, filled_b, [total_fc]*len(BASIC_LABELS)), unsafe_allow_html=True)

        with col_f2:
            field_sel = st.selectbox("Field to drill by agent", BASIC_LABELS, key="ff_field")
            f_key = BASIC_FIELDS[BASIC_LABELS.index(field_sel)]
            agent_fill = (
                df_all.assign(has_val=df_all[f_key].notna())
                .groupby("owner")
                .agg(filled=("has_val","sum"), total=("has_val","count"))
                .reset_index().sort_values("filled", ascending=False)
            )
            st.markdown('<div class="section-header">By Agent</div>', unsafe_allow_html=True)
            st.markdown(
                progress_table(agent_fill["owner"].tolist(),
                               agent_fill["filled"].astype(int).tolist(),
                               agent_fill["total"].astype(int).tolist()),
                unsafe_allow_html=True,
            )

        # Lead-level table
        st.markdown("---")
        st.markdown('<div class="section-header">Lead-Level Audit</div>', unsafe_allow_html=True)
        f_src = st.selectbox("Source", ["All","New Leads","Activity Leads"], key="ff_src")
        f_miss = st.selectbox("Filter", ["Show All","Has Missing Fields","Complete"], key="ff_miss")
        f_owner = st.selectbox("Agent", ["All"] + sorted(df_all["owner"].dropna().unique().tolist()), key="ff_owner")
        f_srch = st.text_input("Search", key="ff_srch")

        base_ff = {"All": df_all, "New Leads": df_new, "Activity Leads": df_act}.get(f_src, df_all)
        if not base_ff.empty:
            if f_miss == "Has Missing Fields":
                base_ff = base_ff[base_ff["missing_basic"] > 0]
            elif f_miss == "Complete":
                base_ff = base_ff[base_ff["missing_basic"] == 0]
            if f_owner != "All":
                base_ff = base_ff[base_ff["owner"] == f_owner]
            if f_srch:
                mask = base_ff.apply(lambda r: r.astype(str).str.contains(f_srch, case=False).any(), axis=1)
                base_ff = base_ff[mask]
            cols_show = ["name","owner","call_response_status","contact_stage","need_another_call",
                         "lead_category","disq_reason","missing_basic"]
            st.markdown(f"**{len(base_ff)} leads**")
            st.dataframe(style_df(base_ff[[c for c in cols_show if c in base_ff.columns]]),
                         use_container_width=True, hide_index=True)

    else:  # Extended (MQL + Responded)
        mql_resp = df_all[df_all["is_mql_responded"]] if not df_all.empty else pd.DataFrame()
        st.metric("MQL + Responded Leads", len(mql_resp))

        if mql_resp.empty:
            st.info("No MQL + Responded leads found in the data.")
        else:
            total_mr = len(mql_resp)
            filled_e = [int(mql_resp[f].notna().sum()) for f in EXTENDED_FIELDS]
            st.markdown("---")
            st.markdown('<div class="section-header">Extended Field Completion</div>', unsafe_allow_html=True)
            st.markdown(progress_table(EXTENDED_LABELS, filled_e, [total_mr]*len(EXTENDED_LABELS)), unsafe_allow_html=True)

            # Detail table
            st.markdown("---")
            cols_mql = ["name","owner"] + [f for f in EXTENDED_FIELDS if f in mql_resp.columns] + ["missing_ext"]
            st.dataframe(style_df(mql_resp[[c for c in cols_mql if c in mql_resp.columns]]),
                         use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════
# TAB 5 — TASK AUDIT
# ═════════════════════════════════════════════
with tabs[4]:
    st.markdown("### ✅ Task Creation Audit")
    st.caption("Leads that are 'Call Not Responded' OR 'Need Another Call: Yes' should have a task (Next Activity Date)")

    if df_all.empty:
        st.warning("No data.")
    else:
        needs = df_all[df_all["needs_task"]]
        has_t = needs[needs["has_task"]]
        miss_t = needs[~needs["has_task"]]
        nac_leads = df_all[df_all["need_another_call"].fillna("").str.strip().eq("Yes")]
        nr_leads  = df_all[df_all["call_response_status"].fillna("").str.strip().eq("Call Not Responded")]

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Leads Requiring Task",   len(needs))
        c2.metric("Task Created ✅",         len(has_t),  f"{pct(len(has_t), len(needs)):.1f}%")
        c3.metric("Task Missing ⚠️",         len(miss_t), f"{pct(len(miss_t), len(needs)):.1f}%", delta_color="inverse")
        c4.metric("Call Not Responded",      len(nr_leads))
        c5.metric("Need Another Call: Yes",  len(nac_leads))

        st.markdown("---")
        col_t1, col_t2 = st.columns(2)

        def task_pie(df_subset, title):
            t = len(df_subset)
            h = int(df_subset["has_task"].sum())
            m = t - h
            return donut(["Has Task","Missing Task"],[h,m],["#3fb950","#f85149"], title)

        with col_t1:
            st.markdown('<div class="section-header">Call Not Responded — Task Status</div>', unsafe_allow_html=True)
            st.plotly_chart(task_pie(nr_leads, ""), use_container_width=True)

        with col_t2:
            st.markdown('<div class="section-header">Need Another Call: Yes — Task Status</div>', unsafe_allow_html=True)
            st.plotly_chart(task_pie(nac_leads, ""), use_container_width=True)

        # By agent
        st.markdown("---")
        st.markdown('<div class="section-header">Task Gap by Agent</div>', unsafe_allow_html=True)
        ag_task = (
            needs.groupby("owner")
            .agg(required=("has_task","count"), created=("has_task","sum"))
            .reset_index()
        )
        ag_task["missing"] = ag_task["required"] - ag_task["created"]
        ag_task = ag_task.sort_values("missing", ascending=False)
        fig_task = go.Figure()
        fig_task.add_trace(go.Bar(y=ag_task["owner"], x=ag_task["created"], name="Created",
                                  orientation="h", marker_color="#3fb950"))
        fig_task.add_trace(go.Bar(y=ag_task["owner"], x=ag_task["missing"], name="Missing",
                                  orientation="h", marker_color="#f85149"))
        fig_task.update_layout(**PLOTLY_THEME, barmode="stack",
                                xaxis=dict(gridcolor="#21262d"),
                                yaxis=dict(tickfont=dict(color="#e6edf3"), autorange="reversed"),
                                height=max(200, len(ag_task)*50))
        st.plotly_chart(fig_task, use_container_width=True)

        # Detail table
        st.markdown("---")
        st.markdown('<div class="section-header">Leads Requiring Task — Detail</div>', unsafe_allow_html=True)
        t_reason = st.selectbox("Filter by reason", ["All","Call Not Responded","Need Another Call: Yes"], key="task_reason")
        t_owner  = st.selectbox("Filter by agent",  ["All"] + sorted(df_all["owner"].dropna().unique().tolist()), key="task_owner")
        t_status = st.selectbox("Task status", ["All","Missing","Created"], key="task_status")
        t_srch   = st.text_input("Search", key="task_srch")

        base_task = needs.copy()
        if t_reason == "Call Not Responded":
            base_task = base_task[base_task["call_response_status"].fillna("").str.strip() == "Call Not Responded"]
        elif t_reason == "Need Another Call: Yes":
            base_task = base_task[base_task["need_another_call"].fillna("").str.strip() == "Yes"]
        if t_owner != "All":
            base_task = base_task[base_task["owner"] == t_owner]
        if t_status == "Missing":
            base_task = base_task[~base_task["has_task"]]
        elif t_status == "Created":
            base_task = base_task[base_task["has_task"]]
        if t_srch:
            mask = base_task.apply(lambda r: r.astype(str).str.contains(t_srch, case=False).any(), axis=1)
            base_task = base_task[mask]

        show_task = base_task[["name","owner","call_response_status","need_another_call",
                               "next_activity_date","has_task","contact_stage"]].copy()
        show_task["next_activity_date"] = show_task["next_activity_date"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
        show_task["has_task"] = show_task["has_task"].map({True:"✅ Yes", False:"❌ Missing"})
        st.markdown(f"**{len(show_task)} leads**")
        st.dataframe(style_df(show_task), use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════
# TAB 6 — LEAD QUALIFICATION
# ═════════════════════════════════════════════
with tabs[5]:
    st.markdown("### 🏷️ Lead Qualification")

    if df_all.empty:
        st.warning("No data.")
    else:
        mql = df_all[df_all["lead_category"].fillna("").str.upper().eq("MQL")] if "lead_category" in df_all.columns else pd.DataFrame()
        mul = df_all[df_all["lead_category"].fillna("").str.upper().eq("MUL")] if "lead_category" in df_all.columns else pd.DataFrame()
        not_cat = df_all[df_all["lead_category"].isna()]
        mul_no_disq = mul[mul["disq_reason"].isna()] if not mul.empty else pd.DataFrame()
        resp_all = df_all[df_all["call_response_status"].fillna("").str.strip() == "Call Responded"]
        resp_no_cat = resp_all[resp_all["lead_category"].isna()]
        raw_leads = df_new_biz[df_new_biz["contact_stage"].isna() | df_new_biz["contact_stage"].fillna("").str.strip().eq("Raw Lead")] if not df_new_biz.empty else pd.DataFrame()

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("MQL", len(mql), f"{pct(len(mql),len(df_all)):.1f}%")
        c2.metric("MUL", len(mul), f"{pct(len(mul),len(df_all)):.1f}%")
        c3.metric("Not Categorized", len(not_cat), f"{pct(len(not_cat),len(df_all)):.1f}%")
        c4.metric("MUL Missing Disq ⚠️", len(mul_no_disq), f"{pct(len(mul_no_disq),max(len(mul),1)):.1f}% of MUL", delta_color="inverse")
        c5.metric("Responded → No Cat ⚠️", len(resp_no_cat), f"{pct(len(resp_no_cat),max(len(resp_all),1)):.1f}%", delta_color="inverse")
        c6.metric(f"Raw Stage ({biz_start}–{biz_end}h)", len(raw_leads))

        st.markdown("---")
        col_lq1, col_lq2 = st.columns(2)

        with col_lq1:
            st.markdown('<div class="section-header">MQL vs MUL — All Leads</div>', unsafe_allow_html=True)
            cat_vc = df_all["lead_category"].fillna("Not Set").str.strip().value_counts()
            st.plotly_chart(donut(cat_vc.index.tolist(), cat_vc.values.tolist(),
                                  [{"MQL":"#58a6ff","MUL":"#f85149","Not Set":"#484f58"}.get(l,"#8b949e") for l in cat_vc.index]),
                            use_container_width=True)

        with col_lq2:
            st.markdown('<div class="section-header">MUL Disqualification Reasons</div>', unsafe_allow_html=True)
            if not mul.empty:
                disq_vc = mul["disq_reason"].fillna("NOT SET").str.strip().value_counts()
                d_colors = ["#f85149" if l == "NOT SET" else "#58a6ff" for l in disq_vc.index]
                st.plotly_chart(bar_h(disq_vc.index.tolist(), disq_vc.values.tolist(), d_colors, pct_of=len(mul)), use_container_width=True)
            else:
                st.info("No MUL leads.")

        # Lead category by agent for responded calls
        st.markdown("---")
        st.markdown('<div class="section-header">Lead Category Completion (Responded Calls) — by Agent</div>', unsafe_allow_html=True)
        if not resp_all.empty:
            ag_cat = (
                resp_all.assign(has_cat=resp_all["lead_category"].notna())
                .groupby("owner")
                .agg(filled=("has_cat","sum"), total=("has_cat","count"))
                .reset_index().sort_values("filled", ascending=False)
            )
            st.markdown(
                progress_table(ag_cat["owner"].tolist(),
                               ag_cat["filled"].astype(int).tolist(),
                               ag_cat["total"].astype(int).tolist()),
                unsafe_allow_html=True,
            )

        # Raw stage info
        st.markdown("---")
        if len(raw_leads) == 0:
            st.success(f"✅ No leads in Raw/unworked stage during {biz_start}:00–{biz_end}:00 window.")
        else:
            st.warning(f"⚠️ {len(raw_leads)} leads still in Raw stage during {biz_start}:00–{biz_end}:00 window.")
            st.dataframe(style_df(raw_leads[["name","owner","create_date","contact_stage"]]),
                         use_container_width=True, hide_index=True)

        # Detail table
        st.markdown("---")
        st.markdown('<div class="section-header">Qualification Detail Table</div>', unsafe_allow_html=True)
        lq_fc1, lq_fc2, lq_fc3, lq_fc4 = st.columns(4)
        lq_src  = lq_fc1.selectbox("Source",   ["All","New Leads","Activity Leads"], key="lq_src")
        lq_cat  = lq_fc2.selectbox("Category", ["All","MQL","MUL","Not Set"], key="lq_cat")
        lq_disq = lq_fc3.selectbox("Disq Reason", ["All","Missing","Filled"], key="lq_disq")
        lq_srch = lq_fc4.text_input("Search", key="lq_srch")

        base_lq = {"All":df_all,"New Leads":df_new,"Activity Leads":df_act}.get(lq_src, df_all)
        if not base_lq.empty:
            if lq_cat == "MQL": base_lq = base_lq[base_lq["lead_category"].fillna("").str.upper().eq("MQL")]
            elif lq_cat == "MUL": base_lq = base_lq[base_lq["lead_category"].fillna("").str.upper().eq("MUL")]
            elif lq_cat == "Not Set": base_lq = base_lq[base_lq["lead_category"].isna()]
            if lq_disq == "Missing": base_lq = base_lq[base_lq["disq_reason"].isna()]
            elif lq_disq == "Filled": base_lq = base_lq[base_lq["disq_reason"].notna()]
            if lq_srch:
                mask = base_lq.apply(lambda r: r.astype(str).str.contains(lq_srch, case=False).any(), axis=1)
                base_lq = base_lq[mask]
            show_lq = base_lq[["name","owner","call_response_status","lead_category","disq_reason","contact_stage","call_outcome"]].copy()
            st.markdown(f"**{len(show_lq)} leads**")
            st.dataframe(style_df(show_lq), use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════
# TAB 7 — BY AGENT
# ═════════════════════════════════════════════
with tabs[6]:
    st.markdown("### 👤 Agent Performance")

    if df_all.empty:
        st.warning("No data.")
    else:
        ag_src = st.selectbox("Data Source", ["All","New Leads","Activity Leads"], key="ag_src")
        ag_base = {"All":df_all,"New Leads":df_new,"Activity Leads":df_act}.get(ag_src, df_all)

        owners = sorted(ag_base["owner"].dropna().unique().tolist())

        rows = []
        for owner in owners:
            ol  = ag_base[ag_base["owner"] == owner]
            rt_v = ol["rt_mins"].dropna()
            resp  = ol[ol["call_response_status"].fillna("").str.strip() == "Call Responded"]
            mql_c = int((ol["lead_category"].fillna("").str.upper() == "MQL").sum())
            mul_c = int((ol["lead_category"].fillna("").str.upper() == "MUL").sum())
            lang_m= int(resp["preferred_language"].isna().sum())
            task_m= int(ol["task_missing"].sum())
            basic_fill = int(((len(BASIC_FIELDS)*len(ol)) - ol["missing_basic"].sum())) / max(len(BASIC_FIELDS)*len(ol),1) * 100

            rows.append({
                "Agent": owner,
                "Total Leads": len(ol),
                "Responded": len(resp),
                "Responded %": f"{pct(len(resp),len(ol)):.1f}%",
                "MQL": mql_c,
                "MUL": mul_c,
                "Lang Missing": lang_m,
                "Task Missing": task_m,
                "Avg RT": fmt_mins(rt_v.mean() if len(rt_v) else None),
                "Avg RT (min)": round(float(rt_v.mean()), 2) if len(rt_v) else 9999,
                "Basic Fill %": f"{basic_fill:.1f}%",
                "Basic Fill": round(basic_fill, 1),
            })

        df_ag = pd.DataFrame(rows)

        # Scorecards
        cols_ag = st.columns(min(len(rows), 3))
        for i, row in enumerate(rows):
            with cols_ag[i % 3]:
                rt_color = "#3fb950" if row["Avg RT (min)"] < 5 else "#d29922" if row["Avg RT (min)"] < 10 else "#f85149"
                fp = row["Basic Fill"]
                fp_color = "#3fb950" if fp >= 90 else "#d29922" if fp >= 70 else "#f85149"
                st.markdown(f"""
                <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;padding:16px;margin-bottom:12px">
                  <div style="font-size:14px;font-weight:700;color:#e6edf3;margin-bottom:10px;border-bottom:1px solid #21262d;padding-bottom:8px">{row['Agent']}</div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;color:#8b949e">Total</span><span style="font-size:12px;font-weight:600;color:#e6edf3">{row['Total Leads']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;color:#8b949e">Responded</span><span style="font-size:12px;font-weight:600;color:#3fb950">{row['Responded']} ({row['Responded %']})</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;color:#8b949e">MQL / MUL</span><span style="font-size:12px;font-weight:600;color:#e6edf3">{row['MQL']} / {row['MUL']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;color:#8b949e">Avg RT</span><span style="font-size:12px;font-weight:600;color:{rt_color}">{row['Avg RT']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;color:#8b949e">Lang Missing</span><span style="font-size:12px;font-weight:600;color:{'#f85149' if row['Lang Missing'] else '#3fb950'}">{row['Lang Missing']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;color:#8b949e">Task Missing</span><span style="font-size:12px;font-weight:600;color:{'#f85149' if row['Task Missing'] else '#3fb950'}">{row['Task Missing']}</span></div>
                  <div style="display:flex;justify-content:space-between"><span style="font-size:12px;color:#8b949e">Field Fill %</span><span style="font-size:12px;font-weight:600;color:{fp_color}">{row['Basic Fill %']}</span></div>
                </div>""", unsafe_allow_html=True)

        # Comparison chart
        st.markdown("---")
        st.markdown('<div class="section-header">Agent Comparison — Responded %</div>', unsafe_allow_html=True)
        resp_pct_vals = [float(r["Responded %"].replace("%","")) for r in rows]
        ag_colors = ["#3fb950" if v >= 70 else "#d29922" if v >= 50 else "#f85149" for v in resp_pct_vals]
        fig_ag = go.Figure(go.Bar(
            y=[r["Agent"] for r in rows], x=resp_pct_vals,
            orientation="h", marker_color=ag_colors,
            text=[f"{v:.1f}%" for v in resp_pct_vals], textposition="outside",
            textfont=dict(color="#e6edf3"),
        ))
        fig_ag.update_layout(**PLOTLY_THEME, xaxis=dict(range=[0,110], gridcolor="#21262d"),
                             yaxis=dict(tickfont=dict(color="#e6edf3"), autorange="reversed"),
                             height=max(200, len(rows)*50))
        st.plotly_chart(fig_ag, use_container_width=True)

        # Summary table
        st.markdown("---")
        disp_cols = ["Agent","Total Leads","Responded","Responded %","MQL","MUL","Lang Missing","Task Missing","Avg RT","Basic Fill %"]
        st.dataframe(df_ag[disp_cols], use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════
# TAB 8 — ALL LEADS TABLE
# ═════════════════════════════════════════════
with tabs[7]:
    st.markdown("### 📑 All Leads — Master Table")

    if df_all.empty:
        st.warning("No data.")
    else:
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)
        al_src   = fc1.selectbox("Source", ["All","New Leads","New Leads (9AM–9PM)","Activity Leads"], key="al_src")
        al_call  = fc2.selectbox("Call Status", ["All","Call Responded","Call Not Responded","Invalid Number"], key="al_call")
        al_stage = fc3.selectbox("Stage", ["All"] + sorted(df_all["contact_stage"].dropna().unique().tolist()), key="al_stage")
        al_own   = fc4.selectbox("Agent", ["All"] + sorted(df_all["owner"].dropna().unique().tolist()), key="al_own")
        al_cat   = fc5.selectbox("Category", ["All","MQL","MUL","Not Set"], key="al_cat")
        al_srch  = st.text_input("🔍 Search name, city, source…", key="al_srch")

        base_al = {
            "All": df_all,
            "New Leads": df_new,
            "New Leads (9AM–9PM)": df_new_biz,
            "Activity Leads": df_act,
        }.get(al_src, df_all)

        if not base_al.empty:
            if al_call != "All":
                base_al = base_al[base_al["call_response_status"].fillna("").str.strip() == al_call]
            if al_stage != "All":
                base_al = base_al[base_al["contact_stage"].fillna("").str.strip() == al_stage]
            if al_own != "All":
                base_al = base_al[base_al["owner"] == al_own]
            if al_cat == "MQL": base_al = base_al[base_al["lead_category"].fillna("").str.upper().eq("MQL")]
            elif al_cat == "MUL": base_al = base_al[base_al["lead_category"].fillna("").str.upper().eq("MUL")]
            elif al_cat == "Not Set": base_al = base_al[base_al["lead_category"].isna()]
            if al_srch:
                mask = base_al.apply(lambda r: r.astype(str).str.contains(al_srch, case=False).any(), axis=1)
                base_al = base_al[mask]

            st.markdown(f"**{len(base_al)} leads matching filters**")

            display_cols = [
                "name","owner","create_date","last_activity",
                "call_response_status","contact_stage","lead_category",
                "need_another_call","preferred_language","rt_mins",
                "disq_reason","product_interest","call_outcome",
                "contact_source","city","state","has_task","task_missing",
                "missing_basic",
            ]
            disp = base_al[[c for c in display_cols if c in base_al.columns]].copy()
            disp["create_date"] = disp["create_date"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
            disp["last_activity"] = disp["last_activity"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
            disp["rt_mins"] = disp["rt_mins"].apply(fmt_mins)
            disp["has_task"] = disp["has_task"].map({True:"✅", False:"❌"})
            disp["task_missing"] = disp["task_missing"].map({True:"⚠️ Yes",False:"—",None:"—"}).fillna("—")
            disp.columns = [c.replace("_"," ").title() for c in disp.columns]

            st.dataframe(style_df(disp), use_container_width=True, hide_index=True, height=500)

            # Download
            csv_buf = io.StringIO()
            disp.to_csv(csv_buf, index=False)
            st.download_button(
                "⬇️ Download filtered data as CSV",
                data=csv_buf.getvalue(),
                file_name="filtered_leads.csv",
                mime="text/csv",
            )

"""
IS Lead Dashboard — Streamlit App
==================================
Run with:
    pip install streamlit plotly pandas openpyxl numpy xlrd
    streamlit run is_lead_dashboard.py
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

[data-testid="stSidebar"] {
    background: #161b22 !important;
    border-right: 1px solid #21262d;
}

[data-testid="metric-container"] {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 10px;
    padding: 14px 18px;
}
[data-testid="metric-container"] label {
    color: #8b949e !important;
    font-size: 11px !important;
    text-transform: uppercase;
    letter-spacing: .06em;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #e6edf3 !important;
    font-size: 26px !important;
    font-weight: 700 !important;
}

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

h1, h2, h3, h4 { color: #e6edf3 !important; }
p, li { color: #8b949e; }

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

.section-header {
    font-size: 14px;
    font-weight: 600;
    color: #e6edf3;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 1px solid #21262d;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTS
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
    "times_contacted":      "Number of times contacted",
}

BASIC_FIELDS    = ["call_response_status", "contact_stage", "need_another_call", "lead_category", "disq_reason"]
BASIC_LABELS    = ["Call Response Status", "Contact Stage", "Need Another Call", "Lead Category", "Disq/Lost Reason"]

STATUS_COLORS = {"Call Responded": "#3fb950", "Call Not Responded": "#f85149", "Invalid Number": "#d29922"}
CAT_COLORS = {"MQL": "#58a6ff", "MUL": "#f85149", "Not Set": "#484f58"}
STAGE_COLORS = {
    "Lost": "#f85149", "Qualification": "#58a6ff", "Opportunity": "#3fb950",
    "Pass to field": "#bc8cff", "Raw Lead": "#d29922",
}
PLOTLY_THEME = dict(
    plot_bgcolor="#161b22", paper_bgcolor="#161b22",
    font=dict(color="#8b949e", family="DM Sans"), margin=dict(l=0, r=10, t=30, b=10),
)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def read_file(uploaded) -> pd.DataFrame:
    if uploaded is None: return pd.DataFrame()
    try:
        if uploaded.name.lower().endswith(".csv"):
            return pd.read_csv(uploaded, low_memory=False)
        return pd.read_excel(uploaded, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading {uploaded.name}: {e}")
        return pd.DataFrame()

def safe_str(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()

def parse_rt(s) -> float:
    try:
        parts = str(s).strip().split(":")
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
        return np.nan
    except: return np.nan

def fmt_mins(m) -> str:
    if m is None or (isinstance(m, float) and np.isnan(m)): return "—"
    m = float(m)
    h, rem = divmod(m, 60)
    if h >= 1: return f"{int(h)}h {int(rem)}m"
    return f"{int(rem)}m {int((rem % 1) * 60)}s"

def pct(num, den):
    return round(num / den * 100, 1) if den > 0 else 0.0

def normalise(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    out = pd.DataFrame(index=df.index)
    for key, raw in COL_MAP.items():
        if raw in df.columns: out[key] = df[raw]
        else:
            matched = [c for c in df.columns if key.replace("_", " ").lower() in c.lower()]
            out[key] = df[matched[0]] if matched else None
    
    # Preferred language fallback
    if out["preferred_language"].isna().all() and "Preferred language" in df.columns:
        out["preferred_language"] = df["Preferred language"]

    # Ensure times_contacted is numeric
    if "times_contacted" in out.columns:
        out["times_contacted"] = pd.to_numeric(out["times_contacted"], errors="coerce")

    for dc in ["create_date", "last_activity", "next_activity_date"]:
        out[dc] = pd.to_datetime(out[dc], errors="coerce")

    fn = out["first_name"].fillna("").astype(str).replace("nan", "")
    ln = out["last_name"].fillna("").astype(str).replace("nan", "")
    out["name"] = (fn + " " + ln).str.strip()
    out["rt_mins"] = out["response_time"].apply(parse_rt)
    out["missing_basic"] = out[BASIC_FIELDS].isnull().sum(axis=1)
    out["has_task"] = out["next_activity_date"].notna()
    out["needs_task"] = (safe_str(out["call_response_status"]).eq("Call Not Responded") | safe_str(out["need_another_call"]).eq("Yes"))
    out["task_missing"] = out["needs_task"] & ~out["has_task"]
    return out

def biz_filter(df: pd.DataFrame, start_h: int, end_h: int) -> pd.DataFrame:
    if df.empty: return df
    dt = pd.to_datetime(df["create_date"], errors="coerce")
    return df[dt.dt.hour.between(start_h, end_h - 1, inclusive="both").fillna(False)].copy()

# ─────────────────────────────────────────────
# PLOT BUILDERS
# ─────────────────────────────────────────────
def make_bar_h(labels, values, colors=None, pct_of=None):
    if not labels: return go.Figure()
    text = [f"{v}" if pct_of is None else f"{v}  ({pct(v, pct_of):.1f}%)" for v in values]
    fig = go.Figure(go.Bar(x=values, y=labels, orientation="h", marker_color=colors or "#58a6ff",
                           text=text, textposition="outside", textfont=dict(color="#e6edf3", size=11)))
    fig.update_layout(**PLOTLY_THEME, xaxis=dict(showgrid=False, showticklabels=False), yaxis=dict(autorange="reversed"), height=max(180, len(labels) * 44))
    return fig

def make_donut(labels, values, colors=None):
    if not labels or sum(values) == 0: return go.Figure()
    fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.55, marker_colors=colors, textinfo="percent+label"))
    fig.update_layout(**PLOTLY_THEME, showlegend=False, height=260)
    return fig

def progress_html(labels, filled, totals):
    rows = ""
    for lbl, f, tot in zip(labels, filled, totals):
        p = pct(f, tot)
        color = "#3fb950" if p >= 90 else "#d29922" if p >= 70 else "#f85149"
        bar = f'<div style="background:#21262d;border-radius:4px;height:7px;width:100%"><div style="width:{p}%;background:{color};height:7px;border-radius:4px"></div></div>'
        rows += f'<tr><td style="padding:8px 12px;color:#e6edf3;font-size:12px">{lbl}</td><td style="padding:8px 12px;width:100%">{bar}</td><td style="padding:8px 12px;color:{color};font-size:12px;font-weight:600">{f}/{tot} ({p}%)</td></tr>'
    return f'<table style="width:100%;border-collapse:collapse;background:#161b22;border-radius:8px"><tbody>{rows}</tbody></table>'

def alert_md(level, icon, msg):
    color = {"ok": "#3fb950", "warn": "#d29922", "err": "#f85149"}[level]
    bg = {"ok": "#0a1a0a", "warn": "#1a1500", "err": "#1c0a0a"}[level]
    return f'<div style="background:{bg};border:1px solid {color};border-left:4px solid {color};border-radius:8px;padding:10px 14px;margin-bottom:8px;font-size:13px;color:{color}">{icon} {msg}</div>'

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 IS Lead Dashboard")
    st.markdown("---")
    file1 = st.file_uploader("**File 1 — New Leads**", type=["xlsx", "xls", "csv"])
    file2 = st.file_uploader("**File 2 — Activity Leads**", type=["xlsx", "xls", "csv"])
    st.markdown("---")
    biz_start = st.slider("Business Hours Start", 0, 23, 9)
    biz_end = st.slider("Business Hours End", 1, 24, 21)

# ─────────────────────────────────────────────
# DATA PROCESSING
# ─────────────────────────────────────────────
raw1, raw2 = read_file(file1), read_file(file2)
df_new = normalise(raw1)
df_act = normalise(raw2)
df_new_biz = biz_filter(df_new, biz_start, biz_end)
frames = [d for d in [df_new, df_act] if not d.empty]
df_all = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

if df_all.empty:
    st.info("📊 Welcome! Please upload lead export files in the sidebar to view the dashboard.")
    st.stop()

# ─────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────
st.markdown("# 📊 IS Lead Dashboard")
st.markdown("---")

tabs = st.tabs(["🏠 Overview", "⏱️ Response Time", "📞 Call Status", "📋 Field Completion", "✅ Task Audit", "🏷️ Lead Qualification", "👤 By Agent", "📑 All Leads"])

# 🏠 OVERVIEW
with tabs[0]:
    n_resp = int(safe_str(df_new["call_response_status"]).eq("Call Responded").sum()) if not df_new.empty else 0
    avg_tc_global = df_all["times_contacted"].mean()
    
    ov1, ov2, ov3, ov4, ov5 = st.columns(5)
    ov1.metric("New Leads", len(df_new))
    ov2.metric("Activity Leads", len(df_act))
    ov3.metric("Responded (New)", n_resp, f"{pct(n_resp, len(df_new))}%")
    ov4.metric("Avg Contacts/Lead", f"{avg_tc_global:.2f}" if not np.isnan(avg_tc_global) else "—")
    ov5.metric("Tasks Missing ⚠️", int(df_all["task_missing"].sum()), delta_color="inverse")

# ⏱️ RESPONSE TIME
with tabs[1]:
    st.markdown("### ⏱️ Response Time Analysis")
    if df_new_biz.empty: st.warning("No data in business hours.")
    else:
        rt_s = df_new_biz["rt_mins"].dropna()
        r1, r2, r3 = st.columns(3)
        r1.metric("Average RT", fmt_mins(rt_s.mean()))
        r2.metric("Median RT", fmt_mins(rt_s.median()))
        r3.metric("Under 5 min ✅", f"{pct((rt_s < 5).sum(), len(rt_s))}%")

# 👤 BY AGENT
with tabs[6]:
    st.markdown("### 👤 Agent Performance Deep-Dive")
    ag_src = st.selectbox("Data Source", ["All", "New Leads", "Activity Leads"], key="ag_src_deep")
    ag_base = {"All": df_all, "New Leads": df_new, "Activity Leads": df_act}[ag_src]
    
    owners = sorted(ag_base["owner"].dropna().unique().tolist())
    rows_ag = []
    for owner in owners:
        ol = ag_base[ag_base["owner"] == owner]
        resp = ol[safe_str(ol["call_response_status"]).eq("Call Responded")]
        rt_v = ol["rt_mins"].dropna()
        tc_v = ol["times_contacted"].dropna()
        
        rows_ag.append({
            "Agent": owner,
            "Total": len(ol),
            "Resp%": pct(len(resp), len(ol)),
            "Avg RT": fmt_mins(rt_v.mean()) if not rt_v.empty else "—",
            "Avg RT Raw": rt_v.mean(),
            "Avg Contacts": round(tc_v.mean(), 2) if not tc_v.empty else 0.0,
            "Total Contacts": int(tc_v.sum()),
            "Task Missing": int(ol["task_missing"].sum()),
            "Fill%": pct(len(ol) * len(BASIC_FIELDS) - ol["missing_basic"].sum(), len(ol) * len(BASIC_FIELDS))
        })

    ag_grid = st.columns(3)
    for i, row in enumerate(rows_ag):
        with ag_grid[i % 3]:
            st.markdown(f"""
            <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;padding:16px;margin-bottom:12px">
                <div style="font-weight:700;color:#e6edf3;border-bottom:1px solid #21262d;padding-bottom:8px;margin-bottom:10px">{row['Agent']}</div>
                <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px"><span style="color:#8b949e">Total Leads</span><span style="color:#e6edf3">{row['Total']}</span></div>
                <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px"><span style="color:#8b949e">Response Rate</span><span style="color:#3fb950">{row['Resp%']}%</span></div>
                <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px"><span style="color:#8b949e">Avg Contacts/Lead</span><span style="color:#58a6ff;font-weight:600">{row['Avg Contacts']}</span></div>
                <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px"><span style="color:#8b949e">Total Contact Attempts</span><span style="color:#e6edf3">{row['Total Contacts']}</span></div>
                <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px"><span style="color:#8b949e">Avg Response Time</span><span style="color:#e6edf3">{row['Avg RT']}</span></div>
                <div style="display:flex;justify-content:space-between;font-size:12px"><span style="color:#8b949e">Task Missing</span><span style="color:#f85149">{row['Task Missing']}</span></div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="section-header">Average Contacts per Lead — by Agent</div>', unsafe_allow_html=True)
    ag_df = pd.DataFrame(rows_ag).sort_values("Avg Contacts", ascending=False)
    fig_tc = go.Figure(go.Bar(y=ag_df["Agent"], x=ag_df["Avg Contacts"], orientation="h", marker_color="#58a6ff",
                              text=ag_df["Avg Contacts"], textposition="outside", textfont=dict(color="#e6edf3")))
    fig_tc.update_layout(**PLOTLY_THEME, xaxis=dict(title="Avg Contacts Made"), yaxis=dict(autorange="reversed"), height=max(200, len(rows_ag)*45))
    st.plotly_chart(fig_tc, use_container_width=True)

    st.dataframe(ag_df[["Agent", "Total", "Resp%", "Avg Contacts", "Total Contacts", "Avg RT", "Fill%"]], use_container_width=True, hide_index=True)

# 📑 ALL LEADS
with tabs[7]:
    st.markdown("### 📑 Master Lead Table")
    st.dataframe(df_all.fillna("—"), use_container_width=True, hide_index=True)
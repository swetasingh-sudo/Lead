"""
IS Lead Dashboard — Streamlit App
==================================
Run with:
    pip install streamlit plotly pandas openpyxl numpy xlrd
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

[data-testid="stExpander"] {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 10px;
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
.health-ok .health-title   { color: #3fb950; }
.health-warn .health-title { color: #d29922; }
.health-err .health-title  { color: #f85149; }

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
}

BASIC_FIELDS    = ["call_response_status", "contact_stage", "need_another_call", "lead_category", "disq_reason"]
EXTENDED_FIELDS = ["product_interest", "primary_reason", "purchase_timeline", "ecg_volume",
                   "echo_volume", "decision_maker", "current_workflow", "num_locations",
                   "customer_segment", "call_outcome"]
BASIC_LABELS    = ["Call Response Status", "Contact Stage", "Need Another Call", "Lead Category", "Disq/Lost Reason"]
EXTENDED_LABELS = ["Product Interest", "Primary Reason", "Purchase Timeline", "ECG Volume",
                   "ECHO Volume", "Decision Maker", "Current Workflow", "No. Locations",
                   "Customer Segment", "Call Outcome"]

STATUS_COLORS = {
    "Call Responded":     "#3fb950",
    "Call Not Responded": "#f85149",
    "Invalid Number":     "#d29922",
}
CAT_COLORS = {"MQL": "#58a6ff", "MUL": "#f85149", "Not Set": "#484f58"}
STAGE_COLORS = {
    "Lost": "#f85149", "Qualification": "#58a6ff", "Opportunity": "#3fb950",
    "Pass to field": "#bc8cff", "Raw Lead": "#d29922",
    "Zigment Unqualified": "#484f58", "Not Set": "#30363d",
    "Contact in Future": "#d29922", "Cold": "#484f58", "International": "#bc8cff",
}
PLOTLY_THEME = dict(
    plot_bgcolor="#161b22",
    paper_bgcolor="#161b22",
    font=dict(color="#8b949e", family="DM Sans"),
    margin=dict(l=0, r=10, t=30, b=10),
)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def read_file(uploaded) -> pd.DataFrame:
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


def safe_str(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def parse_rt(s) -> float:
    try:
        parts = str(s).strip().split(":")
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
        return np.nan
    except Exception:
        return np.nan


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
    if df.empty:
        return df
    out = pd.DataFrame(index=df.index)
    for key, raw in COL_MAP.items():
        if raw in df.columns:
            out[key] = df[raw]
        else:
            matched = [c for c in df.columns if key.replace("_", " ").lower() in c.lower()]
            out[key] = df[matched[0]] if matched else None

    # Preferred language fallback
    if out["preferred_language"].isna().all() and "Preferred language" in df.columns:
        out["preferred_language"] = df["Preferred language"]

    # Dates
    for dc in ["create_date", "last_activity", "last_contacted", "next_activity_date"]:
        out[dc] = pd.to_datetime(out[dc], errors="coerce")

    # Name
    fn = out["first_name"].fillna("").astype(str).replace("nan", "")
    ln = out["last_name"].fillna("").astype(str).replace("nan", "")
    out["name"] = (fn + " " + ln).str.strip()

    # Response time
    out["rt_mins"] = out["response_time"].apply(parse_rt)

    # Missing basic fields count
    out["missing_basic"] = out[BASIC_FIELDS].isnull().sum(axis=1)

    # MQL + responded flag
    out["is_mql_responded"] = (
        safe_str(out["lead_category"]).str.upper().eq("MQL") &
        safe_str(out["call_response_status"]).eq("Call Responded")
    )

    # Task flags
    out["has_task"]    = out["next_activity_date"].notna()
    out["needs_task"]  = (
        safe_str(out["call_response_status"]).eq("Call Not Responded") |
        safe_str(out["need_another_call"]).eq("Yes")
    )
    out["task_missing"] = out["needs_task"] & ~out["has_task"]

    return out


def biz_filter(df: pd.DataFrame, start_h: int, end_h: int) -> pd.DataFrame:
    if df.empty:
        return df
    dt   = pd.to_datetime(df["create_date"], errors="coerce")
    mask = dt.dt.hour.between(start_h, end_h - 1, inclusive="both")
    return df[mask.fillna(False)].copy()


def style_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.fillna("—")


# ─────────────────────────────────────────────
# PLOT BUILDERS
# ─────────────────────────────────────────────
def make_bar_h(labels, values, colors=None, pct_of=None):
    if not labels:
        return go.Figure()
    if colors is None:
        colors = ["#58a6ff"] * len(labels)
    text = [f"{v}" if pct_of is None else f"{v}  ({pct(v, pct_of):.1f}%)" for v in values]
    fig  = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=colors,
        text=text, textposition="outside",
        textfont=dict(color="#e6edf3", size=11),
        hovertemplate="%{y}: %{x}<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_THEME,
        xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
        yaxis=dict(tickfont=dict(color="#e6edf3", size=11), autorange="reversed"),
        height=max(180, len(labels) * 44),
    )
    return fig


def make_donut(labels, values, colors=None):
    if not labels or sum(values) == 0:
        return go.Figure()
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.55,
        marker_colors=colors or px.colors.qualitative.Set2,
        textinfo="percent+label",
        textfont=dict(color="#e6edf3", size=11),
        hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
    ))
    fig.update_layout(**PLOTLY_THEME, showlegend=False, height=260)
    return fig


def progress_html(labels, filled, totals):
    rows = ""
    for lbl, f, tot in zip(labels, filled, totals):
        p     = pct(f, tot) if tot else 0
        color = "#3fb950" if p >= 90 else "#d29922" if p >= 70 else "#f85149"
        bar   = (f'<div style="background:#21262d;border-radius:4px;height:7px;width:100%">'
                 f'<div style="width:{p:.1f}%;background:{color};height:7px;border-radius:4px"></div></div>')
        rows += (f'<tr>'
                 f'<td style="padding:8px 12px;color:#e6edf3;font-size:12px;white-space:nowrap">{lbl}</td>'
                 f'<td style="padding:8px 12px;width:100%">{bar}</td>'
                 f'<td style="padding:8px 12px;color:{color};font-size:12px;font-weight:600;white-space:nowrap">'
                 f'{f}/{tot} ({p:.1f}%)</td></tr>')
    return (f'<table style="width:100%;border-collapse:collapse;background:#161b22;'
            f'border-radius:8px;overflow:hidden"><tbody>{rows}</tbody></table>')


def alert_md(level, icon, msg):
    bg     = {"ok": "#0a1a0a", "warn": "#1a1500", "err": "#1c0a0a"}[level]
    border = {"ok": "#3fb950", "warn": "#d29922", "err": "#f85149"}[level]
    color  = {"ok": "#3fb950", "warn": "#d29922", "err": "#f85149"}[level]
    return (f'<div style="background:{bg};border:1px solid {border};border-left:4px solid {border};'
            f'border-radius:8px;padding:10px 14px;margin-bottom:8px;font-size:13px;color:{color}">'
            f'{icon} {msg}</div>')


# ─────────────────────────────────────────────
# DATA HEALTH
# ─────────────────────────────────────────────
def data_health(df: pd.DataFrame, label: str):
    if df.empty:
        st.markdown(
            f'<div class="health-box health-err"><div class="health-title">❌ {label} — No data loaded</div></div>',
            unsafe_allow_html=True)
        return
    total        = len(df)
    key_cols     = list(COL_MAP.values())
    present      = [c for c in key_cols if c in df.columns]
    missing_cols = [c for c in key_cols if c not in df.columns]
    null_rates   = {c: int(df[c].isna().sum()) for c in present}
    high_null    = {c: v for c, v in null_rates.items() if v > 0 and v / total > 0.5}
    score        = max(0, 100 - len(missing_cols) * 2 - len(high_null) * 3)
    level        = "ok" if score >= 80 else "warn" if score >= 60 else "err"
    icon         = "✅" if level == "ok" else "⚠️" if level == "warn" else "❌"

    st.markdown(f"""
    <div class="health-box health-{level}">
      <div class="health-title">{icon} {label} — Health Score: {score}/100</div>
      <p style="font-size:12px;color:#8b949e;margin:0">
        {total} rows &nbsp;·&nbsp; {len(present)}/{len(key_cols)} key columns found
        {f" &nbsp;·&nbsp; <b style='color:#f85149'>{len(missing_cols)} columns missing</b>" if missing_cols else ""}
      </p>
    </div>""", unsafe_allow_html=True)

    if missing_cols or high_null:
        with st.expander(f"🔍 Details — {label}"):
            if missing_cols:
                st.markdown("**Missing columns** (will show '—' in dashboard):")
                st.code(", ".join(missing_cols))
            if high_null:
                st.markdown("**Columns >50% null:**")
                for c, v in high_null.items():
                    st.markdown(f"- `{c}`: {v}/{total} ({pct(v, total):.1f}%)")


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 IS Lead Dashboard")
    st.markdown("---")
    st.markdown("### 📂 Upload Files")
    file1 = st.file_uploader(
        "**File 1 — New Leads** (created yesterday)",
        type=["xlsx", "xls", "csv"], key="upload_f1",
        help="Leads created yesterday",
    )
    file2 = st.file_uploader(
        "**File 2 — Activity Leads** (worked yesterday)",
        type=["xlsx", "xls", "csv"], key="upload_f2",
        help="Leads that had activity done on them yesterday",
    )
    st.markdown("---")
    st.markdown("### ⚙️ Settings")
    biz_start = st.slider("Business Hours Start", 0, 23, 9,  key="biz_start_slider")
    biz_end   = st.slider("Business Hours End",   1, 24, 21, key="biz_end_slider")
    st.caption("Used for response time & raw-stage calculations")
    st.markdown("---")
    st.caption("Supports .xlsx / .xls / .csv · Missing columns handled automatically")


# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────
raw1 = read_file(file1)
raw2 = read_file(file2)

df_new     = normalise(raw1) if not raw1.empty else pd.DataFrame()
df_act     = normalise(raw2) if not raw2.empty else pd.DataFrame()
df_new_biz = biz_filter(df_new, biz_start, biz_end) if not df_new.empty else pd.DataFrame()

frames = [d for d in [df_new, df_act] if not d.empty]
df_all = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
no_data = df_all.empty


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
hc1, hc2 = st.columns([4, 1])
with hc1:
    st.markdown("# 📊 IS Lead Dashboard")
    if not df_all.empty:
        parts = []
        for df, lbl in [(df_new, "New Leads"), (df_act, "Activity Leads")]:
            if not df.empty and df["create_date"].notna().any():
                d = df["create_date"].dropna().dt.date
                parts.append(f"**{lbl}**: {d.min()} → {d.max()}")
        st.caption("  |  ".join(parts) if parts else "Data loaded")
    else:
        st.caption("Upload files from the sidebar to begin →")
with hc2:
    if not df_all.empty:
        st.metric("Total Leads", len(df_all), label_visibility="visible")


# ─────────────────────────────────────────────
# DATA HEALTH
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("### 🩺 Data Health")
dh1, dh2 = st.columns(2)
with dh1:
    data_health(raw1, "File 1 — New Leads")
with dh2:
    data_health(raw2, "File 2 — Activity Leads")

if no_data:
    st.info("⬆️ Upload at least one file from the sidebar to see the full dashboard.")
    st.stop()


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tabs = st.tabs([
    "🏠 Overview",
    "⏱️ Response Time",
    "📞 Call Status",
    "📋 Field Completion",
    "✅ Task Audit",
    "🏷️ Lead Qualification",
    "👤 By Agent",
    "📑 All Leads",
])


# ══════════════════════════════════════════════
# TAB 0 — OVERVIEW
# ══════════════════════════════════════════════
with tabs[0]:
    nl  = df_new
    al  = df_act
    nlb = df_new_biz

    def _cnt(df, col, val):
        if df.empty or col not in df.columns:
            return 0
        return int(safe_str(df[col]).eq(val).sum())

    n_new  = len(nl)
    n_act  = len(al)
    n_biz  = len(nlb)
    n_resp = _cnt(nl, "call_response_status", "Call Responded")
    rt_v   = nlb["rt_mins"].dropna().tolist() if not nlb.empty else []
    avg_rt = float(np.mean(rt_v)) if rt_v else None
    med_rt = float(np.median(rt_v)) if rt_v else None
    n_mql  = int(safe_str(df_all["lead_category"]).str.upper().eq("MQL").sum())
    n_mul  = int(safe_str(df_all["lead_category"]).str.upper().eq("MUL").sum())
    n_task = int(df_all["task_missing"].sum())

    ov1, ov2, ov3, ov4, ov5, ov6 = st.columns(6)
    ov1.metric("New Leads",         n_new,          f"{n_biz} in {biz_start}–{biz_end}h")
    ov2.metric("Activity Leads",    n_act)
    ov3.metric("Responded (New)",   n_resp,         f"{pct(n_resp, n_new):.1f}%")
    ov4.metric("Avg Response Time", fmt_mins(avg_rt),f"Median {fmt_mins(med_rt)}")
    ov5.metric("MQL / MUL",         f"{n_mql} / {n_mul}")
    ov6.metric("Tasks Missing ⚠️",  n_task,         delta_color="inverse")

    st.markdown("---")
    ov_c1, ov_c2 = st.columns(2)

    with ov_c1:
        st.markdown('<div class="section-header">📞 Call Response — New Leads</div>', unsafe_allow_html=True)
        if not nl.empty:
            vc_nl = nl["call_response_status"].str.strip().value_counts()
            s_lbls = [s for s in ["Call Responded","Call Not Responded","Invalid Number"] if s in vc_nl.index]
            s_vals = [int(vc_nl.get(s, 0)) for s in s_lbls]
            s_clrs = [STATUS_COLORS.get(s, "#8b949e") for s in s_lbls]
            st.plotly_chart(make_donut(s_lbls, s_vals, s_clrs),
                            use_container_width=True, key="ov_donut_call_status")

    with ov_c2:
        st.markdown('<div class="section-header">🏷️ Lead Category — All Leads</div>', unsafe_allow_html=True)
        cat_vc = df_all["lead_category"].fillna("Not Set").str.strip().value_counts()
        c_lbls = cat_vc.index.tolist()
        c_vals = cat_vc.values.tolist()
        c_clrs = [CAT_COLORS.get(l, "#8b949e") for l in c_lbls]
        st.plotly_chart(make_donut(c_lbls, c_vals, c_clrs),
                        use_container_width=True, key="ov_donut_lead_category")

    ov_c3, ov_c4 = st.columns(2)

    with ov_c3:
        st.markdown('<div class="section-header">📋 Contact Stage — New Leads</div>', unsafe_allow_html=True)
        if not nl.empty:
            stg    = nl["contact_stage"].fillna("Not Set").str.strip().value_counts()
            st_l   = stg.index.tolist()
            st_v   = stg.values.tolist()
            st_c   = [STAGE_COLORS.get(l, "#8b949e") for l in st_l]
            st.plotly_chart(make_bar_h(st_l, st_v, st_c, pct_of=n_new),
                            use_container_width=True, key="ov_bar_contact_stage")

    with ov_c4:
        st.markdown('<div class="section-header">⚠️ Key Alerts</div>', unsafe_allow_html=True)
        nl_resp   = nl[safe_str(nl["call_response_status"]).eq("Call Responded")] if not nl.empty else pd.DataFrame()
        lang_miss = int(nl_resp["preferred_language"].isna().sum()) if not nl_resp.empty else 0
        cat_miss  = int(nl_resp["lead_category"].isna().sum())      if not nl_resp.empty else 0
        mul_nd    = int((safe_str(df_all["lead_category"]).str.upper().eq("MUL") & df_all["disq_reason"].isna()).sum())
        raw_cnt   = int((nlb["contact_stage"].isna() | safe_str(nlb["contact_stage"]).eq("Raw Lead")).sum()) if not nlb.empty else 0

        html_al  = ""
        html_al += alert_md("err" if lang_miss > 0 else "ok", "🌐",
                            f"{lang_miss} responded calls missing preferred language ({pct(lang_miss, max(len(nl_resp),1)):.1f}%)")
        html_al += alert_md("err" if cat_miss > 0 else "ok",  "🏷️",
                            f"{cat_miss} responded calls missing lead category")
        html_al += alert_md("warn" if mul_nd > 0 else "ok",   "❌",
                            f"{mul_nd} MUL leads missing disqualification reason")
        html_al += alert_md("err" if n_task > 0 else "ok",    "📅",
                            f"{n_task} leads requiring tasks have NO task created")
        html_al += alert_md("warn" if raw_cnt > 0 else "ok",  "🔴",
                            f"{raw_cnt} leads in Raw/unworked stage ({biz_start}–{biz_end}h window)")
        st.markdown(html_al, unsafe_allow_html=True)


# ══════════════════════════════════════════════
# TAB 1 — RESPONSE TIME
# ══════════════════════════════════════════════
with tabs[1]:
    st.markdown("### ⏱️ Response Time Analysis")
    st.caption(f"Based on new leads created between {biz_start}:00–{biz_end}:00 only")

    if df_new_biz.empty:
        st.warning("No new leads in the selected business-hours window. Adjust the slider in the sidebar.")
    else:
        rt_s  = df_new_biz["rt_mins"].dropna()
        avg_v = float(rt_s.mean())   if len(rt_s) else None
        med_v = float(rt_s.median()) if len(rt_s) else None
        u5    = int((rt_s < 5).sum())
        o15   = int((rt_s > 15).sum())

        r1, r2, r3, r4, r5 = st.columns(5)
        r1.metric("Leads with RT",  len(rt_s))
        r2.metric("Average RT",     fmt_mins(avg_v))
        r3.metric("Median RT",      fmt_mins(med_v))
        r4.metric("Under 5 min ✅", u5,  f"{pct(u5,  len(rt_s)):.1f}%")
        r5.metric("Over 15 min ⚠️", o15, f"{pct(o15, len(rt_s)):.1f}%", delta_color="inverse")

        st.markdown("---")
        rt_c1, rt_c2 = st.columns(2)

        with rt_c1:
            st.markdown('<div class="section-header">By Agent — Avg & Median RT</div>', unsafe_allow_html=True)
            ag_rt = (
                df_new_biz.dropna(subset=["rt_mins"])
                .groupby("owner")["rt_mins"]
                .agg(["mean", "median", "count"])
                .reset_index()
                .sort_values("mean")
            )
            if not ag_rt.empty:
                ag_rt.columns = ["Agent", "Avg", "Median", "Count"]
                rt_colors = [
                    "#3fb950" if v < 5 else "#d29922" if v < 10 else "#f85149"
                    for v in ag_rt["Avg"]
                ]
                fig_rta = go.Figure()
                fig_rta.add_trace(go.Bar(
                    y=ag_rt["Agent"], x=ag_rt["Avg"], name="Avg RT",
                    orientation="h", marker_color=rt_colors,
                    text=[fmt_mins(v) for v in ag_rt["Avg"]],
                    textposition="outside", textfont=dict(color="#e6edf3", size=10),
                    hovertemplate="%{y}<br>Avg: %{x:.1f}m<extra></extra>",
                ))
                fig_rta.add_trace(go.Scatter(
                    y=ag_rt["Agent"], x=ag_rt["Median"], name="Median RT",
                    mode="markers", marker=dict(color="#bc8cff", size=10, symbol="diamond"),
                    hovertemplate="%{y}<br>Median: %{x:.1f}m<extra></extra>",
                ))
                fig_rta.update_layout(
                    **PLOTLY_THEME,
                    xaxis=dict(showgrid=True, gridcolor="#21262d", zeroline=False),
                    yaxis=dict(tickfont=dict(color="#e6edf3"), autorange="reversed"),
                    legend=dict(orientation="h", y=1.08, x=0),
                    height=max(250, len(ag_rt) * 52),
                )
                st.plotly_chart(fig_rta, use_container_width=True, key="rt_agent_combo_bar")
                ag_rt_disp = ag_rt.copy()
                ag_rt_disp["Avg"]    = ag_rt_disp["Avg"].apply(fmt_mins)
                ag_rt_disp["Median"] = ag_rt_disp["Median"].apply(fmt_mins)
                st.dataframe(ag_rt_disp, use_container_width=True, hide_index=True)

        with rt_c2:
            st.markdown('<div class="section-header">Distribution Buckets</div>', unsafe_allow_html=True)
            bkts = {
                "< 2 min":   int((rt_s < 2).sum()),
                "2–5 min":   int(((rt_s >= 2)  & (rt_s < 5)).sum()),
                "5–10 min":  int(((rt_s >= 5)  & (rt_s < 10)).sum()),
                "10–15 min": int(((rt_s >= 10) & (rt_s < 15)).sum()),
                "15–30 min": int(((rt_s >= 15) & (rt_s < 30)).sum()),
                "> 30 min":  int((rt_s >= 30).sum()),
            }
            b_clrs = ["#3fb950","#3fb950","#d29922","#d29922","#f85149","#f85149"]
            st.plotly_chart(
                make_bar_h(list(bkts.keys()), list(bkts.values()), b_clrs, pct_of=len(rt_s)),
                use_container_width=True, key="rt_bucket_bar",
            )
            st.markdown('<div class="section-header" style="margin-top:12px">Histogram (minutes)</div>',
                        unsafe_allow_html=True)
            fig_hist = px.histogram(rt_s, nbins=20, color_discrete_sequence=["#58a6ff"],
                                    labels={"value": "Minutes"})
            fig_hist.update_layout(**PLOTLY_THEME, height=200,
                                   xaxis=dict(gridcolor="#21262d"),
                                   yaxis=dict(gridcolor="#21262d"))
            st.plotly_chart(fig_hist, use_container_width=True, key="rt_histogram_px")

        st.markdown("---")
        st.markdown('<div class="section-header">Individual Response Times</div>', unsafe_allow_html=True)
        rt_srch = st.text_input("Search agent / name", key="rt_search_text")
        rt_disp = df_new_biz[["name","owner","create_date","call_response_status","rt_mins","contact_stage"]].copy()
        rt_disp["create_date"] = rt_disp["create_date"].dt.strftime("%H:%M").fillna("—")
        rt_disp["rt_mins"]     = rt_disp["rt_mins"].apply(fmt_mins)
        rt_disp.columns = ["Name","Agent","Created","Call Status","Response Time","Stage"]
        if rt_srch:
            mask    = rt_disp.apply(lambda r: r.astype(str).str.contains(rt_srch, case=False).any(), axis=1)
            rt_disp = rt_disp[mask]
        st.dataframe(style_df(rt_disp), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# TAB 2 — CALL STATUS
# ══════════════════════════════════════════════
with tabs[2]:
    st.markdown("### 📞 Call Status & Preferred Language")

    all_resp  = df_all[safe_str(df_all["call_response_status"]).eq("Call Responded")] if not df_all.empty else pd.DataFrame()
    nl_resp   = df_new[safe_str(df_new["call_response_status"]).eq("Call Responded")] if not df_new.empty else pd.DataFrame()
    al_resp   = df_act[safe_str(df_act["call_response_status"]).eq("Call Responded")] if not df_act.empty else pd.DataFrame()
    lang_miss = int(all_resp["preferred_language"].isna().sum()) if not all_resp.empty else 0
    lang_fill = len(all_resp) - lang_miss

    cs1,cs2,cs3,cs4,cs5 = st.columns(5)
    cs1.metric("Responded — New",      len(nl_resp), f"{pct(len(nl_resp), max(len(df_new),1)):.1f}%")
    cs2.metric("Responded — Activity", len(al_resp), f"{pct(len(al_resp), max(len(df_act),1)):.1f}%")
    cs3.metric("Total Responded",      len(all_resp),f"{pct(len(all_resp),len(df_all)):.1f}%")
    cs4.metric("Language Filled ✅",   lang_fill,    f"{pct(lang_fill, max(len(all_resp),1)):.1f}%")
    cs5.metric("Language Missing ⚠️",  lang_miss,    f"{pct(lang_miss, max(len(all_resp),1)):.1f}%", delta_color="inverse")

    st.markdown("---")
    cs_c1, cs_c2 = st.columns(2)
    STATUS_ORDER = ["Call Responded", "Call Not Responded", "Invalid Number"]

    with cs_c1:
        st.markdown('<div class="section-header">New Leads Call Status</div>', unsafe_allow_html=True)
        if not df_new.empty:
            vc_n = df_new["call_response_status"].str.strip().value_counts()
            n_v  = [int(vc_n.get(s, 0)) for s in STATUS_ORDER]
            st.plotly_chart(
                make_bar_h(STATUS_ORDER, n_v, [STATUS_COLORS[s] for s in STATUS_ORDER], pct_of=len(df_new)),
                use_container_width=True, key="cs_new_leads_bar",
            )

    with cs_c2:
        st.markdown('<div class="section-header">Activity Leads Call Status</div>', unsafe_allow_html=True)
        if not df_act.empty:
            vc_a = df_act["call_response_status"].str.strip().value_counts()
            a_v  = [int(vc_a.get(s, 0)) for s in STATUS_ORDER]
            st.plotly_chart(
                make_bar_h(STATUS_ORDER, a_v, [STATUS_COLORS[s] for s in STATUS_ORDER], pct_of=len(df_act)),
                use_container_width=True, key="cs_activity_leads_bar",
            )

    st.markdown("---")
    st.markdown('<div class="section-header">🌐 Preferred Language Completion by Agent (Responded Calls)</div>',
                unsafe_allow_html=True)
    if not all_resp.empty:
        lang_ag = (
            all_resp.assign(has_lang=all_resp["preferred_language"].notna())
            .groupby("owner").agg(filled=("has_lang","sum"), total=("has_lang","count"))
            .reset_index().sort_values("filled", ascending=False)
        )
        st.markdown(progress_html(lang_ag["owner"].tolist(),
                                  lang_ag["filled"].astype(int).tolist(),
                                  lang_ag["total"].astype(int).tolist()),
                    unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="section-header">Responded Calls Detail Table</div>', unsafe_allow_html=True)
    fc1, fc2, fc3 = st.columns(3)
    cs_src  = fc1.selectbox("Source",   ["All","New Leads","Activity Leads"], key="cs_src_selectbox")
    cs_lang = fc2.selectbox("Language", ["All","Filled","Missing"],           key="cs_lang_selectbox")
    cs_srch = fc3.text_input("Search",                                        key="cs_search_text")

    base_cs = {"All": df_all, "New Leads": df_new, "Activity Leads": df_act}[cs_src]
    if not base_cs.empty:
        base_cs = base_cs[safe_str(base_cs["call_response_status"]).eq("Call Responded")]
        if cs_lang == "Filled":  base_cs = base_cs[base_cs["preferred_language"].notna()]
        if cs_lang == "Missing": base_cs = base_cs[base_cs["preferred_language"].isna()]
        if cs_srch:
            mask    = base_cs.apply(lambda r: r.astype(str).str.contains(cs_srch, case=False).any(), axis=1)
            base_cs = base_cs[mask]
        st.markdown(f"**{len(base_cs)} leads**")
        st.dataframe(style_df(base_cs[["name","owner","call_response_status","preferred_language",
                                       "lead_category","contact_stage"]]),
                     use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# TAB 3 — FIELD COMPLETION
# ══════════════════════════════════════════════
with tabs[3]:
    st.markdown("### 📋 Field Completion Audit")
    fc_view = st.radio("View", ["Basic Fields (All Leads)", "Extended Fields (MQL + Responded)"],
                       horizontal=True, key="fc_view_radio_tab3")

    if df_all.empty:
        st.warning("No data.")
    elif "Basic" in fc_view:
        total_fc = len(df_all)
        filled_b = [int(df_all[f].notna().sum()) for f in BASIC_FIELDS]

        bfc = st.columns(len(BASIC_FIELDS))
        for i, (lbl, f) in enumerate(zip(BASIC_LABELS, filled_b)):
            bfc[i].metric(lbl, f"{pct(f, total_fc):.1f}%", f"{total_fc-f} missing")

        st.markdown("---")
        ff_c1, ff_c2 = st.columns(2)

        with ff_c1:
            st.markdown('<div class="section-header">Overall Field Completion</div>', unsafe_allow_html=True)
            st.markdown(progress_html(BASIC_LABELS, filled_b, [total_fc]*len(BASIC_FIELDS)),
                        unsafe_allow_html=True)

        with ff_c2:
            f_sel = st.selectbox("Drill by Agent — Field", BASIC_LABELS, key="ff_field_selectbox")
            f_key = BASIC_FIELDS[BASIC_LABELS.index(f_sel)]
            ag_ff = (
                df_all.assign(has_val=df_all[f_key].notna())
                .groupby("owner").agg(filled=("has_val","sum"), total=("has_val","count"))
                .reset_index().sort_values("filled", ascending=False)
            )
            st.markdown('<div class="section-header">By Agent</div>', unsafe_allow_html=True)
            st.markdown(progress_html(ag_ff["owner"].tolist(),
                                      ag_ff["filled"].astype(int).tolist(),
                                      ag_ff["total"].astype(int).tolist()),
                        unsafe_allow_html=True)

        st.markdown("---")
        st.markdown('<div class="section-header">Lead-Level Audit Table</div>', unsafe_allow_html=True)
        ff_f1, ff_f2, ff_f3, ff_f4 = st.columns(4)
        ff_src   = ff_f1.selectbox("Source",  ["All","New Leads","Activity Leads"], key="ff_src_selectbox")
        ff_miss  = ff_f2.selectbox("Filter",  ["Show All","Has Missing","Complete"], key="ff_miss_selectbox")
        ff_owner = ff_f3.selectbox("Agent",
                                   ["All"] + sorted(df_all["owner"].dropna().unique().tolist()),
                                   key="ff_owner_selectbox")
        ff_srch  = ff_f4.text_input("Search", key="ff_search_text")

        base_ff = {"All": df_all, "New Leads": df_new, "Activity Leads": df_act}[ff_src]
        if not base_ff.empty:
            if ff_miss == "Has Missing": base_ff = base_ff[base_ff["missing_basic"] > 0]
            if ff_miss == "Complete":    base_ff = base_ff[base_ff["missing_basic"] == 0]
            if ff_owner != "All":        base_ff = base_ff[base_ff["owner"] == ff_owner]
            if ff_srch:
                mask    = base_ff.apply(lambda r: r.astype(str).str.contains(ff_srch, case=False).any(), axis=1)
                base_ff = base_ff[mask]
            show_ff = base_ff[["name","owner","call_response_status","contact_stage",
                               "need_another_call","lead_category","disq_reason","missing_basic"]].copy()
            st.markdown(f"**{len(show_ff)} leads**")
            st.dataframe(style_df(show_ff), use_container_width=True, hide_index=True)

    else:  # Extended
        mql_r = df_all[df_all["is_mql_responded"]] if not df_all.empty else pd.DataFrame()
        st.metric("MQL + Responded Leads", len(mql_r))
        if mql_r.empty:
            st.info("No MQL + Responded leads found.")
        else:
            total_mr = len(mql_r)
            filled_e = [int(mql_r[f].notna().sum()) for f in EXTENDED_FIELDS]
            efc = st.columns(min(5, len(EXTENDED_LABELS)))
            for i, (lbl, f) in enumerate(zip(EXTENDED_LABELS, filled_e)):
                efc[i % 5].metric(lbl, f"{pct(f, total_mr):.1f}%", f"{total_mr-f} missing")
            st.markdown("---")
            st.markdown(progress_html(EXTENDED_LABELS, filled_e, [total_mr]*len(EXTENDED_LABELS)),
                        unsafe_allow_html=True)
            st.markdown("---")
            show_mql = mql_r[["name","owner"] + [f for f in EXTENDED_FIELDS if f in mql_r.columns] + ["missing_basic"]]
            st.dataframe(style_df(show_mql), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# TAB 4 — TASK AUDIT
# ══════════════════════════════════════════════
with tabs[4]:
    st.markdown("### ✅ Task Creation Audit")
    st.caption("'Call Not Responded' OR 'Need Another Call: Yes' → must have a Next Activity Date (task)")

    if df_all.empty:
        st.warning("No data.")
    else:
        needs_t = df_all[df_all["needs_task"]]
        has_t   = needs_t[needs_t["has_task"]]
        miss_t  = needs_t[~needs_t["has_task"]]
        nac_all = df_all[safe_str(df_all["need_another_call"]).eq("Yes")]
        nr_all  = df_all[safe_str(df_all["call_response_status"]).eq("Call Not Responded")]

        ta1,ta2,ta3,ta4,ta5 = st.columns(5)
        ta1.metric("Requiring Task",     len(needs_t))
        ta2.metric("Task Created ✅",    len(has_t),  f"{pct(len(has_t),  max(len(needs_t),1)):.1f}%")
        ta3.metric("Task Missing ⚠️",    len(miss_t), f"{pct(len(miss_t), max(len(needs_t),1)):.1f}%", delta_color="inverse")
        ta4.metric("Call Not Responded", len(nr_all))
        ta5.metric("Need Another Call",  len(nac_all))

        st.markdown("---")
        tc1, tc2 = st.columns(2)

        with tc1:
            st.markdown('<div class="section-header">Call Not Responded — Task Status</div>', unsafe_allow_html=True)
            nr_h = int(nr_all["has_task"].sum())
            nr_m = len(nr_all) - nr_h
            st.plotly_chart(
                make_donut(["Has Task","Missing Task"], [nr_h, nr_m], ["#3fb950","#f85149"]),
                use_container_width=True, key="task_donut_not_responded",
            )

        with tc2:
            st.markdown('<div class="section-header">Need Another Call: Yes — Task Status</div>', unsafe_allow_html=True)
            nac_h = int(nac_all["has_task"].sum())
            nac_m = len(nac_all) - nac_h
            st.plotly_chart(
                make_donut(["Has Task","Missing Task"], [nac_h, nac_m], ["#3fb950","#f85149"]),
                use_container_width=True, key="task_donut_need_another_call",
            )

        st.markdown("---")
        st.markdown('<div class="section-header">Task Gap by Agent</div>', unsafe_allow_html=True)
        if not needs_t.empty:
            ag_tsk = (
                needs_t.groupby("owner")
                .agg(required=("has_task","count"), created=("has_task","sum"))
                .reset_index()
            )
            ag_tsk["missing"] = ag_tsk["required"] - ag_tsk["created"]
            ag_tsk = ag_tsk.sort_values("missing", ascending=False)
            fig_tk = go.Figure()
            fig_tk.add_trace(go.Bar(y=ag_tsk["owner"], x=ag_tsk["created"], name="Created",
                                    orientation="h", marker_color="#3fb950"))
            fig_tk.add_trace(go.Bar(y=ag_tsk["owner"], x=ag_tsk["missing"], name="Missing",
                                    orientation="h", marker_color="#f85149"))
            fig_tk.update_layout(**PLOTLY_THEME, barmode="stack",
                                 xaxis=dict(gridcolor="#21262d"),
                                 yaxis=dict(tickfont=dict(color="#e6edf3"), autorange="reversed"),
                                 height=max(200, len(ag_tsk)*52))
            st.plotly_chart(fig_tk, use_container_width=True, key="task_agent_stacked_bar")

        st.markdown("---")
        st.markdown('<div class="section-header">Leads Requiring Task — Detail Table</div>', unsafe_allow_html=True)
        tf1, tf2, tf3, tf4 = st.columns(4)
        t_rsn  = tf1.selectbox("Reason", ["All","Call Not Responded","Need Another Call: Yes"], key="task_reason_selectbox")
        t_own  = tf2.selectbox("Agent",  ["All"] + sorted(df_all["owner"].dropna().unique().tolist()), key="task_owner_selectbox")
        t_stat = tf3.selectbox("Status", ["All","Missing","Created"], key="task_status_selectbox")
        t_srch = tf4.text_input("Search",                                                       key="task_search_text")

        base_tk = needs_t.copy()
        if t_rsn == "Call Not Responded":     base_tk = base_tk[safe_str(base_tk["call_response_status"]).eq("Call Not Responded")]
        if t_rsn == "Need Another Call: Yes": base_tk = base_tk[safe_str(base_tk["need_another_call"]).eq("Yes")]
        if t_own != "All":                     base_tk = base_tk[base_tk["owner"] == t_own]
        if t_stat == "Missing":                base_tk = base_tk[~base_tk["has_task"]]
        if t_stat == "Created":                base_tk = base_tk[base_tk["has_task"]]
        if t_srch:
            mask    = base_tk.apply(lambda r: r.astype(str).str.contains(t_srch, case=False).any(), axis=1)
            base_tk = base_tk[mask]

        disp_tk = base_tk[["name","owner","call_response_status","need_another_call",
                            "next_activity_date","has_task","contact_stage"]].copy()
        disp_tk["next_activity_date"] = disp_tk["next_activity_date"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
        disp_tk["has_task"] = disp_tk["has_task"].map({True:"✅ Yes", False:"❌ Missing"})
        st.markdown(f"**{len(disp_tk)} leads**")
        st.dataframe(style_df(disp_tk), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# TAB 5 — LEAD QUALIFICATION
# ══════════════════════════════════════════════
with tabs[5]:
    st.markdown("### 🏷️ Lead Qualification")

    if df_all.empty:
        st.warning("No data.")
    else:
        lq_mql  = df_all[safe_str(df_all["lead_category"]).str.upper().eq("MQL")]
        lq_mul  = df_all[safe_str(df_all["lead_category"]).str.upper().eq("MUL")]
        lq_none = df_all[df_all["lead_category"].isna()]
        mul_nd  = lq_mul[lq_mul["disq_reason"].isna()]
        resp_lq = df_all[safe_str(df_all["call_response_status"]).eq("Call Responded")]
        resp_nc = resp_lq[resp_lq["lead_category"].isna()]
        raw_lq  = df_new_biz[
            (df_new_biz["contact_stage"].isna() | safe_str(df_new_biz["contact_stage"]).eq("Raw Lead"))
        ] if not df_new_biz.empty else pd.DataFrame()

        lq1,lq2,lq3,lq4,lq5,lq6 = st.columns(6)
        lq1.metric("MQL",                  len(lq_mql),  f"{pct(len(lq_mql),  len(df_all)):.1f}%")
        lq2.metric("MUL",                  len(lq_mul),  f"{pct(len(lq_mul),  len(df_all)):.1f}%")
        lq3.metric("Not Categorized",      len(lq_none), f"{pct(len(lq_none), len(df_all)):.1f}%")
        lq4.metric("MUL Missing Disq ⚠️",  len(mul_nd),  f"{pct(len(mul_nd),  max(len(lq_mul),1)):.1f}% of MUL", delta_color="inverse")
        lq5.metric("Responded → No Cat ⚠️",len(resp_nc), f"{pct(len(resp_nc), max(len(resp_lq),1)):.1f}%", delta_color="inverse")
        lq6.metric(f"Raw Stage ({biz_start}–{biz_end}h)", len(raw_lq))

        st.markdown("---")
        lqc1, lqc2 = st.columns(2)

        with lqc1:
            st.markdown('<div class="section-header">MQL vs MUL Distribution</div>', unsafe_allow_html=True)
            cat_v2 = df_all["lead_category"].fillna("Not Set").str.strip().value_counts()
            st.plotly_chart(
                make_donut(cat_v2.index.tolist(), cat_v2.values.tolist(),
                           [CAT_COLORS.get(l,"#8b949e") for l in cat_v2.index]),
                use_container_width=True, key="lq_mql_mul_donut",
            )

        with lqc2:
            st.markdown('<div class="section-header">MUL Disqualification Reasons</div>', unsafe_allow_html=True)
            if not lq_mul.empty:
                dq_vc  = lq_mul["disq_reason"].fillna("NOT SET").str.strip().value_counts()
                dq_clrs = ["#f85149" if l=="NOT SET" else "#58a6ff" for l in dq_vc.index]
                st.plotly_chart(
                    make_bar_h(dq_vc.index.tolist(), dq_vc.values.tolist(), dq_clrs, pct_of=len(lq_mul)),
                    use_container_width=True, key="lq_disq_reasons_bar",
                )
            else:
                st.info("No MUL leads found.")

        st.markdown("---")
        st.markdown('<div class="section-header">Lead Category Completion — Responded Calls by Agent</div>',
                    unsafe_allow_html=True)
        if not resp_lq.empty:
            ag_cat = (
                resp_lq.assign(has_cat=resp_lq["lead_category"].notna())
                .groupby("owner").agg(filled=("has_cat","sum"), total=("has_cat","count"))
                .reset_index().sort_values("filled", ascending=False)
            )
            st.markdown(progress_html(ag_cat["owner"].tolist(),
                                      ag_cat["filled"].astype(int).tolist(),
                                      ag_cat["total"].astype(int).tolist()),
                        unsafe_allow_html=True)

        st.markdown("---")
        if len(raw_lq) == 0:
            st.success(f"✅ No leads in Raw/unworked stage during {biz_start}:00–{biz_end}:00 window.")
        else:
            st.warning(f"⚠️ {len(raw_lq)} leads still in Raw stage during {biz_start}:00–{biz_end}:00.")
            st.dataframe(style_df(raw_lq[["name","owner","create_date","contact_stage"]]),
                         use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown('<div class="section-header">Qualification Detail Table</div>', unsafe_allow_html=True)
        lf1, lf2, lf3, lf4 = st.columns(4)
        lq_src  = lf1.selectbox("Source",      ["All","New Leads","Activity Leads"], key="lq_source_selectbox")
        lq_cat  = lf2.selectbox("Category",    ["All","MQL","MUL","Not Set"],       key="lq_cat_selectbox")
        lq_dq   = lf3.selectbox("Disq Reason", ["All","Missing","Filled"],           key="lq_disq_selectbox")
        lq_srch = lf4.text_input("Search",                                           key="lq_search_text")

        base_lq = {"All":df_all,"New Leads":df_new,"Activity Leads":df_act}[lq_src]
        if not base_lq.empty:
            if lq_cat == "MQL":      base_lq = base_lq[safe_str(base_lq["lead_category"]).str.upper().eq("MQL")]
            elif lq_cat == "MUL":    base_lq = base_lq[safe_str(base_lq["lead_category"]).str.upper().eq("MUL")]
            elif lq_cat == "Not Set":base_lq = base_lq[base_lq["lead_category"].isna()]
            if lq_dq == "Missing":   base_lq = base_lq[base_lq["disq_reason"].isna()]
            elif lq_dq == "Filled":  base_lq = base_lq[base_lq["disq_reason"].notna()]
            if lq_srch:
                mask    = base_lq.apply(lambda r: r.astype(str).str.contains(lq_srch, case=False).any(), axis=1)
                base_lq = base_lq[mask]
            show_lq = base_lq[["name","owner","call_response_status","lead_category",
                                "disq_reason","contact_stage","call_outcome"]].copy()
            st.markdown(f"**{len(show_lq)} leads**")
            st.dataframe(style_df(show_lq), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# TAB 6 — BY AGENT
# ══════════════════════════════════════════════
with tabs[6]:
    st.markdown("### 👤 Agent Performance")

    if df_all.empty:
        st.warning("No data.")
    else:
        ag_src  = st.selectbox("Data Source", ["All","New Leads","Activity Leads"], key="ag_source_selectbox")
        ag_base = {"All":df_all,"New Leads":df_new,"Activity Leads":df_act}[ag_src]
        owners  = sorted(ag_base["owner"].dropna().unique().tolist())

        rows_ag = []
        for owner in owners:
            ol    = ag_base[ag_base["owner"] == owner]
            resp  = ol[safe_str(ol["call_response_status"]).eq("Call Responded")]
            rt_v2 = ol["rt_mins"].dropna()
            mql_c = int(safe_str(ol["lead_category"]).str.upper().eq("MQL").sum())
            mul_c = int(safe_str(ol["lead_category"]).str.upper().eq("MUL").sum())
            l_mis = int(resp["preferred_language"].isna().sum())
            t_mis = int(ol["task_missing"].sum())
            bf    = (len(BASIC_FIELDS)*len(ol) - int(ol["missing_basic"].sum())) / max(len(BASIC_FIELDS)*len(ol),1)*100
            avg_r = float(rt_v2.mean()) if len(rt_v2) else None
            rows_ag.append({
                "Agent":        owner,
                "Total Leads":  len(ol),
                "Responded":    len(resp),
                "Responded %":  pct(len(resp), len(ol)),
                "MQL":          mql_c,
                "MUL":          mul_c,
                "Lang Missing": l_mis,
                "Task Missing": t_mis,
                "Avg RT raw":   avg_r,
                "Avg RT":       fmt_mins(avg_r),
                "Basic Fill %": round(bf, 1),
            })

        ag_grid = st.columns(min(3, max(len(rows_ag), 1)))
        for i, row in enumerate(rows_ag):
            rt_c = "#3fb950" if (row["Avg RT raw"] or 999) < 5 else "#d29922" if (row["Avg RT raw"] or 999) < 10 else "#f85149"
            fp   = row["Basic Fill %"]
            fp_c = "#3fb950" if fp >= 90 else "#d29922" if fp >= 70 else "#f85149"
            lm_c = "#f85149" if row["Lang Missing"] > 0 else "#3fb950"
            tm_c = "#f85149" if row["Task Missing"] > 0 else "#3fb950"
            with ag_grid[i % 3]:
                st.markdown(f"""
                <div style="background:#161b22;border:1px solid #21262d;border-radius:12px;
                            padding:16px;margin-bottom:12px">
                  <div style="font-size:14px;font-weight:700;color:#e6edf3;margin-bottom:10px;
                              border-bottom:1px solid #21262d;padding-bottom:8px">{row['Agent']}</div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:12px;color:#8b949e">Total Leads</span>
                    <span style="font-size:12px;font-weight:600;color:#e6edf3">{row['Total Leads']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:12px;color:#8b949e">Responded</span>
                    <span style="font-size:12px;font-weight:600;color:#3fb950">{row['Responded']} ({row['Responded %']:.1f}%)</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:12px;color:#8b949e">MQL / MUL</span>
                    <span style="font-size:12px;font-weight:600;color:#e6edf3">{row['MQL']} / {row['MUL']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:12px;color:#8b949e">Avg RT</span>
                    <span style="font-size:12px;font-weight:600;color:{rt_c}">{row['Avg RT']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:12px;color:#8b949e">Lang Missing</span>
                    <span style="font-size:12px;font-weight:600;color:{lm_c}">{row['Lang Missing']}</span></div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:12px;color:#8b949e">Task Missing</span>
                    <span style="font-size:12px;font-weight:600;color:{tm_c}">{row['Task Missing']}</span></div>
                  <div style="display:flex;justify-content:space-between">
                    <span style="font-size:12px;color:#8b949e">Field Fill %</span>
                    <span style="font-size:12px;font-weight:600;color:{fp_c}">{fp:.1f}%</span></div>
                </div>""", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown('<div class="section-header">Responded % by Agent</div>', unsafe_allow_html=True)
        rp_vals = [r["Responded %"] for r in rows_ag]
        rp_clrs = ["#3fb950" if v >= 70 else "#d29922" if v >= 50 else "#f85149" for v in rp_vals]
        fig_ag  = go.Figure(go.Bar(
            y=[r["Agent"] for r in rows_ag], x=rp_vals, orientation="h",
            marker_color=rp_clrs,
            text=[f"{v:.1f}%" for v in rp_vals], textposition="outside",
            textfont=dict(color="#e6edf3"),
        ))
        fig_ag.update_layout(**PLOTLY_THEME,
                             xaxis=dict(range=[0,110], gridcolor="#21262d"),
                             yaxis=dict(tickfont=dict(color="#e6edf3"), autorange="reversed"),
                             height=max(200, len(rows_ag)*52))
        st.plotly_chart(fig_ag, use_container_width=True, key="ag_responded_pct_bar")

        st.markdown("---")
        disp_ag_cols = ["Agent","Total Leads","Responded","Responded %","MQL","MUL",
                        "Lang Missing","Task Missing","Avg RT","Basic Fill %"]
        st.dataframe(pd.DataFrame(rows_ag)[disp_ag_cols], use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# TAB 7 — ALL LEADS
# ══════════════════════════════════════════════
with tabs[7]:
    st.markdown("### 📑 All Leads — Master Table")

    if df_all.empty:
        st.warning("No data.")
    else:
        all_stages = sorted(df_all["contact_stage"].dropna().unique().tolist())
        all_owners = sorted(df_all["owner"].dropna().unique().tolist())

        al1, al2, al3, al4, al5 = st.columns(5)
        al_src   = al1.selectbox("Source",      ["All","New Leads","New Leads (Biz Hrs)","Activity Leads"], key="al_source_selectbox")
        al_call  = al2.selectbox("Call Status", ["All","Call Responded","Call Not Responded","Invalid Number"], key="al_call_selectbox")
        al_stage = al3.selectbox("Stage",       ["All"] + all_stages, key="al_stage_selectbox")
        al_own   = al4.selectbox("Agent",       ["All"] + all_owners,  key="al_owner_selectbox")
        al_cat   = al5.selectbox("Category",    ["All","MQL","MUL","Not Set"], key="al_cat_selectbox")
        al_srch  = st.text_input("🔍 Search name, city, source…", key="al_search_text")

        base_al = {
            "All":                 df_all,
            "New Leads":           df_new,
            "New Leads (Biz Hrs)": df_new_biz,
            "Activity Leads":      df_act,
        }[al_src]

        if not base_al.empty:
            if al_call != "All":    base_al = base_al[safe_str(base_al["call_response_status"]).eq(al_call)]
            if al_stage != "All":   base_al = base_al[safe_str(base_al["contact_stage"]).eq(al_stage)]
            if al_own != "All":     base_al = base_al[base_al["owner"] == al_own]
            if al_cat == "MQL":     base_al = base_al[safe_str(base_al["lead_category"]).str.upper().eq("MQL")]
            elif al_cat == "MUL":   base_al = base_al[safe_str(base_al["lead_category"]).str.upper().eq("MUL")]
            elif al_cat == "Not Set": base_al = base_al[base_al["lead_category"].isna()]
            if al_srch:
                mask    = base_al.apply(lambda r: r.astype(str).str.contains(al_srch, case=False).any(), axis=1)
                base_al = base_al[mask]

            st.markdown(f"**{len(base_al)} leads matching filters**")

            disp_cols = ["name","owner","create_date","last_activity","call_response_status",
                         "contact_stage","lead_category","need_another_call","preferred_language",
                         "rt_mins","disq_reason","product_interest","call_outcome",
                         "contact_source","city","state","has_task","task_missing","missing_basic"]
            disp_al = base_al[[c for c in disp_cols if c in base_al.columns]].copy()
            disp_al["create_date"]   = disp_al["create_date"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
            disp_al["last_activity"] = disp_al["last_activity"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
            disp_al["rt_mins"]       = disp_al["rt_mins"].apply(fmt_mins)
            disp_al["has_task"]      = disp_al["has_task"].map({True:"✅", False:"❌"})
            disp_al["task_missing"]  = disp_al["task_missing"].map({True:"⚠️ Yes", False:"—", None:"—"}).fillna("—")
            disp_al.columns = [c.replace("_"," ").title() for c in disp_al.columns]

            st.dataframe(style_df(disp_al), use_container_width=True, hide_index=True, height=520)

            csv_buf = io.StringIO()
            disp_al.to_csv(csv_buf, index=False)
            st.download_button(
                "⬇️ Download filtered data as CSV",
                data=csv_buf.getvalue(),
                file_name="filtered_leads.csv",
                mime="text/csv",
                key="al_csv_download_button",
            )
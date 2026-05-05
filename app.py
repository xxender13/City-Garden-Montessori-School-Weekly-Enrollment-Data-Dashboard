"""
City Garden Montessori School — Weekly Dashboard
Streamlit App  ·  Upload your summary .xlsx → instant interactive HTML dashboard
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
import warnings
import io
warnings.filterwarnings("ignore")

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CGMS Weekly Dashboard",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── GLOBAL CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');

/* ── reset streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 !important; max-width: 100% !important; }
.stApp { background: #080C16; }

/* ── sidebar ── */
[data-testid="stSidebar"] {
    background: #0F1525 !important;
    border-right: 1px solid #1E2740;
}
[data-testid="stSidebar"] * { color: #C5CCDF !important; font-family: 'DM Sans', sans-serif !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
    color: #E8EEFF !important; font-family: 'Lora', serif !important;
}
[data-testid="stSidebarContent"] { padding: 24px 20px; }

/* ── file uploader ── */
[data-testid="stFileUploader"] {
    background: #111827;
    border: 2px dashed #2A3755;
    border-radius: 12px;
    padding: 8px;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #4F8EF7; }
[data-testid="stFileUploader"] label { color: #7A85A3 !important; font-size: 0.82rem !important; }

/* ── buttons ── */
.stButton > button {
    background: #4F8EF7 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    padding: 8px 20px !important;
    width: 100%;
}
.stButton > button:hover { background: #3A7AE8 !important; }

/* ── main page text ── */
h1, h2, h3 { font-family: 'Lora', serif !important; color: #E8EEFF !important; }
p, li, span { font-family: 'DM Sans', sans-serif !important; color: #C5CCDF !important; }

/* ── landing card ── */
.landing-wrap {
    min-height: 80vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 40px 20px;
}
.landing-card {
    background: #0F1525;
    border: 1px solid #1E2740;
    border-radius: 20px;
    padding: 52px 56px;
    max-width: 820px;
    width: 100%;
}
.school-logo {
    font-size: 3rem;
    text-align: center;
    margin-bottom: 12px;
}
.school-name {
    font-family: 'Lora', serif;
    font-size: 2rem;
    font-weight: 700;
    color: #E8EEFF;
    text-align: center;
    margin-bottom: 4px;
}
.school-tagline {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.92rem;
    color: #5A6280;
    text-align: center;
    margin-bottom: 40px;
    letter-spacing: 1px;
    text-transform: uppercase;
}
.how-it-works {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 20px;
    margin-bottom: 40px;
}
.step-card {
    background: #151D30;
    border: 1px solid #1E2740;
    border-radius: 12px;
    padding: 20px 16px;
    text-align: center;
}
.step-num {
    width: 32px; height: 32px;
    background: #4F8EF7;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-family: 'Lora', serif;
    font-weight: 700;
    color: white;
    font-size: 0.95rem;
    margin-bottom: 10px;
}
.step-title {
    font-family: 'Lora', serif;
    font-size: 0.92rem;
    font-weight: 600;
    color: #E8EEFF;
    margin-bottom: 6px;
}
.step-desc {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.78rem;
    color: #5A6280;
    line-height: 1.5;
}
.divider {
    border: none;
    border-top: 1px solid #1E2740;
    margin: 28px 0;
}
.what-inside {
    font-family: 'Lora', serif;
    font-size: 1rem;
    font-weight: 600;
    color: #E8EEFF;
    margin-bottom: 16px;
    text-align: center;
}
.feature-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 10px;
    margin-bottom: 32px;
}
.feature-item {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    background: #151D30;
    border-radius: 8px;
    padding: 12px 14px;
}
.feature-icon { font-size: 1.1rem; flex-shrink: 0; margin-top: 1px; }
.feature-text { font-family: 'DM Sans', sans-serif; font-size: 0.80rem; color: #7A85A3; line-height: 1.4; }
.feature-text b { color: #C5CCDF; }
.note-box {
    background: rgba(79,142,247,0.08);
    border: 1px solid rgba(79,142,247,0.2);
    border-radius: 10px;
    padding: 14px 18px;
    font-family: 'DM Sans', sans-serif;
    font-size: 0.78rem;
    color: #7A85A3;
    line-height: 1.6;
    text-align: center;
}
.note-box b { color: #4F8EF7; }

/* ── processing banner ── */
.proc-banner {
    background: #0F1525;
    border: 1px solid #1E2740;
    border-radius: 12px;
    padding: 24px 28px;
    margin: 16px 0;
    font-family: 'DM Sans', sans-serif;
    font-size: 0.88rem;
    color: #7A85A3;
}

/* ── iframe wrapper ── */
.dashboard-frame-wrap {
    width: 100%;
    background: #080C16;
    padding: 0;
}
</style>
""", unsafe_allow_html=True)


# ── DESIGN TOKENS (shared with chart builder) ─────────────────────────────────
BG     = "#0B0F1A"; CARD   = "#131929"; PANEL  = "#1A2236"; BORDER = "#252D42"
T1     = "#E8EEFF"; T2     = "#7A85A3"; T3     = "#4A5270"
BLUE   = "#4F8EF7"; TEAL   = "#3ECFBF"; GOLD   = "#F5C842"; RED    = "#F45B69"
PURPLE = "#9B72F2"; GREEN  = "#43D9A2"; ORANGE = "#FF8C5A"
RACE_PAL = [BLUE, TEAL, GOLD, RED, PURPLE, GREEN]
FONT_T   = "Lora, Georgia, serif"
FONT_B   = "'DM Sans', 'Trebuchet MS', sans-serif"
ATT_TARGET = 93.0

COMMON = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family=FONT_B, color=T2, size=11),
    margin=dict(t=48, b=32, l=44, r=24),
    hoverlabel=dict(bgcolor="#1E2845", bordercolor=BORDER,
                    font=dict(color=T1, family=FONT_B, size=12)),
)


# ── DATA PARSER ───────────────────────────────────────────────────────────────
def parse_excel(uploaded_file):
    df = pd.read_excel(uploaded_file,sheet_name="summary", header=None)

    def cell(r, c):
        try:
            v = df.iloc[r, c]
            return v if pd.notna(v) else None
        except: return None

    def num(r, c, default=0):
        v = cell(r, c)
        try:    return float(v)
        except: return default

    def pct(r, c, default=0):
        v = num(r, c, default)
        return round(v * 100 if (v != 0 and v <= 1) else v, 2)

    data = {}
    data["week_date"]   = str(cell(1, 1))[:10]
    data["school_year"] = str(cell(3, 0))

    GRADE_ROWS = {
        "Grade 1":5,"Grade 2":6,"Grade 3":7,"Grade 4":8,
        "Grade 5":9,"Grade 6":10,"Grade 7":11,"Grade 8":12,
        "PS":14,"PK":15,"K":16,
    }
    data["grades"]        = list(GRADE_ROWS.keys())
    data["enroll_actual"] = [num(r,1) for r in GRADE_ROWS.values()]
    data["enroll_budget"] = [num(r,2) for r in GRADE_ROWS.values()]
    data["enroll_var"]    = [num(r,3) for r in GRADE_ROWS.values()]
    data["att_week"]      = [pct(r,4) for r in GRADE_ROWS.values()]
    data["att_ytd"]       = [pct(r,5) for r in GRADE_ROWS.values()]
    data["att_ada"]       = [num(r,6) for r in GRADE_ROWS.values()]

    data["total_enrolled"]   = num(18,1); data["total_budget_enr"] = num(18,2)
    data["total_var_enr"]    = num(18,3); data["att_week_total"]   = pct(18,4)
    data["att_ytd_total"]    = pct(18,5); data["ada_total"]        = num(18,6)
    data["charter_enrolled"] = num(20,1); data["charter_att_week"] = pct(20,4)
    data["charter_att_ytd"]  = pct(20,5); data["charter_ada"]      = num(20,6)
    data["charter_9090"]     = pct(5,7)

    RACE_CODE = {"A":"Asian","B":"Black","H":"Hispanic","I":"Indigenous","M":"Multiracial","W":"White"}
    data["race_labels"] = [RACE_CODE.get(str(cell(r,0)), str(cell(r,0))) for r in range(29,35)]
    data["race_ws_n"]   = [num(r,7) for r in range(29,35)]

    FRL_SEGS = {"Charter":37,"Preschool":38,"Whole School":39,"EAEC":40,"ECEC":41}
    ENR_TOT  = {"Charter":511,"Preschool":91,"Whole School":602,"EAEC":419,"ECEC":183}
    data["frl_segs"]  = list(FRL_SEGS.keys())
    data["frl_enr"]   = ENR_TOT
    data["frl_free"]  = {s:num(r,1) for s,r in FRL_SEGS.items()}
    data["frl_red"]   = {s:num(r,2) for s,r in FRL_SEGS.items()}
    data["frl_total"] = {s:num(r,3) for s,r in FRL_SEGS.items()}
    data["frl_pct"]   = {s:pct(r,4) for s,r in FRL_SEGS.items()}

    data["iep"]  = {"Charter":{"n":num(44,1),"pct":pct(44,2)},
                    "Preschool":{"n":num(45,1),"pct":pct(45,2)},
                    "Whole School":{"n":num(46,1),"pct":pct(46,2)}}
    data["s504"] = {"Charter":{"n":num(44,4),"pct":pct(44,5)},
                    "Preschool":{"n":num(45,4),"pct":pct(45,5)},
                    "Whole School":{"n":num(46,4),"pct":pct(46,5)}}
    data["ell"]  = {"Charter":{"n":num(49,1),"pct":pct(49,2)},
                    "Preschool":{"n":num(50,1),"pct":pct(50,2)},
                    "Whole School":{"n":num(51,1),"pct":pct(51,2)}}

    data["iss_week"] = num(54,1); data["iss_ytd"] = num(54,2)
    data["oss_week"] = num(55,1); data["oss_ytd"] = num(55,2)
    return data


# ── CHART BUILDERS ────────────────────────────────────────────────────────────
def att_color(v):
    if v == 0: return T3
    return GREEN if v >= ATT_TARGET else (GOLD if v >= 90 else RED)

def fig_html(fig, div_id, h=270):
    return pio.to_html(fig, full_html=False, include_plotlyjs=False,
                       div_id=div_id, default_height=f"{h}px",
                       config={"displayModeBar": False, "responsive": True})

def build_enrollment(d):
    g = [x.replace("Grade ","Gr. ") for x in d["grades"]]
    ymax = max(max(d["enroll_actual"]), max(d["enroll_budget"])) * 1.22
    fig = go.Figure()
    fig.add_trace(go.Bar(x=g, y=d["enroll_actual"], name="Enrolled",
        marker_color=BLUE, marker_line_width=0,
        text=[f"{int(v)}" for v in d["enroll_actual"]],
        textposition="outside", textfont=dict(color=T1, size=9, family=FONT_B),
        offsetgroup=1, hovertemplate="<b>%{x}</b><br>Enrolled: %{y}<extra></extra>"))
    fig.add_trace(go.Bar(x=g, y=d["enroll_budget"], name="Budget",
        marker_color=PANEL, marker_line_color=BLUE, marker_line_width=1.5,
        offsetgroup=2, hovertemplate="<b>%{x}</b><br>Budget: %{y}<extra></extra>"))
    fig.update_layout(**COMMON, barmode="group",
        title=dict(text="Enrollment vs. Budget by Grade", font=dict(color=T1,size=13,family=FONT_T), x=0),
        legend=dict(orientation="h", x=1, xanchor="right", y=1.15, bgcolor="rgba(0,0,0,0)", font=dict(size=10)),
        xaxis=dict(gridcolor=BORDER, showgrid=False, tickfont=dict(size=9, color=T2)),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(size=9, color=T2), range=[0, ymax]))
    return fig_html(fig, "chart-enroll")

def build_attendance(d):
    valid = [(g, aw, ay) for g, aw, ay in zip(d["grades"], d["att_week"], d["att_ytd"]) if aw > 0]
    if not valid: return "<p style='color:#5A6280;padding:20px'>No attendance data</p>"
    gv = [x[0].replace("Grade ","Gr. ") for x in valid]
    aw = [x[1] for x in valid]; ay = [x[2] for x in valid]
    ymin = max(80, min(aw + ay) - 4)
    fig = go.Figure()
    fig.add_hrect(y0=ATT_TARGET, y1=103, fillcolor="rgba(67,217,162,0.04)", line_width=0, layer="below")
    fig.add_hline(y=ATT_TARGET, line_color=RED, line_dash="dash", line_width=1.5,
                  annotation=dict(text=f"  {ATT_TARGET}% Target", font=dict(color=RED, size=10), x=1, xanchor="left"))
    fig.add_trace(go.Bar(x=gv, y=aw, name="This Week",
        marker_color=[att_color(v) for v in aw], marker_line_width=0,
        text=[f"{v:.1f}%" for v in aw], textposition="outside", textfont=dict(size=9, color=T1),
        hovertemplate="<b>%{x}</b><br>This Week: %{y:.1f}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=gv, y=ay, name="YTD Avg",
        mode="lines+markers", line=dict(color=PURPLE, width=2.5, dash="dot"),
        marker=dict(size=7, color=PURPLE, line=dict(color=BG, width=1.5)),
        hovertemplate="<b>%{x}</b><br>YTD: %{y:.1f}%<extra></extra>"))
    fig.update_layout(**COMMON,
        title=dict(text="Attendance — This Week vs. YTD", font=dict(color=T1,size=13,family=FONT_T), x=0),
        legend=dict(orientation="h", x=1, xanchor="right", y=1.15, bgcolor="rgba(0,0,0,0)", font=dict(size=10)),
        xaxis=dict(gridcolor=BORDER, showgrid=False, tickfont=dict(size=9, color=T2)),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(size=9, color=T2), range=[ymin, 103], ticksuffix="%"))
    return fig_html(fig, "chart-att")

def build_race(d):
    fig = go.Figure(go.Pie(
        labels=d["race_labels"], values=d["race_ws_n"], hole=0.60,
        marker=dict(colors=RACE_PAL, line=dict(color=BG, width=3)),
        textinfo="label+percent", textfont=dict(size=11, family=FONT_B, color=T1),
        hovertemplate="<b>%{label}</b><br>%{value} students — %{percent}<extra></extra>",
        showlegend=False, direction="clockwise", sort=True,
        pull=[0.03 if v == max(d["race_ws_n"]) else 0 for v in d["race_ws_n"]]))
    fig.add_annotation(x=0.5, y=0.5, showarrow=False, align="center", xref="paper", yref="paper",
        text=f"<b style='font-size:22px'>{int(sum(d['race_ws_n']))}</b><br><span style='font-size:10px;color:{T2}'>Students</span>",
        font=dict(family=FONT_T, color=T1))
    rc = {**COMMON, "margin": dict(t=48, b=12, l=12, r=12)}
    fig.update_layout(**rc, title=dict(text="Race / Ethnicity — Whole School", font=dict(color=T1,size=13,family=FONT_T), x=0))
    return fig_html(fig, "chart-race", h=300)

def build_frl(d):
    segs  = d["frl_segs"]
    enr   = d["frl_enr"]
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Free",    x=segs, y=[d["frl_free"][s] for s in segs], marker_color=RED,    marker_line_width=0, hovertemplate="<b>%{x}</b><br>Free: %{y}<extra></extra>"))
    fig.add_trace(go.Bar(name="Reduced", x=segs, y=[d["frl_red"][s]  for s in segs], marker_color=GOLD,   marker_line_width=0, hovertemplate="<b>%{x}</b><br>Reduced: %{y}<extra></extra>"))
    fig.add_trace(go.Bar(name="Paid",    x=segs, y=[max(0, enr[s]-d["frl_total"][s]) for s in segs], marker_color=BORDER, marker_line_width=0, hovertemplate="<b>%{x}</b><br>Paid: %{y}<extra></extra>"))
    for s in segs:
        fig.add_annotation(x=s, y=enr[s]+16, showarrow=False,
            text=f"<b>{d['frl_pct'][s]:.1f}%</b>", font=dict(size=10, color=T1, family=FONT_B))
    fig.update_layout(**COMMON, barmode="stack",
        title=dict(text="Free & Reduced Lunch (FRL) by Segment", font=dict(color=T1,size=13,family=FONT_T), x=0),
        legend=dict(orientation="h", x=1, xanchor="right", y=1.15, bgcolor="rgba(0,0,0,0)", font=dict(size=10)),
        xaxis=dict(gridcolor=BORDER, showgrid=False, tickfont=dict(size=9, color=T2)),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(size=9, color=T2)))
    return fig_html(fig, "chart-frl")

def build_programs(d):
    iep  = d["iep"];  s504 = d["s504"];  ell = d["ell"]
    segs   = ["Charter", "Whole School"]
    colors = [PURPLE, TEAL, GOLD]
    progs  = ["IEPs", "504 Plans", "ELL"]
    pdata  = [iep, s504, ell]
    fig = make_subplots(rows=1, cols=3, subplot_titles=progs, horizontal_spacing=0.10)
    for ci, (prog, pd_item, color) in enumerate(zip(progs, pdata, colors), start=1):
        for seg in segs:
            n   = pd_item[seg]["n"]
            p   = pd_item[seg]["pct"]
            tot = d["total_enrolled"] if seg == "Whole School" else d["charter_enrolled"]
            fig.add_trace(go.Bar(y=[seg], x=[n], orientation="h",
                marker_color=color, marker_line_width=0, showlegend=False,
                text=f"  {int(n)} ({p:.1f}%)", textposition="inside",
                textfont=dict(size=11, color=T1, family=FONT_B), width=0.5,
                hovertemplate=f"<b>{seg}</b><br>{prog}: {int(n)} ({p:.1f}%)<extra></extra>"), row=1, col=ci)
            fig.add_trace(go.Bar(y=[seg], x=[tot - n], orientation="h",
                marker_color=BORDER, marker_line_width=0, showlegend=False, width=0.5,
                hovertemplate=f"<b>{seg}</b><br>Other: {int(tot-n)}<extra></extra>"), row=1, col=ci)
        fig.layout.annotations[ci-1].update(font=dict(color=T1, size=12, family=FONT_T))
    cm = {k:v for k,v in COMMON.items() if k not in ("paper_bgcolor","plot_bgcolor")}
    AXS = dict(gridcolor=BORDER, linecolor=BORDER, zerolinecolor=BORDER)
    fig.update_layout(**cm, barmode="stack",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        title=dict(text="Special Programs — Rate Comparison (Charter vs. Whole School)", font=dict(color=T1,size=13,family=FONT_T), x=0),
        yaxis= dict(**AXS, showgrid=False, tickfont=dict(size=11, color=T1)),
        yaxis2=dict(**AXS, showgrid=False, tickfont=dict(size=11, color=T1)),
        yaxis3=dict(**AXS, showgrid=False, tickfont=dict(size=11, color=T1)),
        xaxis= dict(**AXS, showgrid=True, tickfont=dict(size=9, color=T2)),
        xaxis2=dict(**AXS, showgrid=True, tickfont=dict(size=9, color=T2)),
        xaxis3=dict(**AXS, showgrid=True, tickfont=dict(size=9, color=T2)))
    return fig_html(fig, "chart-prog", h=220)

def build_discipline(d):
    iss_ytd = d["iss_ytd"]; oss_ytd = d["oss_ytd"]
    iss_wk  = d["iss_week"]; oss_wk  = d["oss_week"]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=["ISS — In-School Suspension", "OSS — Out-of-School Suspension"],
        y=[iss_ytd, oss_ytd], marker_color=[RED, ORANGE], marker_line_width=0,
        text=[f"{int(iss_ytd)}", f"{int(oss_ytd)}"],
        textposition="outside", textfont=dict(size=22, color=T1, family=FONT_T, weight="bold"),
        width=0.35, name="YTD",
        hovertemplate="<b>%{x}</b><br>YTD Incidents: %{y}<extra></extra>"))
    fig.add_annotation(x=0.5, y=1.1, xref="paper", yref="paper", showarrow=False,
        text=f"This week: ISS = {int(iss_wk)}   ·   OSS = {int(oss_wk)}",
        font=dict(size=11, color=T2, family=FONT_B))
    fig.update_layout(**COMMON, showlegend=False,
        title=dict(text="Discipline Incidents — Year to Date", font=dict(color=T1,size=13,family=FONT_T), x=0),
        xaxis=dict(gridcolor=BORDER, showgrid=False, tickfont=dict(size=11, color=T1)),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(size=9, color=T2),
                   range=[0, max(iss_ytd, oss_ytd, 1) * 1.5]))
    return fig_html(fig, "chart-disc", h=240)

def build_variance(d):
    g      = [x.replace("Grade ","Gr. ") for x in d["grades"]]
    colors = [GREEN if v >= 0 else RED for v in d["enroll_var"]]
    fig = go.Figure()
    fig.add_vline(x=0, line_color=T2, line_dash="dash", line_width=1.5)
    fig.add_trace(go.Scatter(
        x=d["enroll_var"], y=g, mode="markers+text",
        marker=dict(color=colors, size=16, line=dict(color=BG, width=2)),
        text=[f"  {int(v):+}" for v in d["enroll_var"]],
        textposition="middle right", textfont=dict(size=11, color=T1, family=FONT_B),
        hovertemplate="<b>%{y}</b><br>Variance: %{x:+}<extra></extra>", showlegend=False))
    cm = {**COMMON, "margin": dict(t=48, b=32, l=60, r=40)}
    fig.update_layout(**cm,
        title=dict(text="Enrollment Delta vs. Budget (per Grade)", font=dict(color=T1,size=13,family=FONT_T), x=0),
        xaxis=dict(gridcolor=BORDER, zeroline=False, tickfont=dict(size=10, color=T2),
                   title=dict(text="Students vs. Budget", font=dict(size=10, color=T2))),
        yaxis=dict(gridcolor=BORDER, showgrid=False, tickfont=dict(size=10, color=T1)))
    return fig_html(fig, "chart-var", h=320)


# ── FULL HTML DASHBOARD BUILDER ───────────────────────────────────────────────
def build_dashboard_html(d):
    ch_enroll = build_enrollment(d)
    ch_att    = build_attendance(d)
    ch_race   = build_race(d)
    ch_frl    = build_frl(d)
    ch_prog   = build_programs(d)
    ch_disc   = build_discipline(d)
    ch_var    = build_variance(d)

    fmt_pct = lambda v: f"{v:.1f}%"
    fmt_n   = lambda v: f"{int(v):,}"

    def kpi(label, value, sub, color, icon):
        return f"""<div class="kpi-card">
          <div class="kpi-icon">{icon}</div>
          <div class="kpi-value" style="color:{color}">{value}</div>
          <div class="kpi-label">{label}</div>
          <div class="kpi-sub">{sub}</div>
        </div>"""

    def prog_card(title, icon, color, ws_n, ws_p, ch_n, ch_p, ps_n, ps_p, desc):
        rgb = tuple(int(color.lstrip('#')[i:i+2], 16) for i in (0,2,4))
        return f"""<div class="prog-card" style="border-left:4px solid {color}">
          <div class="prog-header">
            <div class="prog-icon" style="background:rgba{rgb+(0.15,)};color:{color}">{icon}</div>
            <div>
              <div class="prog-title">{title}</div>
              <div class="prog-desc">{desc}</div>
            </div>
          </div>
          <div class="prog-stats">
            <div class="prog-stat">
              <div class="ps-num" style="color:{color}">{int(ws_n)}</div>
              <div class="ps-lbl">Whole School</div>
              <div class="ps-pct">{ws_p:.1f}% of {int(d['total_enrolled'])}</div>
            </div>
            <div class="prog-stat">
              <div class="ps-num" style="color:{color}">{int(ch_n)}</div>
              <div class="ps-lbl">Charter K–8</div>
              <div class="ps-pct">{ch_p:.1f}% of {int(d['charter_enrolled'])}</div>
            </div>
            <div class="prog-stat">
              <div class="ps-num" style="color:{color}">{int(ps_n)}</div>
              <div class="ps-lbl">Preschool</div>
              <div class="ps-pct">{'—' if ps_n==0 else f'{ps_p:.1f}%'}</div>
            </div>
          </div>
        </div>"""

    def att_badge(v):
        if v == 0: return f'<span class="badge badge-gray">N/A</span>'
        cls = "badge-green" if v >= ATT_TARGET else ("badge-gold" if v >= 90 else "badge-red")
        return f'<span class="badge {cls}">{v:.1f}%</span>'

    grade_display = {"Grade 1":"1st","Grade 2":"2nd","Grade 3":"3rd","Grade 4":"4th",
                     "Grade 5":"5th","Grade 6":"6th","Grade 7":"7th","Grade 8":"8th",
                     "K":"Kindergarten","PK":"Pre-K","PS":"Pre-School"}

    table_rows = ""
    for i, g in enumerate(d["grades"]):
        v = int(d["enroll_var"][i])
        vc = "badge-green" if v >= 0 else "badge-red"
        table_rows += f"""<tr>
          <td><b>{grade_display.get(g,g)}</b></td>
          <td>{int(d['enroll_actual'][i])}</td>
          <td>{int(d['enroll_budget'][i])}</td>
          <td><span class="badge {vc}">{v:+}</span></td>
          <td>{att_badge(d['att_week'][i])}</td>
          <td>{att_badge(d['att_ytd'][i])}</td>
          <td style="color:{T2}">{d['att_ada'][i]:.1f}</td>
        </tr>"""

    v_total = int(d["total_var_enr"])
    table_rows += f"""<tr class="total-row">
      <td><b>TOTAL</b></td>
      <td><b>{int(d['total_enrolled'])}</b></td>
      <td><b>{int(d['total_budget_enr'])}</b></td>
      <td><span class="badge {'badge-green' if v_total>=0 else 'badge-red'}">{v_total:+}</span></td>
      <td>{att_badge(d['att_week_total'])}</td>
      <td>{att_badge(d['att_ytd_total'])}</td>
      <td style="color:{T2}"><b>{d['ada_total']:.1f}</b></td>
    </tr>"""

    frl = d["frl_pct"]
    iep = d["iep"]; s504 = d["s504"]; ell = d["ell"]

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CGMS Dashboard — {d['week_date']}</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{background:{BG};color:{T1};font-family:'DM Sans',sans-serif;}}
.page-header{{background:linear-gradient(135deg,#0D1526 0%,#172040 60%,#0D1526 100%);border-bottom:1px solid {BORDER};padding:22px 40px 18px;display:flex;align-items:flex-end;justify-content:space-between;flex-wrap:wrap;gap:12px}}
.header-left h1{{font-family:'Lora',serif;font-size:1.75rem;color:{T1};letter-spacing:.4px}}
.header-left .sub{{font-size:.82rem;color:{T2};margin-top:4px}}
.week-badge{{background:{BLUE};color:#fff;padding:5px 16px;border-radius:20px;font-size:.80rem;font-weight:700;letter-spacing:.6px}}
.sys-badge{{font-size:.70rem;color:{T3};margin-top:6px;text-align:right}}
.page-body{{padding:24px 36px;max-width:1440px;margin:0 auto}}
.section-header{{margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid {BORDER}}}
.section-header h2{{font-family:'Lora',serif;font-size:1.05rem;color:{T1};letter-spacing:.4px}}
.section-header p{{font-size:.75rem;color:{T2};margin-top:3px}}
.section{{margin-bottom:30px}}
.kpi-row{{display:grid;grid-template-columns:repeat(auto-fit,minmax(138px,1fr));gap:12px;margin-bottom:24px}}
.kpi-card{{background:{CARD};border:1px solid {BORDER};border-radius:10px;padding:16px 12px 12px;text-align:center;transition:border-color .2s,transform .15s}}
.kpi-card:hover{{border-color:{BLUE};transform:translateY(-2px)}}
.kpi-icon{{font-size:1.3rem;margin-bottom:5px}}
.kpi-value{{font-family:'Lora',serif;font-size:1.8rem;font-weight:700;line-height:1}}
.kpi-label{{font-size:.66rem;color:{T2};margin-top:5px;text-transform:uppercase;letter-spacing:.9px}}
.kpi-sub{{font-size:.70rem;color:{T3};margin-top:3px}}
.grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}}
.grid-3-1{{display:grid;grid-template-columns:2fr 1fr;gap:16px}}
.chart-card{{background:{CARD};border:1px solid {BORDER};border-radius:10px;padding:4px 6px 6px;overflow:hidden}}
.data-table{{width:100%;border-collapse:collapse;font-size:.80rem}}
.data-table th{{background:{PANEL};color:{T2};text-transform:uppercase;letter-spacing:.7px;font-size:.66rem;padding:9px 10px;text-align:left;border-bottom:1px solid {BORDER};position:sticky;top:0}}
.data-table td{{padding:8px 10px;border-bottom:1px solid {BORDER};color:{T1};vertical-align:middle}}
.data-table tr:hover td{{background:{PANEL}}}
.total-row td{{background:{PANEL}!important;border-top:2px solid {BORDER}}}
.table-wrap{{overflow-x:auto;border-radius:10px;border:1px solid {BORDER};margin-top:16px}}
.badge{{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.68rem;font-weight:700}}
.badge-green{{background:rgba(67,217,162,.15);color:{GREEN}}}
.badge-red{{background:rgba(244,91,105,.15);color:{RED}}}
.badge-gold{{background:rgba(245,200,66,.15);color:{GOLD}}}
.badge-gray{{background:rgba(90,100,130,.2);color:{T2}}}
.stats-strip{{display:grid;grid-template-columns:repeat(auto-fit,minmax(105px,1fr));gap:10px;margin:14px 0}}
.stat-chip{{background:{CARD};border:1px solid {BORDER};border-radius:8px;padding:10px 12px;text-align:center}}
.sc-val{{font-size:1.2rem;font-weight:700;font-family:'Lora',serif}}
.sc-lbl{{font-size:.66rem;color:{T2};margin-top:3px;text-transform:uppercase;letter-spacing:.6px}}
.prog-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:18px}}
.prog-card{{background:{CARD};border:1px solid {BORDER};border-radius:10px;padding:18px 16px;transition:transform .15s}}
.prog-card:hover{{transform:translateY(-2px)}}
.prog-header{{display:flex;align-items:flex-start;gap:12px;margin-bottom:14px}}
.prog-icon{{width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.2rem;flex-shrink:0}}
.prog-title{{font-family:'Lora',serif;font-size:.95rem;color:{T1};font-weight:700}}
.prog-desc{{font-size:.70rem;color:{T2};margin-top:2px;line-height:1.4}}
.prog-stats{{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;border-top:1px solid {BORDER};padding-top:12px}}
.prog-stat{{text-align:center}}
.ps-num{{font-size:1.45rem;font-weight:700;font-family:'Lora',serif;line-height:1}}
.ps-lbl{{font-size:.66rem;color:{T2};margin-top:3px;text-transform:uppercase;letter-spacing:.6px}}
.ps-pct{{font-size:.70rem;color:{T3};margin-top:2px}}
.side-cards{{display:flex;flex-direction:column;gap:12px}}
.side-card{{background:{CARD};border:1px solid {BORDER};border-radius:10px;padding:16px;text-align:center;flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center}}
.sc-big{{font-size:2rem;font-weight:700;font-family:'Lora',serif;line-height:1}}
.page-footer{{margin-top:32px;padding:14px 40px;border-top:1px solid {BORDER};display:flex;justify-content:space-between;font-size:.68rem;color:{T3};flex-wrap:wrap;gap:6px}}
</style>
</head>
<body>
<header class="page-header">
  <div class="header-left">
    <h1>🌱 City Garden Montessori School</h1>
    <div class="sub">Weekly Enrollment &amp; Data Dashboard &nbsp;·&nbsp; School Year {d['school_year']}</div>
  </div>
  <div class="header-right">
    <div class="week-badge">Week Ending {d['week_date']}</div>
    <div class="sys-badge">Infinite Campus &amp; Spreadsheet Calc</div>
  </div>
</header>

<div class="page-body">

<div class="section">
  <div class="section-header"><h2>📊 At a Glance</h2><p>Key metrics for this week</p></div>
  <div class="kpi-row">
    {kpi("Total Enrolled",  fmt_n(d['total_enrolled']),   f"Budget: {fmt_n(d['total_budget_enr'])} · Δ {int(d['total_var_enr']):+}", BLUE,   "🎒")}
    {kpi("Charter K–8",     fmt_n(d['charter_enrolled']), "Enrolled",                                                                  TEAL,   "🏫")}
    {kpi("This Week Att.",  fmt_pct(d['att_week_total']), "Whole School",  GREEN if d['att_week_total']>=ATT_TARGET else GOLD,           "📅" if d['att_week_total']>=ATT_TARGET else "⚠️")}
    {kpi("YTD Attendance",  fmt_pct(d['att_ytd_total']),  "Whole School",  GREEN if d['att_ytd_total']>=ATT_TARGET else GOLD,            "📈")}
    {kpi("Charter 90/90",   fmt_pct(d['charter_9090']),   "YTD Rate",      PURPLE,                                                      "🏆")}
    {kpi("Total ADA",       f"{d['ada_total']:,.1f}",     "Avg Daily Att.", ORANGE,                                                     "📊")}
    {kpi("ISS YTD",         fmt_n(d['iss_ytd']),          f"{int(d['iss_week'])} this week", GOLD if d['iss_ytd']>0 else GREEN,          "📋")}
    {kpi("OSS YTD",         fmt_n(d['oss_ytd']),          f"{int(d['oss_week'])} this week", RED  if d['oss_ytd']>20 else GOLD,          "⚠️")}
  </div>
</div>

<div class="section">
  <div class="section-header"><h2>🎒 Enrollment &amp; Attendance</h2><p>Actual enrollment vs. budgeted targets and weekly attendance by grade</p></div>
  <div class="grid-2" style="margin-bottom:16px">
    <div class="chart-card">{ch_enroll}</div>
    <div class="chart-card">{ch_att}</div>
  </div>
  <div class="table-wrap">
    <table class="data-table">
      <thead><tr>
        <th>Grade</th><th>Enrolled</th><th>Budget</th><th>Variance</th>
        <th>This Wk Att.</th><th>YTD Att.</th><th>ADA</th>
      </tr></thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
  <div class="chart-card" style="margin-top:16px">{ch_var}</div>
</div>

<div class="section">
  <div class="section-header"><h2>👥 Student Demographics</h2><p>Race/ethnicity and Free &amp; Reduced Lunch eligibility</p></div>
  <div class="grid-2" style="margin-bottom:12px">
    <div class="chart-card">{ch_race}</div>
    <div class="chart-card">{ch_frl}</div>
  </div>
  <div class="stats-strip">
    <div class="stat-chip"><div class="sc-val" style="color:{RED}">{fmt_pct(frl['Whole School'])}</div><div class="sc-lbl">Whole School FRL</div></div>
    <div class="stat-chip"><div class="sc-val" style="color:{ORANGE}">{fmt_pct(frl['Charter'])}</div><div class="sc-lbl">Charter FRL</div></div>
    <div class="stat-chip"><div class="sc-val" style="color:{GOLD}">{fmt_pct(frl['Preschool'])}</div><div class="sc-lbl">Preschool FRL</div></div>
    <div class="stat-chip"><div class="sc-val" style="color:{TEAL}">{fmt_pct(frl['EAEC'])}</div><div class="sc-lbl">EAEC FRL</div></div>
    <div class="stat-chip"><div class="sc-val" style="color:{BLUE}">{fmt_pct(frl['ECEC'])}</div><div class="sc-lbl">ECEC FRL</div></div>
    <div class="stat-chip"><div class="sc-val" style="color:{RED}">{fmt_n(d['frl_free']['Whole School'])}</div><div class="sc-lbl">Free Lunch</div></div>
    <div class="stat-chip"><div class="sc-val" style="color:{GOLD}">{fmt_n(d['frl_red']['Whole School'])}</div><div class="sc-lbl">Reduced Lunch</div></div>
  </div>
</div>

<div class="section">
  <div class="section-header"><h2>🎯 Special Programs</h2><p>IEP (SPED) · 504 Accommodation Plans · English Language Learners (ELL)</p></div>
  <div class="prog-grid">
    {prog_card("IEP — Individualized Education Plans","🎯",PURPLE,
      iep['Whole School']['n'], iep['Whole School']['pct'],
      iep['Charter']['n'],      iep['Charter']['pct'],
      iep['Preschool']['n'],    iep['Preschool']['pct'],
      "Students receiving special education services under IDEA")}
    {prog_card("504 — Accommodation Plans","📝",TEAL,
      s504['Whole School']['n'], s504['Whole School']['pct'],
      s504['Charter']['n'],      s504['Charter']['pct'],
      s504['Preschool']['n'],    s504['Preschool']['pct'],
      "Students with accommodations under Section 504")}
    {prog_card("ELL — English Language Learners","🌐",GOLD,
      ell['Whole School']['n'], ell['Whole School']['pct'],
      ell['Charter']['n'],      ell['Charter']['pct'],
      ell['Preschool']['n'],    ell['Preschool']['pct'],
      "Students currently receiving English language support")}
  </div>
  <div class="chart-card">{ch_prog}</div>
</div>

<div class="section">
  <div class="section-header"><h2>⚠️ Discipline</h2><p>In-school and out-of-school suspension — this week and year-to-date</p></div>
  <div class="grid-3-1">
    <div class="chart-card">{ch_disc}</div>
    <div class="side-cards">
      <div class="side-card">
        <div class="sc-big" style="color:{RED}">{int(d['iss_ytd']+d['oss_ytd'])}</div>
        <div class="sc-lbl">Total YTD Incidents</div>
      </div>
      <div class="side-card">
        <div class="sc-big" style="color:{GREEN}">{int(d['iss_week']+d['oss_week'])}</div>
        <div class="sc-lbl">This Week</div>
      </div>
      <div class="side-card" style="gap:6px">
        <div style="font-size:.68rem;color:{T2};text-transform:uppercase;letter-spacing:.6px;margin-bottom:4px">YTD Breakdown</div>
        <div style="font-size:.95rem;color:{RED}">ISS &nbsp;<b>{int(d['iss_ytd'])}</b></div>
        <div style="font-size:.95rem;color:{ORANGE}">OSS &nbsp;<b>{int(d['oss_ytd'])}</b></div>
      </div>
    </div>
  </div>
</div>

</div>
<footer class="page-footer">
  <span>🌱 City Garden Montessori School · {d['school_year']}</span>
  <span>Auto-generated from weekly summary Excel · Week ending {d['week_date']}</span>
  <span>Data: Infinite Campus + Spreadsheet Calc</span>
</footer>
</body></html>"""


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;margin-bottom:24px">
      <div style="font-size:2.4rem">🌱</div>
      <div style="font-family:'Lora',serif;font-size:1.1rem;color:#E8EEFF;font-weight:700;margin-top:4px">CGMS Dashboard</div>
      <div style="font-size:0.72rem;color:#3A4560;margin-top:3px;text-transform:uppercase;letter-spacing:1px">City Garden Montessori</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### Upload Weekly File")
    uploaded = st.file_uploader(
        "Drop your summary .xlsx here",
        type=["xlsx"]
    )

    st.markdown("<hr style='border-color:#1E2740;margin:20px 0'>", unsafe_allow_html=True)

    st.markdown("""
    <div style="font-size:0.72rem;color:#3A4560;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">Dashboard Sections</div>
    """, unsafe_allow_html=True)

    sections = [
        ("📊", "At a Glance", "8 key KPI cards"),
        ("🎒", "Enrollment", "Actual vs. budget by grade"),
        ("📅", "Attendance", "This week & YTD rates"),
        ("👥", "Demographics", "Race/ethnicity & FRL"),
        ("🎯", "Special Programs", "IEP · 504 · ELL breakdown"),
        ("⚠️", "Discipline", "ISS & OSS year-to-date"),
    ]
    for icon, title, desc in sections:
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;margin-bottom:4px;background:#111827">
          <span style="font-size:1rem">{icon}</span>
          <div>
            <div style="font-size:0.82rem;color:#C5CCDF;font-weight:500">{title}</div>
            <div style="font-size:0.70rem;color:#3A4560">{desc}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<hr style='border-color:#1E2740;margin:20px 0'>", unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:0.70rem;color:#2A3450;line-height:1.6">
      <b style="color:#3A4560">Format requirements</b><br>
      • Sheet name: <code style="color:#4F8EF7">summary</code><br>
      • Same row/column layout each week<br>
      • Source: Infinite Campus export
    </div>
    """, unsafe_allow_html=True)


# ── MAIN AREA ─────────────────────────────────────────────────────────────────
if uploaded is None:
    # Landing page
    st.markdown("""
    <div class="landing-wrap">
      <div class="landing-card">
        <div class="school-logo">🌱</div>
        <div class="school-name">City Garden Montessori School</div>
        <div class="school-tagline">Weekly Data Dashboard · Powered by Streamlit</div>

      </div>
    </div>
    """, unsafe_allow_html=True)

else:
    # Process and render dashboard
    try:
        with st.spinner("Parsing data and building dashboard…"):
            data = parse_excel(uploaded)
            html = build_dashboard_html(data)

        # Render inside scrollable iframe
        st.components.v1.html(html, height=7200, scrolling=True)

    except ValueError as e:
        if "summary" in str(e).lower() or "Worksheet" in str(e):
            st.error("❌ Sheet named **'summary'** not found in this file. Please upload the correct weekly summary export.", icon="🚫")
        else:
            st.error(f"❌ Could not parse the file: {e}", icon="🚫")
    except Exception as e:
        st.error(f"❌ Unexpected error: {e}", icon="🚫")
        with st.expander("Show error details"):
            import traceback
            st.code(traceback.format_exc())

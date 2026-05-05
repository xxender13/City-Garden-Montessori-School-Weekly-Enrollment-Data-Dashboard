
"""
City Garden Montessori School — Weekly Dashboard
Branded UI + Clean Sidebar + Fixed Landing Page
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
import warnings
warnings.filterwarnings("ignore")

# ── PAGE CONFIG ─────────────────────────────────────────
st.set_page_config(
    page_title="CGMS Dashboard",
    page_icon="🌱",
    layout="wide",
)

# ── CLEAN UI CSS ───────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');

#MainMenu, footer, header {visibility: hidden;}
.block-container {padding: 0 !important; max-width: 100%;}
.stApp {background: #0A0F1C;}

/* SIDEBAR */
[data-testid="stSidebar"] {
    background: #0F172A;
    border-right: 1px solid #1E293B;
}
[data-testid="stSidebarContent"] {
    padding: 24px 18px;
}
[data-testid="stSidebar"] * {
    font-family: 'DM Sans', sans-serif !important;
}

/* UPLOAD */
[data-testid="stFileUploader"] {
    background: #111827;
    border: 2px dashed #334155;
    border-radius: 10px;
    padding: 10px;
}
[data-testid="stFileUploader"]:hover {
    border-color: #4F8EF7;
}

/* LANDING */
.hero {
    min-height: 85vh;
    display:flex;
    align-items:center;
    justify-content:center;
}
.hero-card {
    background:#111827;
    border:1px solid #1E293B;
    border-radius:16px;
    padding:50px;
    max-width:820px;
}
.title {
    font-family:'Lora',serif;
    font-size:2.2rem;
    color:#E2E8F0;
    text-align:center;
}
.subtitle {
    text-align:center;
    color:#64748B;
    margin-bottom:30px;
}
.brand {
    text-align:center;
    font-size:0.8rem;
    color:#475569;
    margin-bottom:20px;
}

/* STEPS */
.steps {
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:16px;
}
.step {
    background:#1E293B;
    padding:18px;
    border-radius:10px;
    text-align:center;
}
.step b {color:white;}
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR ───────────────────────────────────────────
with st.sidebar:

    st.markdown("""
    <div style="text-align:center;margin-bottom:20px">
        <div style="font-size:2.5rem">🌱</div>
        <div style="font-size:1.1rem;color:white;font-weight:700">
            CGMS Dashboard
        </div>
        <div style="font-size:0.75rem;color:#64748B">
            City Garden Montessori School
        </div>
        <div style="font-size:0.7rem;color:#334155;margin-top:6px">
            PRiME Center · Saint Louis University
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Upload Weekly File")

    uploaded = st.file_uploader(
        "Upload summary.xlsx",
        type=["xlsx"]
    )

    st.markdown("---")

    st.markdown("### Sections")
    st.markdown("""
    • At a Glance  
    • Enrollment  
    • Attendance  
    • Demographics  
    • Programs  
    • Discipline  
    """)

    st.markdown("---")
    st.caption("Sheet name must be 'summary'")

# ── PARSER (same as yours) ────────────────────────────
def parse_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name="summary", header=None)

    def num(r,c):
        try: return float(df.iloc[r,c])
        except: return 0

    def pct(r,c):
        v = num(r,c)
        return v*100 if v<=1 else v

    data = {}
    data["week"] = str(df.iloc[1,1])[:10]
    data["enrolled"] = num(18,1)
    data["attendance"] = pct(18,4)
    data["ytd"] = pct(18,5)
    data["iss"] = num(54,2)
    data["oss"] = num(55,2)
    return data

# ── SIMPLE KPI VIEW (lightweight demo render) ─────────
def render_dashboard(d):

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Enrolled", int(d["enrolled"]))
    col2.metric("Attendance", f"{d['attendance']:.1f}%")
    col3.metric("YTD", f"{d['ytd']:.1f}%")
    col4.metric("Discipline", int(d["iss"] + d["oss"]))

    st.markdown("---")

    st.success("Dashboard loaded successfully")

# ── MAIN ─────────────────────────────────────────────
if uploaded is None:

    st.markdown("""
    <div class="hero">
        <div class="hero-card">

            <div class="title">City Garden Montessori School</div>
            <div class="subtitle">Weekly Data Dashboard</div>

            <div class="brand">
                Built by PRiME Center · Saint Louis University
            </div>

            <div class="steps">

                <div class="step">
                    <b>1. Upload File</b><br>
                    Add summary.xlsx
                </div>

                <div class="step">
                    <b>2. Auto Generate</b><br>
                    Dashboard builds instantly
                </div>

                <div class="step">
                    <b>3. Review</b><br>
                    Analyze insights
                </div>

            </div>

        </div>
    </div>
    """, unsafe_allow_html=True)

else:
    try:
        data = parse_excel(uploaded)
        render_dashboard(data)

    except Exception as e:
        st.error(f"Error: {e}")
```

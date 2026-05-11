"""
Exam Scheduler — Redesigned UI (Arabic-first, RTL, step-by-step UX)
Drop-in replacement for the original schedule_app.py
"""

from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from excel_io import read_input_workbook, write_schedule
from i18n import _
from logging_utils import get_log_buffer, reset_log_buffer, log
from scheduler import build_schedule

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="جدول المراقبة | Exam Scheduler",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# Inject custom CSS — Tajawal font, navy/gold palette, RTL
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap" rel="stylesheet">
<style>

/* ── Root tokens ─────────────────────────────────────────────────────── */
:root {
  --navy:   #0f1e3c;
  --navy2:  #162447;
  --gold:   #c9963a;
  --gold2:  #e8b84b;
  --cream:  #f5f0e8;
  --white:  #ffffff;
  --text:   #1a1a2e;
  --muted:  #6b7280;
  --success:#16a34a;
  --error:  #dc2626;
  --radius: 14px;
  --shadow: 0 4px 24px rgba(15,30,60,.10);
}

/* ── Global ───────────────────────────────────────────────────────────── */
html, body, [class*="css"] {
  font-family: 'Tajawal', sans-serif !important;
}

.stApp {
  background: var(--cream);
  direction: rtl;
}

/* ── Sidebar ──────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
  background: var(--navy) !important;
  border-left: 4px solid var(--gold);
  border-right: none;
}

[data-testid="stSidebar"] * {
  color: var(--white) !important;
  font-family: 'Tajawal', sans-serif !important;
  direction: rtl;
}

[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stFileUploader label,
[data-testid="stSidebar"] .stSlider label {
  font-size: 0.9rem;
  color: var(--gold2) !important;
  font-weight: 500;
  letter-spacing: .3px;
}

[data-testid="stSidebar"] .stButton > button {
  width: 100%;
  background: var(--gold) !important;
  color: var(--navy) !important;
  border: none !important;
  border-radius: var(--radius) !important;
  font-family: 'Tajawal', sans-serif !important;
  font-weight: 700 !important;
  font-size: 1.05rem !important;
  padding: 0.7rem 1rem !important;
  margin-top: 1rem;
  cursor: pointer;
  transition: background .2s, transform .15s;
  box-shadow: 0 2px 10px rgba(201,150,58,.3);
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: var(--gold2) !important;
  transform: translateY(-1px);
}

/* slider accent */
[data-testid="stSidebar"] .stSlider [data-baseweb="slider"] div[role="slider"] {
  background: var(--gold) !important;
}

/* ── Header banner ───────────────────────────────────────────────────── */
.app-header {
  background: linear-gradient(135deg, var(--navy) 0%, #1e3a8a 60%, var(--navy2) 100%);
  border-radius: var(--radius);
  padding: 2.2rem 2.5rem;
  margin-bottom: 1.8rem;
  box-shadow: var(--shadow);
  display: flex;
  align-items: center;
  gap: 1.2rem;
  position: relative;
  overflow: hidden;
}
.app-header::before {
  content: '';
  position: absolute;
  top: -40px; left: -40px;
  width: 180px; height: 180px;
  border-radius: 50%;
  background: rgba(201,150,58,.15);
  pointer-events: none;
}
.app-header::after {
  content: '';
  position: absolute;
  bottom: -30px; right: 60px;
  width: 120px; height: 120px;
  border-radius: 50%;
  background: rgba(255,255,255,.05);
  pointer-events: none;
}
.app-header .icon {
  font-size: 3.2rem;
  line-height: 1;
  z-index: 1;
}
.app-header .titles { z-index: 1; }
.app-header h1 {
  color: var(--white) !important;
  font-size: 1.9rem !important;
  font-weight: 900 !important;
  margin: 0 !important;
  padding: 0 !important;
  line-height: 1.2 !important;
}
.app-header .sub {
  color: var(--gold2);
  font-size: .95rem;
  font-weight: 400;
  margin-top: .3rem;
}

/* ── Step indicator ─────────────────────────────────────────────────── */
.steps-bar {
  display: flex;
  align-items: center;
  gap: 0;
  margin-bottom: 1.8rem;
  direction: rtl;
}
.step {
  display: flex;
  align-items: center;
  gap: .5rem;
  background: var(--white);
  border-radius: var(--radius);
  padding: .65rem 1.1rem;
  box-shadow: var(--shadow);
  font-weight: 600;
  font-size: .88rem;
  color: var(--muted);
  flex: 1;
  transition: all .25s;
}
.step.active {
  background: var(--navy);
  color: var(--white);
  box-shadow: 0 4px 20px rgba(15,30,60,.25);
}
.step.done {
  background: #dcfce7;
  color: var(--success);
}
.step-num {
  width: 26px; height: 26px;
  border-radius: 50%;
  background: currentColor;
  color: var(--white);
  display: flex; align-items: center; justify-content: center;
  font-size: .78rem; font-weight: 700;
  flex-shrink: 0;
}
.step.active .step-num { background: var(--gold); color: var(--navy); }
.step.done .step-num { background: var(--success); color: var(--white); }
.step-sep { flex: 0; padding: 0 .3rem; color: var(--muted); font-size: 1.2rem; }

/* ── Metric cards ────────────────────────────────────────────────────── */
.metrics-row {
  display: flex;
  gap: 1rem;
  margin-bottom: 1.8rem;
  direction: rtl;
}
.metric-card {
  flex: 1;
  background: var(--white);
  border-radius: var(--radius);
  padding: 1.2rem 1.5rem;
  box-shadow: var(--shadow);
  border-top: 4px solid var(--gold);
  text-align: right;
  transition: transform .2s;
}
.metric-card:hover { transform: translateY(-3px); }
.metric-card .mc-icon { font-size: 1.8rem; margin-bottom: .4rem; }
.metric-card .mc-val {
  font-size: 2.2rem;
  font-weight: 900;
  color: var(--navy);
  line-height: 1;
}
.metric-card .mc-label {
  font-size: .85rem;
  color: var(--muted);
  margin-top: .25rem;
  font-weight: 500;
}

/* ── Section card ────────────────────────────────────────────────────── */
.section-card {
  background: var(--white);
  border-radius: var(--radius);
  padding: 1.6rem 1.8rem;
  margin-bottom: 1.4rem;
  box-shadow: var(--shadow);
  direction: rtl;
}
.section-title {
  color: var(--navy);
  font-size: 1.1rem;
  font-weight: 700;
  margin-bottom: 1rem;
  padding-bottom: .6rem;
  border-bottom: 2px solid var(--gold);
  display: flex; align-items: center; gap: .5rem;
}

/* ── Upload zone styling ─────────────────────────────────────────────── */
[data-testid="stFileUploader"] {
  border: 2px dashed var(--gold) !important;
  border-radius: var(--radius) !important;
  background: #fffbf2 !important;
  padding: 1.2rem !important;
  direction: rtl;
}
[data-testid="stFileUploader"] label { display: none !important; }

/* ── Success / info / error banners ──────────────────────────────────── */
[data-testid="stAlert"] {
  border-radius: var(--radius) !important;
  font-family: 'Tajawal', sans-serif !important;
  direction: rtl;
}

/* ── DataFrame ───────────────────────────────────────────────────────── */
.stDataFrame {
  border-radius: var(--radius) !important;
  overflow: hidden !important;
  box-shadow: var(--shadow) !important;
  direction: rtl;
}
[data-testid="stDataFrame"] th {
  background: var(--navy) !important;
  color: var(--white) !important;
  font-family: 'Tajawal', sans-serif !important;
  font-weight: 600 !important;
}

/* ── Charts ──────────────────────────────────────────────────────────── */
[data-testid="stVegaLiteChart"],
[data-testid="stArrowVegaLiteChart"] {
  border-radius: var(--radius) !important;
  overflow: hidden;
}

/* ── Download button ─────────────────────────────────────────────────── */
.stDownloadButton > button {
  background: var(--gold) !important;
  color: var(--navy) !important;
  border: none !important;
  border-radius: var(--radius) !important;
  font-family: 'Tajawal', sans-serif !important;
  font-weight: 700 !important;
  font-size: 1rem !important;
  padding: .65rem 1.6rem !important;
  box-shadow: 0 2px 10px rgba(201,150,58,.35) !important;
  transition: background .2s, transform .15s !important;
}
.stDownloadButton > button:hover {
  background: var(--gold2) !important;
  transform: translateY(-2px) !important;
}

/* ── Expander ────────────────────────────────────────────────────────── */
[data-testid="stExpander"] {
  border-radius: var(--radius) !important;
  border: 1px solid #e5e7eb !important;
  direction: rtl;
}

/* ── Sidebar logo area ───────────────────────────────────────────────── */
.sidebar-logo {
  text-align: center;
  padding: 1.5rem 1rem 1rem;
  border-bottom: 1px solid rgba(201,150,58,.3);
  margin-bottom: 1.2rem;
}
.sidebar-logo .logo-icon { font-size: 2.8rem; }
.sidebar-logo h2 {
  color: var(--white) !important;
  font-size: 1.1rem !important;
  font-weight: 700 !important;
  margin: .4rem 0 0 !important;
}
.sidebar-logo .logo-sub {
  color: var(--gold2) !important;
  font-size: .8rem;
}

/* ── Info prompt ─────────────────────────────────────────────────────── */
.upload-prompt {
  background: linear-gradient(135deg, #fffbf2, #fef9ee);
  border: 1px solid #fde68a;
  border-right: 4px solid var(--gold);
  border-radius: var(--radius);
  padding: 2rem;
  text-align: center;
  color: var(--text);
  font-size: 1rem;
  margin-top: .5rem;
}
.upload-prompt .up-icon { font-size: 2.5rem; margin-bottom: .6rem; }

/* ── hide streamlit branding ─────────────────────────────────────────── */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
      <div class="logo-icon">🎓</div>
      <h2>جدول المراقبة</h2>
      <div class="logo-sub">Exam Scheduler</div>
    </div>
    """, unsafe_allow_html=True)

    lang = st.selectbox("🌐 اللغة / Language", ["AR", "FR"])

    st.markdown("---")

    uploaded = st.file_uploader(_("upload", lang), type=["xlsx"])

    st.markdown("---")

    time_limit = st.slider(_("time_limit", lang), 1, 30, 10)

    generate_btn = st.button(_("generate", lang))


# ─────────────────────────────────────────────────────────────────────────────
# Header banner
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="icon">📋</div>
  <div class="titles">
    <h1>جدول توزيع المراقبين</h1>
    <div class="sub">نظام تلقائي لإعداد جداول الامتحانات وتوزيع الأساتذة على القاعات</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Step indicator
# ─────────────────────────────────────────────────────────────────────────────
def step_bar(step: int):
    """step: 1 = upload, 2 = generate, 3 = done"""
    s1 = "active" if step == 1 else ("done" if step > 1 else "")
    s2 = "active" if step == 2 else ("done" if step > 2 else "")
    s3 = "active" if step == 3 else ("done" if step > 3 else "")

    labels = {
        "AR": ("رفع الملف", "توليد الجدول", "التحميل"),
        "FR": ("Importer", "Générer", "Télécharger"),
    }
    l1, l2, l3 = labels.get(lang, labels["AR"])

    st.markdown(f"""
    <div class="steps-bar">
      <div class="step {s1}">
        <span class="step-num">١</span>{l1}
      </div>
      <span class="step-sep">←</span>
      <div class="step {s2}">
        <span class="step-num">٢</span>{l2}
      </div>
      <span class="step-sep">←</span>
      <div class="step {s3}">
        <span class="step-num">٣</span>{l3}
      </div>
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Main logic
# ─────────────────────────────────────────────────────────────────────────────
if not uploaded:
    step_bar(1)
    st.markdown("""
    <div class="upload-prompt">
      <div class="up-icon">📂</div>
      <p><strong>ارفع ملف Excel الخاص بالقالب من الشريط الجانبي للبدء</strong></p>
      <p style="font-size:.88rem; color:#6b7280; margin-top:.5rem;">
        Téléversez votre fichier .xlsx depuis la barre latérale pour commencer
      </p>
    </div>
    """, unsafe_allow_html=True)

else:
    # ── Load file ──────────────────────────────────────────────────────────
    temp_dir = tempfile.gettempdir()
    source = Path(temp_dir) / uploaded.name
    with open(source, "wb") as f:
        f.write(uploaded.getbuffer())

    input_data = read_input_workbook(source)

    n_teachers = len(input_data.teachers)
    n_rooms    = len(input_data.rooms)
    n_sessions = len(input_data.sessions)

    step_bar(2 if not generate_btn else 3)

    # ── Metric cards ───────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="metrics-row">
      <div class="metric-card">
        <div class="mc-icon">👨‍🏫</div>
        <div class="mc-val">{n_teachers}</div>
        <div class="mc-label">{_("teachers", lang)}</div>
      </div>
      <div class="metric-card">
        <div class="mc-icon">🚪</div>
        <div class="mc-val">{n_rooms}</div>
        <div class="mc-label">{_("rooms", lang)}</div>
      </div>
      <div class="metric-card">
        <div class="mc-icon">📅</div>
        <div class="mc-val">{n_sessions}</div>
        <div class="mc-label">{_("sessions", lang)}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"✅  {_('file_loaded', lang)} — **{uploaded.name}**")

    # ── Generate ───────────────────────────────────────────────────────────
    if generate_btn:
        reset_log_buffer()
        with st.spinner("⏳  جارٍ إنشاء الجدول…" if lang == "AR" else "⏳  Génération en cours…"):
            try:
                schedule = build_schedule(
                    input_data.teachers,
                    input_data.sessions,
                    input_data.rooms,
                    input_data.active_pairs,
                    time_limit=time_limit,
                )
                ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                out_path = Path(temp_dir) / f"scheduled_{ts}.xlsx"
                write_schedule(source, input_data, schedule, out_path)

                st.success("🎉  " + ("تم إنشاء الجدول بنجاح!" if lang == "AR" else "Planning généré avec succès!"))

                # ── Distribution grid ──────────────────────────────────────
                st.markdown("""
                <div class="section-card">
                  <div class="section-title">🗺️ شبكة التوزيع</div>
                </div>
                """, unsafe_allow_html=True)

                grid = {}
                for room in input_data.rooms:
                    grid[room] = [
                        " / ".join(schedule.supervisors.get((room, s), []))
                        for s in input_data.sessions
                    ]
                df_grid = pd.DataFrame(grid, index=input_data.sessions).T

                st.dataframe(
                    df_grid.style.set_properties(**{
                        "text-align": "center",
                        "font-family": "Tajawal, sans-serif",
                    }).highlight_null(color="#fef9ee"),
                    use_container_width=True,
                )

                # ── Workload chart ─────────────────────────────────────────
                st.markdown("""
                <div class="section-card" style="margin-top:1.4rem">
                  <div class="section-title">📊 عبء العمل — الأساتذة</div>
                </div>
                """, unsafe_allow_html=True)

                df_load = (
                    pd.DataFrame(list(schedule.load.items()), columns=["الأستاذ", "الحصص"])
                    .sort_values("الحصص", ascending=False)
                )
                st.bar_chart(df_load, x="الأستاذ", y="الحصص", use_container_width=True)

                # ── Download ───────────────────────────────────────────────
                st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)
                with open(out_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️  {_('download', lang)}",
                        data=f,
                        file_name=out_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except ValueError as err:
                st.error(f"❌  {err}")
                log(str(err))

        with st.expander(f"📋  {_('log', lang)}"):
            st.code(get_log_buffer().getvalue(), language="")

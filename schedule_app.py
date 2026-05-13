"""
Exam Scheduler — v5
• session_subjects is now a list per session
• Subject input uses st.multiselect or free-text with comma separation
"""
from __future__ import annotations

import io
import math
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

from excel_io import InputData, ScheduleOutput, read_input_workbook
from scheduler import build_schedule

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="جدول المراقبة", page_icon="🎓",
                   layout="wide", initial_sidebar_state="collapsed")

_CSS = """
:root{--navy:#0f1e3c;--navy2:#162447;--gold:#c9963a;--gold2:#e8b84b;
--cream:#f2ede4;--white:#fff;--text:#1a1a2e;--muted:#6b7280;
--success:#16a34a;--error:#dc2626;--radius:14px;
--shadow:0 4px 24px rgba(15,30,60,.09);}
html,body,[class*="css"]{font-family:'Tajawal',sans-serif!important;}
.stApp{background:var(--cream);direction:rtl;}
.app-header{background:linear-gradient(135deg,var(--navy) 0%,#1e3a8a 60%,var(--navy2) 100%);
border-radius:var(--radius);padding:1.8rem 2.5rem;margin-bottom:1.8rem;
box-shadow:var(--shadow);position:relative;overflow:hidden;}
.app-header::before{content:'';position:absolute;top:-50px;left:-50px;width:200px;height:200px;
border-radius:50%;background:rgba(201,150,58,.12);pointer-events:none;}
.app-header h1{color:#fff!important;font-size:1.8rem!important;font-weight:900!important;margin:0!important;padding:0!important;}
.app-header .sub{color:var(--gold2);font-size:.9rem;margin-top:.3rem;}
.card{background:var(--white);border-radius:var(--radius);padding:1.6rem 1.8rem;margin-bottom:1.2rem;box-shadow:var(--shadow);direction:rtl;}
.card-title{font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:1rem;padding-bottom:.6rem;border-bottom:2px solid var(--gold);}
.opt-card{background:var(--white);border-radius:var(--radius);padding:2rem;box-shadow:var(--shadow);border-top:5px solid var(--gold);text-align:right;transition:transform .2s,box-shadow .2s;}
.opt-card:hover{transform:translateY(-4px);box-shadow:0 8px 32px rgba(15,30,60,.14);}
.opt-card .oc-icon{font-size:2.6rem;margin-bottom:.8rem;}
.opt-card h3{font-size:1.15rem;font-weight:800;color:var(--navy);margin:0 0 .5rem;}
.opt-card p{font-size:.88rem;color:var(--muted);line-height:1.6;}
.opt-card ul{font-size:.85rem;color:var(--text);padding-right:1.2rem;margin-top:.5rem;line-height:1.8;}
.blk{display:flex;align-items:center;gap:.6rem;font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:.9rem;padding-bottom:.6rem;border-bottom:1px solid #e5e7eb;}
.blk .bn{width:26px;height:26px;border-radius:50%;background:var(--navy);color:#fff;display:flex;align-items:center;justify-content:center;font-size:.78rem;font-weight:700;flex-shrink:0;}
.mrow{display:flex;gap:1rem;margin-bottom:1.6rem;direction:rtl;}
.mc{flex:1;background:var(--white);border-radius:var(--radius);padding:1rem 1.3rem;box-shadow:var(--shadow);border-top:4px solid var(--gold);text-align:right;}
.mc .mv{font-size:1.9rem;font-weight:900;color:var(--navy);line-height:1;}
.mc .ml{font-size:.8rem;color:var(--muted);margin-top:.2rem;}
.rtitle{font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:.8rem;margin-top:1.4rem;padding-bottom:.5rem;border-bottom:2px solid var(--gold);}
.stButton>button{border-radius:10px!important;font-family:'Tajawal',sans-serif!important;font-weight:600!important;transition:all .18s!important;}
.stButton>button:hover{transform:translateY(-1px)!important;}
.stDownloadButton>button{background:var(--navy)!important;color:#fff!important;border:none!important;border-radius:10px!important;font-family:'Tajawal',sans-serif!important;font-weight:600!important;}
.stTextInput input,.stNumberInput input{border-radius:10px!important;direction:rtl!important;font-family:'Tajawal',sans-serif!important;}
.stTextInput input:focus,.stNumberInput input:focus{border-color:var(--gold)!important;box-shadow:0 0 0 3px rgba(201,150,58,.15)!important;}
.stCheckbox label{direction:rtl!important;font-family:'Tajawal',sans-serif!important;}
[data-testid="stAlert"]{border-radius:var(--radius)!important;font-family:'Tajawal',sans-serif!important;direction:rtl;}
.stDataFrame{border-radius:var(--radius)!important;overflow:hidden!important;}
#MainMenu,footer,header{visibility:hidden;}
"""
st.markdown('<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap" rel="stylesheet">',
            unsafe_allow_html=True)
st.markdown(f"<style>{_CSS}</style>", unsafe_allow_html=True)

# ─── Session state ────────────────────────────────────────────────────────────
_DEFAULTS = dict(
    page="welcome", input_mode=None,
    teachers=[], teacher_subjects={},
    rooms=[],
    sessions=[], session_dates={},
    session_subjects={},   # Dict[str, List[str]]
    matrix={},

    result=None,
)
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


def go(page):
    st.session_state.page = page
    st.rerun()


def sync_matrix():
    valid = {(r, s) for r in st.session_state.rooms for s in st.session_state.sessions}
    for k in list(st.session_state.matrix):
        if k not in valid:
            del st.session_state.matrix[k]
    for k in valid:
        if k not in st.session_state.matrix:
            st.session_state.matrix[k] = False





# ─── Excel style helpers ──────────────────────────────────────────────────────
_CA  = Alignment(horizontal="center", vertical="center", wrap_text=True)
_RA  = Alignment(horizontal="right",  vertical="center")
_TH  = Border(left=Side(style="thin"), right=Side(style="thin"),
              top=Side(style="thin"),  bottom=Side(style="thin"))
_HF  = PatternFill("solid", fgColor="0F1E3C")
_HFT = Font(color="FFFFFF", bold=True, name="Arial")
_GF  = PatternFill("solid", fgColor="C9963A")
_GFT = Font(color="FFFFFF", bold=True, name="Arial")
_BF  = PatternFill("solid", fgColor="FEF9EE")
_RF  = PatternFill("solid", fgColor="F8FAFC")


def _h(ws, row, col, val, w=None):
    c = ws.cell(row, col, val)
    c.fill = _HF; c.font = _HFT; c.alignment = _CA; c.border = _TH
    if w:
        ws.column_dimensions[c.column_letter].width = w
    return c


def _c(ws, row, col, val="", bold=False, fill=None, align=_CA):
    c = ws.cell(row, col, val)
    c.font = Font(bold=bold, name="Arial")
    if fill:
        c.fill = fill
    c.alignment = align; c.border = _TH
    return c


def _subjects_label(data: InputData, session: str) -> str:
    subjs = data.session_subjects.get(session, [])
    return "، ".join(subjs) if subjs else ""


# ─── Excel builders ───────────────────────────────────────────────────────────

def build_distribution_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "التوزيع"
    _h(ws, 1, 1, "القاعة / الحصة", w=22)
    for j, s in enumerate(data.sessions, 2):
        parts = [s]
        if data.session_dates.get(s):   parts.append(data.session_dates[s])
        subjs = data.session_subjects.get(s, [])
        if subjs: parts.append("(" + "، ".join(subjs) + ")")
        _h(ws, 1, j, "\n".join(parts))
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 26
    _c(ws, 2, 1, "الاحتياطيون", fill=_BF)
    ws.cell(2, 1).font = Font(italic=True, name="Arial")
    for j, s in enumerate(data.sessions, 2):
        c = _c(ws, 2, j, " / ".join(out.backups.get(s, [])), fill=_BF)
        c.font = Font(italic=True, name="Arial")
    # مداوم row
    _MF = PatternFill("solid", fgColor="E0F2FE")  # light blue
    _c(ws, 3, 1, "المداومون", fill=_MF)
    ws.cell(3, 1).font = Font(bold=True, name="Arial")
    for j, s in enumerate(data.sessions, 2):
        mou = out.moudawim.get(s, [])
        c = _c(ws, 3, j, " / ".join(mou), fill=_MF)
        c.font = Font(bold=True, name="Arial")
    for i, room in enumerate(data.rooms, 4):
        _c(ws, i, 1, room, bold=True, fill=_RF)
        for j, session in enumerate(data.sessions, 2):
            v = " / ".join(out.supervisors.get((room, session), [])) if (room, session) in data.active_pairs else "—"
            _c(ws, i, j, v)
    ws.freeze_panes = "B3"
    ws2 = wb.create_sheet("الحمل الوظيفي")
    _h(ws2, 1, 1, "الأستاذ", w=26); _h(ws2, 1, 2, "المادة", w=20); _h(ws2, 1, 3, "عدد الحصص", w=16)
    for i, (t, load) in enumerate(sorted(out.load.items(), key=lambda x: -x[1]), 2):
        _c(ws2, i, 1, t); _c(ws2, i, 2, data.teacher_subjects.get(t, "")); _c(ws2, i, 3, load)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def build_mapping_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "الخريطة التفصيلية"
    _h(ws, 1, 1, "الأستاذ", w=26); _h(ws, 1, 2, "مادة الأستاذ", w=20)
    _h(ws, 1, 3, "القاعة", w=22);  _h(ws, 1, 4, "الحصة", w=22)
    _h(ws, 1, 5, "التاريخ", w=18); _h(ws, 1, 6, "مواد الحصة", w=30); _h(ws, 1, 7, "الدور", w=14)
    row = 2
    for t in sorted(data.teachers):
        subj = data.teacher_subjects.get(t, "")
        for (room, session), sups in sorted(out.supervisors.items(), key=lambda x: x[0][1]):
            if t in sups:
                _c(ws, row, 1, t); _c(ws, row, 2, subj)
                _c(ws, row, 3, room); _c(ws, row, 4, session)
                _c(ws, row, 5, data.session_dates.get(session, ""))
                _c(ws, row, 6, "، ".join(data.session_subjects.get(session, [])))
                _c(ws, row, 7, "مراقب"); row += 1
        for session, bups in sorted(out.backups.items()):
            if t in bups:
                _c(ws, row, 1, t); _c(ws, row, 2, subj)
                _c(ws, row, 3, "—"); _c(ws, row, 4, session)
                _c(ws, row, 5, data.session_dates.get(session, ""))
                _c(ws, row, 6, "، ".join(data.session_subjects.get(session, [])))
                _c(ws, row, 7, "احتياطي"); row += 1
        for session, mous in sorted(out.moudawim.items()):
            if t in mous:
                _c(ws, row, 1, t); _c(ws, row, 2, subj)
                _c(ws, row, 3, "—"); _c(ws, row, 4, session)
                _c(ws, row, 5, data.session_dates.get(session, ""))
                _c(ws, row, 6, "، ".join(data.session_subjects.get(session, [])))
                _c(ws, row, 7, "مداوم"); row += 1
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def build_convocations_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    date_str = datetime.now().strftime("%d/%m/%Y")

    for t in sorted(data.teachers):
        duties = []
        for (room, session), sups in out.supervisors.items():
            if t in sups:
                duties.append({"session": session,
                               "date":    data.session_dates.get(session, ""),
                               "subjects": "، ".join(data.session_subjects.get(session, [])),
                               "role":    "مراقب"})
        for session, bups in out.backups.items():
            if t in bups:
                duties.append({"session": session,
                               "date":    data.session_dates.get(session, ""),
                               "subjects": "، ".join(data.session_subjects.get(session, [])),
                               "role":    "احتياطي"})
        for session, mous in out.moudawim.items():
            if t in mous:
                duties.append({"session": session,
                               "date":    data.session_dates.get(session, ""),
                               "subjects": "، ".join(data.session_subjects.get(session, [])),
                               "role":    "مداوم"})
        duties.sort(key=lambda d: (d["date"], d["session"]))

        ws = wb.create_sheet(t[:28])
        ws.column_dimensions["A"].width = 8;  ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 30; ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 16

        ws.merge_cells("A1:E1")
        c = ws["A1"]; c.value = "إشعار بمهمة المراقبة"
        c.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
        c.fill = _HF; c.alignment = _CA; ws.row_dimensions[1].height = 34

        ws.merge_cells("A2:E2")
        c = ws["A2"]
        c.value = f"الأستاذ(ة): {t}    |    المادة: {data.teacher_subjects.get(t, '—')}"
        c.font = Font(bold=True, size=12, color="FFFFFF", name="Arial")
        c.fill = _GF; c.alignment = _CA; c.border = _TH; ws.row_dimensions[2].height = 28

        ws.merge_cells("A3:E3")
        c = ws["A3"]; c.value = f"تاريخ الطباعة: {date_str}"
        c.font = Font(size=10, italic=True, name="Arial"); c.alignment = _CA; c.border = _TH

        for col, label in enumerate(["الرقم", "الحصة", "المواد الممتحَنة", "التاريخ", "الدور"], 1):
            _h(ws, 4, col, label)
        ws.row_dimensions[4].height = 22

        if duties:
            for i, d in enumerate(duties, 1):
                fill = _BF if d["role"] == "احتياطي" else (PatternFill("solid", fgColor="E0F2FE") if d["role"] == "مداوم" else None)
                _c(ws, 4+i, 1, i,             fill=fill)
                _c(ws, 4+i, 2, d["session"],  fill=fill)
                _c(ws, 4+i, 3, d["subjects"], fill=fill)
                _c(ws, 4+i, 4, d["date"],     fill=fill)
                _c(ws, 4+i, 5, d["role"],     fill=fill)
        else:
            ws.merge_cells("A5:E5")
            c = ws["A5"]; c.value = "لا توجد مهام مُسنَدة"
            c.font = Font(italic=True, color="6B7280", name="Arial"); c.alignment = _CA

        last = 4 + max(len(duties), 1) + 2
        ws.merge_cells(f"A{last}:E{last}")
        c = ws.cell(last, 1, "توقيع الأستاذ(ة): ___________________________")
        c.alignment = _RA; c.font = Font(size=10, name="Arial")
        ws.row_dimensions[last].height = 30

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ─── Validation & scheduler runner ───────────────────────────────────────────
def _run_scheduler():
    teachers         = st.session_state.teachers
    rooms            = st.session_state.rooms
    sessions_        = st.session_state.sessions
    teacher_subjects = st.session_state.teacher_subjects
    session_subjects = st.session_state.session_subjects  # Dict[str, List[str]]
    session_dates    = st.session_state.session_dates
    active_pairs     = {k for k, v in st.session_state.matrix.items() if v}

    errors = []
    if len(teachers) < 2:  errors.append("يجب إدخال أستاذين على الأقل")
    if not rooms:           errors.append("يجب إدخال قاعة واحدة على الأقل")
    if not sessions_:       errors.append("يجب إدخال حصة واحدة على الأقل")
    if not active_pairs:    errors.append("يجب تفعيل زوج واحد على الأقل في المصفوفة")

    if not errors:
        from collections import Counter
        rooms_per_sess = Counter(s for (r, s) in active_pairs)
        for s, cnt in rooms_per_sess.items():
            bup_s  = math.ceil(0.22 * 2 * cnt)
            # teachers locked as moudawim-only (teach an examined subject)
            locked_s = {t for t in teachers
                        for subj in session_subjects.get(s, [])
                        if teacher_subjects.get(t, "") == subj}
            # free pool = all teachers minus locked ones
            free_s = len(teachers) - len(locked_s)
            needed = 2 * cnt + bup_s
            if free_s < needed:
                shortage = needed - free_s
                errors.append(
                    f"عدد الأساتذة غير كافٍ للحصة «{s}» — "
                    f"يلزم {needed} أستاذاً حراً (مراقبون + احتياطيون) "
                    f"لكن المتاح {free_s} (الكلي {len(teachers)} - {len(locked_s)} مداوم مُقيَّد). "
                    f"نقص: {shortage} أستاذ"
                )
                break
        max_usable = 2 * len(active_pairs)
        if len(teachers) > max_usable:
            errors.append(f"عدد الأساتذة كبير جداً — الحد الأقصى {max_usable} لـ {len(active_pairs)} زوج نشط")
        # each teacher is backup exactly once →
        # N teachers must be >= N active sessions (each needs at least 1 backup)
        n_active_sessions = sum(1 for s in rooms_per_sess if rooms_per_sess[s] > 0)
        if len(teachers) < n_active_sessions:
            errors.append(
                f"عدد الأساتذة ({len(teachers)}) أقل من عدد الحصص النشطة "
                f"({n_active_sessions}) — كل حصة تحتاج على الأقل احتياطياً واحداً."
            )

    if errors:
        for e in errors:
            st.error(f"❌  {e}")
        return

    data = InputData(
        teachers=teachers, rooms=rooms, sessions=sessions_,
        active_pairs=active_pairs,
        teacher_subjects=teacher_subjects,
        session_dates=session_dates,
        session_subjects=session_subjects,
    )
    with st.spinner("⏳  جارٍ حساب التوزيع الأمثل…"):
        try:
            schedule = build_schedule(
                data.teachers, data.sessions, data.rooms, data.active_pairs,
                teacher_subjects=data.teacher_subjects,
                session_subjects=data.session_subjects,
                time_limit=20,
            )
            st.session_state.result = (data, schedule)
            go("results")
        except Exception as err:
            st.error(f"❌  {err}")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: WELCOME
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "welcome":
    go("input_excel")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: INPUT — EXCEL
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "input_excel":
    st.markdown("""
    <div class="app-header">
      <h1>🎓 جدول توزيع المراقبين</h1>
      <div class="sub">ارفع ملف القالب لتوليد الجدول تلقائياً</div>
    </div>""", unsafe_allow_html=True)

    # Template download
    col_info, col_dl = st.columns([3, 1], gap="large")
    with col_info:
        st.markdown("""<div class="card">
        <div class="card-title">📋 طريقة الاستخدام</div>
        <ol style="direction:rtl;line-height:2.2;font-size:.95rem">
          <li>حمّل القالب الرسمي بالضغط على الزر المجاور</li>
          <li>افتح الملف وتوجّه إلى ورقة <b>الإعدادات</b> — أدخل عدد القاعات والحصص</li>
          <li>انتقل إلى ورقة <b>الأساتذة</b> — أدخل الأسماء والمواد وأرقام التأجير</li>
          <li>انتقل إلى ورقة <b>الحصص</b> — أدخل التاريخ والتوقيت والمواد وعدد القاعات العاملة</li>
          <li>ورقة <b>المصفوفة</b> تُحسب تلقائياً — لا تعدّلها</li>
          <li>ارفع الملف أدناه وولّد الجدول</li>
        </ol>
        </div>""", unsafe_allow_html=True)
    with col_dl:
        st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)
        try:
            import pathlib
            tpl = pathlib.Path(__file__).parent / "قالب_جدول_المراقبة.xlsx"
            if tpl.exists():
                st.download_button(
                    "⬇️  تحميل القالب",
                    data=open(tpl,"rb").read(),
                    file_name="قالب_جدول_المراقبة.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        except Exception:
            pass

    st.markdown("<div style='margin-top:.5rem'></div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("📂  ارفع ملف Excel هنا", type=["xlsx"])
    if uploaded:
        import tempfile, pathlib
        tmp = pathlib.Path(tempfile.gettempdir()) / uploaded.name
        tmp.write_bytes(uploaded.getbuffer())
        try:
            data = read_input_workbook(tmp)
            st.session_state.teachers         = list(data.teachers)
            st.session_state.teacher_subjects = dict(data.teacher_subjects)
            st.session_state.rooms            = list(data.rooms)
            st.session_state.sessions         = list(data.sessions)
            st.session_state.session_dates    = dict(data.session_dates)
            st.session_state.session_subjects = dict(data.session_subjects)
            st.session_state.matrix           = {k: True for k in data.active_pairs}
            st.session_state.input_mode       = "excel"

            c1,c2,c3,c4 = st.columns(4)
            c1.metric("👨‍🏫 أستاذ",     len(data.teachers))
            c2.metric("🚪 قاعة",         len(data.rooms))
            c3.metric("📅 حصة",           len(data.sessions))
            c4.metric("✅ زوج نشط",       len(data.active_pairs))

            _, bc, _ = st.columns([1,2,1])
            with bc:
                if st.button("⚡  توليد الجدول", type="primary", use_container_width=True):
                    _run_scheduler()
        except Exception as e:
            st.error(f"❌  {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: INPUT — QUICK
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: RESULTS
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "results":
    st.markdown("""
    <div class="app-header">
      <h1>🎉 الجدول جاهز</h1>
      <div class="sub">تحميل الجدول، الخريطة التفصيلية، والإشعارات</div>
    </div>""", unsafe_allow_html=True)

    if st.button("← عودة"):
        go("input_excel")

    data, schedule = st.session_state.result
    # total backups = 1 per teacher (exactly-once constraint)
    total_bup = len(data.teachers)
    bup_per_sess = {}
    for s in data.sessions:
        bup_per_sess[s] = len(schedule.backups.get(s, []))

    st.success(f"✅  تم توليد الجدول بنجاح!")

    avg_load = round(sum(schedule.load.values()) / max(len(schedule.load), 1), 1)
    st.markdown(f"""
    <div class="mrow">
      <div class="mc"><div class="mv">{len(data.teachers)}</div><div class="ml">👨‍🏫 أستاذ</div></div>
      <div class="mc"><div class="mv">{total_bup}</div><div class="ml">🔁 إجمالي الاحتياطيين</div></div>
      <div class="mc"><div class="mv">{len(data.active_pairs)}</div><div class="ml">✅ زوج نشط</div></div>
      <div class="mc"><div class="mv">{avg_load}</div><div class="ml">📊 متوسط الحصص</div></div>
    </div>""", unsafe_allow_html=True)

    # Show moudawim per session
    if any(schedule.moudawim.values()):
        st.markdown('<div class="rtitle">🎓 المداومون لكل حصة</div>', unsafe_allow_html=True)
        mou_rows = []
        for s in data.sessions:
            mou_list = schedule.moudawim.get(s, [])
            if mou_list:
                for t in mou_list:
                    mou_rows.append({"الحصة": s, "التاريخ": data.session_dates.get(s,""),
                                     "المواد": "، ".join(data.session_subjects.get(s,[])),
                                     "المداوم": t, "مادته": data.teacher_subjects.get(t,"")})
        if mou_rows:
            st.dataframe(pd.DataFrame(mou_rows), use_container_width=True)

    st.markdown('<div class="rtitle">🗺️ شبكة التوزيع</div>', unsafe_allow_html=True)
    grid = {}
    for room in data.rooms:
        grid[room] = [
            " / ".join(schedule.supervisors.get((room, s), ["—"])) if (room, s) in data.active_pairs else "—"
            for s in data.sessions
        ]
    df_grid = pd.DataFrame(grid, index=data.sessions).T
    df_grid.index.name = "القاعة"
    st.dataframe(df_grid, use_container_width=True)

    st.markdown('<div class="rtitle">📊 عبء العمل</div>', unsafe_allow_html=True)
    df_load = pd.DataFrame(list(schedule.load.items()), columns=["الأستاذ", "الحصص"]).sort_values("الحصص", ascending=False)
    st.bar_chart(df_load.set_index("الأستاذ"), use_container_width=True)

    st.markdown('<div class="rtitle">⬇️ التحميلات</div>', unsafe_allow_html=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.markdown("**📋 جدول التوزيع**")
        st.caption("شبكة القاعات × الحصص مع المواد والتواريخ")
        st.download_button("⬇️  تحميل جدول التوزيع",
                           data=build_distribution_excel(data, schedule),
                           file_name=f"jadwal_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with dl2:
        st.markdown("**🗂️ الخريطة التفصيلية**")
        st.caption("صف لكل مهمة: أستاذ · مادة · قاعة · حصة · مواد الحصة · دور")
        st.download_button("⬇️  تحميل الخريطة",
                           data=build_mapping_excel(data, schedule),
                           file_name=f"kharita_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with dl3:
        st.markdown("**📨 إشعارات المراقبة**")
        st.caption("مراقبة + احتياط — المواد الممتحَنة — بدون ذكر القاعة")
        st.download_button("⬇️  تحميل الإشعارات",
                           data=build_convocations_excel(data, schedule),
                           file_name=f"ish3arat_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
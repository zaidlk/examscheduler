"""
Exam Scheduler — v7
Excel-only input. No manual mode.
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

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="جدول المراقبة",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
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
.app-header h1{color:#fff!important;font-size:1.8rem!important;font-weight:900!important;
margin:0!important;padding:0!important;}
.app-header .sub{color:var(--gold2);font-size:.9rem;margin-top:.3rem;}
.card{background:var(--white);border-radius:var(--radius);padding:1.6rem 1.8rem;
margin-bottom:1.2rem;box-shadow:var(--shadow);direction:rtl;}
.card-title{font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:1rem;
padding-bottom:.6rem;border-bottom:2px solid var(--gold);}
.mrow{display:flex;gap:1rem;margin-bottom:1.6rem;direction:rtl;}
.mc{flex:1;background:var(--white);border-radius:var(--radius);padding:1rem 1.3rem;
box-shadow:var(--shadow);border-top:4px solid var(--gold);text-align:right;}
.mc .mv{font-size:1.9rem;font-weight:900;color:var(--navy);line-height:1;}
.mc .ml{font-size:.8rem;color:var(--muted);margin-top:.2rem;}
.rtitle{font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:.8rem;
margin-top:1.4rem;padding-bottom:.5rem;border-bottom:2px solid var(--gold);}
.stButton>button{border-radius:10px!important;font-family:'Tajawal',sans-serif!important;
font-weight:600!important;transition:all .18s!important;}
.stButton>button:hover{transform:translateY(-1px)!important;}
.stDownloadButton>button{background:var(--navy)!important;color:#fff!important;
border:none!important;border-radius:10px!important;
font-family:'Tajawal',sans-serif!important;font-weight:600!important;}
[data-testid="stAlert"]{border-radius:var(--radius)!important;
font-family:'Tajawal',sans-serif!important;direction:rtl;}
.stDataFrame{border-radius:var(--radius)!important;overflow:hidden!important;}
#MainMenu,footer,header{visibility:hidden;}
"""
st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap" rel="stylesheet">',
    unsafe_allow_html=True,
)
st.markdown(f"<style>{_CSS}</style>", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for k, v in [("page", "upload"), ("result", None)]:
    if k not in st.session_state:
        st.session_state[k] = v


def go(page: str):
    st.session_state.page = page
    st.rerun()


# ── Excel style helpers ───────────────────────────────────────────────────────
_CA  = Alignment(horizontal="center", vertical="center", wrap_text=True)
_RA  = Alignment(horizontal="right",  vertical="center")
_TH  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
_HF  = PatternFill("solid", fgColor="0F1E3C")
_HFT = Font(color="FFFFFF", bold=True, name="Arial")
_GF  = PatternFill("solid", fgColor="C9963A")
_GFT = Font(color="FFFFFF", bold=True, name="Arial")
_BF  = PatternFill("solid", fgColor="FEF9EE")
_BLF = PatternFill("solid", fgColor="E0F2FE")
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
    c.alignment = align
    c.border = _TH
    return c


# ── Excel output builders ─────────────────────────────────────────────────────

def _subjects_str(data: InputData, session: str) -> str:
    return "، ".join(data.session_subjects.get(session, []))


def build_distribution_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "التوزيع"

    _h(ws, 1, 1, "القاعة / الحصة", w=22)
    for j, s in enumerate(data.sessions, 2):
        parts = [s]
        if data.session_dates.get(s):
            parts.append(data.session_dates[s])
        subjs = data.session_subjects.get(s, [])
        if subjs:
            parts.append("(" + "، ".join(subjs) + ")")
        _h(ws, 1, j, "\n".join(parts))
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 26

    # Moudawim row
    _c(ws, 2, 1, "المداومون", fill=_BLF)
    ws.cell(2, 1).font = Font(bold=True, name="Arial")
    for j, s in enumerate(data.sessions, 2):
        mou = out.moudawim.get(s, [])
        c = _c(ws, 2, j, " / ".join(mou), fill=_BLF)
        c.font = Font(bold=True, name="Arial")

    # Backup row
    _c(ws, 3, 1, "الاحتياطيون", fill=_BF)
    ws.cell(3, 1).font = Font(italic=True, name="Arial")
    for j, s in enumerate(data.sessions, 2):
        c = _c(ws, 3, j, " / ".join(out.backups.get(s, [])), fill=_BF)
        c.font = Font(italic=True, name="Arial")

    # Room rows
    for i, room in enumerate(data.rooms, 4):
        _c(ws, i, 1, room, bold=True, fill=_RF)
        for j, session in enumerate(data.sessions, 2):
            v = (
                " / ".join(out.supervisors.get((room, session), []))
                if (room, session) in data.active_pairs
                else "—"
            )
            _c(ws, i, j, v)

    ws.freeze_panes = "B4"

    # Workload sheet
    ws2 = wb.create_sheet("الحمل الوظيفي")
    _h(ws2, 1, 1, "الأستاذ", w=26)
    _h(ws2, 1, 2, "المادة", w=20)
    _h(ws2, 1, 3, "عدد الحصص", w=16)
    for i, (t, load) in enumerate(sorted(out.load.items(), key=lambda x: -x[1]), 2):
        _c(ws2, i, 1, t)
        _c(ws2, i, 2, data.teacher_subjects.get(t, ""))
        _c(ws2, i, 3, load)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_mapping_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "الخريطة التفصيلية"
    for col, (label, w) in enumerate(
        [("الأستاذ", 26), ("مادة الأستاذ", 20), ("القاعة", 22),
         ("الحصة", 22), ("التاريخ", 18), ("مواد الحصة", 30), ("الدور", 14)], 1
    ):
        _h(ws, 1, col, label, w=w)

    row = 2
    for t in sorted(data.teachers):
        subj = data.teacher_subjects.get(t, "")
        for (room, session), sups in sorted(out.supervisors.items(), key=lambda x: x[0][1]):
            if t in sups:
                _c(ws, row, 1, t); _c(ws, row, 2, subj); _c(ws, row, 3, room)
                _c(ws, row, 4, session); _c(ws, row, 5, data.session_dates.get(session, ""))
                _c(ws, row, 6, _subjects_str(data, session)); _c(ws, row, 7, "مراقب")
                row += 1
        for session, bups in sorted(out.backups.items()):
            if t in bups:
                _c(ws, row, 1, t); _c(ws, row, 2, subj); _c(ws, row, 3, "—")
                _c(ws, row, 4, session); _c(ws, row, 5, data.session_dates.get(session, ""))
                _c(ws, row, 6, _subjects_str(data, session)); _c(ws, row, 7, "احتياطي")
                row += 1
        for session, mous in sorted(out.moudawim.items()):
            if t in mous:
                _c(ws, row, 1, t); _c(ws, row, 2, subj); _c(ws, row, 3, "—")
                _c(ws, row, 4, session); _c(ws, row, 5, data.session_dates.get(session, ""))
                _c(ws, row, 6, _subjects_str(data, session)); _c(ws, row, 7, "مداوم")
                row += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_convocations_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    date_str = datetime.now().strftime("%d/%m/%Y")

    for t in sorted(data.teachers):
        duties = []
        for (room, session), sups in out.supervisors.items():
            if t in sups:
                duties.append({
                    "session": session, "date": data.session_dates.get(session, ""),
                    "subjects": _subjects_str(data, session), "role": "مراقب",
                })
        for session, bups in out.backups.items():
            if t in bups:
                duties.append({
                    "session": session, "date": data.session_dates.get(session, ""),
                    "subjects": _subjects_str(data, session), "role": "احتياطي",
                })
        for session, mous in out.moudawim.items():
            if t in mous:
                duties.append({
                    "session": session, "date": data.session_dates.get(session, ""),
                    "subjects": _subjects_str(data, session), "role": "مداوم",
                })
        duties.sort(key=lambda d: (d["date"], d["session"]))

        ws = wb.create_sheet(t[:28])
        for col, w in zip("ABCDE", [8, 26, 30, 18, 16]):
            ws.column_dimensions[col].width = w

        ws.merge_cells("A1:E1")
        c = ws["A1"]
        c.value = "إشعار بمهمة المراقبة"
        c.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
        c.fill = _HF; c.alignment = _CA; ws.row_dimensions[1].height = 34

        ws.merge_cells("A2:E2")
        c = ws["A2"]
        c.value = f"الأستاذ(ة): {t}    |    المادة: {data.teacher_subjects.get(t, '—')}"
        c.font = Font(bold=True, size=12, color="FFFFFF", name="Arial")
        c.fill = _GF; c.alignment = _CA; c.border = _TH; ws.row_dimensions[2].height = 28

        ws.merge_cells("A3:E3")
        c = ws["A3"]
        c.value = f"تاريخ الطباعة: {date_str}"
        c.font = Font(size=10, italic=True, name="Arial"); c.alignment = _CA; c.border = _TH

        for col, label in enumerate(["الرقم", "الحصة", "المواد الممتحَنة", "التاريخ", "الدور"], 1):
            _h(ws, 4, col, label)
        ws.row_dimensions[4].height = 22

        if duties:
            for i, d in enumerate(duties, 1):
                fill = (
                    _BF  if d["role"] == "احتياطي" else
                    _BLF if d["role"] == "مداوم"   else
                    None
                )
                _c(ws, 4+i, 1, i,            fill=fill)
                _c(ws, 4+i, 2, d["session"],  fill=fill)
                _c(ws, 4+i, 3, d["subjects"], fill=fill)
                _c(ws, 4+i, 4, d["date"],     fill=fill)
                _c(ws, 4+i, 5, d["role"],     fill=fill)
        else:
            ws.merge_cells("A5:E5")
            c = ws["A5"]
            c.value = "لا توجد مهام مُسنَدة"
            c.font = Font(italic=True, color="6B7280", name="Arial"); c.alignment = _CA

        last = 4 + max(len(duties), 1) + 2
        ws.merge_cells(f"A{last}:E{last}")
        c = ws.cell(last, 1, "توقيع الأستاذ(ة): ___________________________")
        c.alignment = _RA; c.font = Font(size=10, name="Arial")
        ws.row_dimensions[last].height = 30

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Validation ────────────────────────────────────────────────────────────────

def validate(data: InputData) -> list:
    errors = []
    if len(data.teachers) < 2:
        errors.append("يجب إدخال أستاذين على الأقل")
    if not data.rooms:
        errors.append("يجب إدخال قاعة واحدة على الأقل")
    if not data.sessions:
        errors.append("يجب إدخال حصة واحدة على الأقل")
    if not data.active_pairs:
        errors.append("لا توجد أزواج نشطة في المصفوفة")

    from collections import Counter
    rooms_per_sess = Counter(s for (r, s) in data.active_pairs)
    N = len(data.teachers)

    # Per-session: check free pool vs needed supervisors + backups
    # backup slots per session ≈ N × rooms_s / total_rooms (floor/ceil)
    total_rooms = sum(rooms_per_sess.values())
    for s, cnt in rooms_per_sess.items():
        mou_count  = len(data.moudawim_fixed.get(s, []))
        free_pool  = N - mou_count
        bup_target = math.ceil(N * cnt / total_rooms) if total_rooms else 0
        needed     = 2 * cnt + bup_target
        if free_pool < needed:
            errors.append(
                f"الحصة «{s}»: المراقبون المتاحون ({free_pool}) "
                f"أقل من المطلوب ({cnt}×2 + {bup_target} احتياطي = {needed}). "
                f"قلّل المداومين أو أضف أساتذة."
            )

    # Global: N must equal total backup slots
    total_backup = sum(
        math.ceil(N * rooms_per_sess[s] / total_rooms) if total_rooms else 0
        for s in rooms_per_sess
    )
    n_active = sum(1 for s in rooms_per_sess if rooms_per_sess[s] > 0)
    if N < n_active:
        errors.append(f"عدد الأساتذة ({N}) أقل من عدد الحصص النشطة ({n_active})")

    return errors


# ── Scheduler runner ──────────────────────────────────────────────────────────

def run_scheduler(data: InputData):
    errors = validate(data)
    if errors:
        for e in errors:
            st.error(f"❌  {e}")
        return

    with st.spinner("⏳  جارٍ حساب التوزيع الأمثل…"):
        try:
            schedule = build_schedule(
                teachers        = data.teachers,
                sessions        = data.sessions,
                rooms           = data.rooms,
                active_pairs    = data.active_pairs,
                teacher_subjects= data.teacher_subjects,
                session_subjects= data.session_subjects,
                moudawim_fixed  = data.moudawim_fixed,
                time_limit      = 20,
            )
            st.session_state.result = (data, schedule)
            go("results")
        except Exception as err:
            st.error(f"❌  {err}")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: UPLOAD
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "upload":
    st.markdown("""
    <div class="app-header">
      <h1>🎓 جدول توزيع المراقبين</h1>
      <div class="sub">ارفع ملف القالب لتوليد الجدول تلقائياً</div>
    </div>""", unsafe_allow_html=True)

    col_info, col_dl = st.columns([3, 1], gap="large")

    with col_info:
        st.markdown("""<div class="card">
        <div class="card-title">📋 طريقة الاستخدام</div>
        <ol style="direction:rtl;line-height:2.2;font-size:.95rem">
          <li>حمّل القالب الرسمي بالضغط على الزر المجاور</li>
          <li><b>الإعدادات</b> — أدخل عدد القاعات والحصص</li>
          <li><b>الأساتذة</b> — أدخل الأسماء والمواد وأرقام التأجير</li>
          <li><b>الحصص</b> — أدخل التاريخ والتوقيت والمواد وعدد القاعات العاملة</li>
          <li><b>المداومون</b> — اختر مداوماً لكل مادة في كل حصة من القائمة المنسدلة</li>
          <li><b>المصفوفة</b> — تُحسب تلقائياً</li>
          <li>ارفع الملف وولّد الجدول</li>
        </ol>
        </div>""", unsafe_allow_html=True)

    with col_dl:
        st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)
        import pathlib
        tpl = pathlib.Path(__file__).parent / "قالب_جدول_المراقبة.xlsx"
        if tpl.exists():
            st.download_button(
                "⬇️  تحميل القالب",
                data=open(tpl, "rb").read(),
                file_name="قالب_جدول_المراقبة.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.markdown("<div style='margin-top:.5rem'></div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("📂  ارفع ملف Excel هنا", type=["xlsx"])

    if uploaded:
        import tempfile, pathlib as pl
        tmp = pl.Path(tempfile.gettempdir()) / uploaded.name
        tmp.write_bytes(uploaded.getbuffer())
        try:
            data = read_input_workbook(tmp)
            st.session_state.uploaded_data = data

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("👨‍🏫 أستاذ",      len(data.teachers))
            c2.metric("🚪 قاعة",          len(data.rooms))
            c3.metric("📅 حصة",            len(data.sessions))
            c4.metric("✅ زوج نشط",        len(data.active_pairs))
            c5.metric("🎓 مداوم",          sum(len(v) for v in data.moudawim_fixed.values()))

            _, bc, _ = st.columns([1, 2, 1])
            with bc:
                if st.button("⚡  توليد الجدول", type="primary", use_container_width=True):
                    run_scheduler(data)
        except Exception as e:
            st.error(f"❌  {e}")


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
        go("upload")

    data, schedule = st.session_state.result

    bup_per_sess = {s: len(schedule.backups.get(s, [])) for s in data.sessions}
    avg_load     = round(sum(schedule.load.values()) / max(len(schedule.load), 1), 1)

    st.success("✅  تم توليد الجدول بنجاح!")

    st.markdown(f"""
    <div class="mrow">
      <div class="mc"><div class="mv">{len(data.teachers)}</div><div class="ml">👨‍🏫 أستاذ</div></div>
      <div class="mc"><div class="mv">{sum(len(v) for v in schedule.moudawim.values())}</div><div class="ml">🎓 مداوم</div></div>
      <div class="mc"><div class="mv">{len(data.teachers)}</div><div class="ml">🔁 إجمالي الاحتياطيين</div></div>
      <div class="mc"><div class="mv">{len(data.active_pairs)}</div><div class="ml">✅ زوج نشط</div></div>
      <div class="mc"><div class="mv">{avg_load}</div><div class="ml">📊 متوسط الحصص</div></div>
    </div>""", unsafe_allow_html=True)

    # Moudawim table
    if any(schedule.moudawim.values()):
        st.markdown('<div class="rtitle">🎓 المداومون لكل حصة</div>', unsafe_allow_html=True)
        rows = []
        for s in data.sessions:
            for t in schedule.moudawim.get(s, []):
                rows.append({
                    "الحصة": s,
                    "التاريخ": data.session_dates.get(s, ""),
                    "المواد": "، ".join(data.session_subjects.get(s, [])),
                    "المداوم": t,
                    "مادته": data.teacher_subjects.get(t, ""),
                })
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

    # Distribution grid
    st.markdown('<div class="rtitle">🗺️ شبكة التوزيع</div>', unsafe_allow_html=True)
    grid = {}
    for room in data.rooms:
        grid[room] = [
            " / ".join(schedule.supervisors.get((room, s), ["—"]))
            if (room, s) in data.active_pairs else "—"
            for s in data.sessions
        ]
    df_grid = pd.DataFrame(grid, index=data.sessions).T
    df_grid.index.name = "القاعة"
    st.dataframe(df_grid, use_container_width=True)

    # Workload chart
    st.markdown('<div class="rtitle">📊 عبء العمل</div>', unsafe_allow_html=True)
    df_load = (
        pd.DataFrame(list(schedule.load.items()), columns=["الأستاذ", "الحصص"])
        .sort_values("الحصص", ascending=False)
    )
    st.bar_chart(df_load.set_index("الأستاذ"), use_container_width=True)

    # Downloads
    st.markdown('<div class="rtitle">⬇️ التحميلات</div>', unsafe_allow_html=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.markdown("**📋 جدول التوزيع**")
        st.caption("شبكة القاعات × الحصص مع المداومين والاحتياطيين")
        st.download_button(
            "⬇️  تحميل جدول التوزيع",
            data=build_distribution_excel(data, schedule),
            file_name=f"jadwal_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dl2:
        st.markdown("**🗂️ الخريطة التفصيلية**")
        st.caption("صف لكل مهمة: أستاذ · مادة · قاعة · حصة · دور")
        st.download_button(
            "⬇️  تحميل الخريطة",
            data=build_mapping_excel(data, schedule),
            file_name=f"kharita_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dl3:
        st.markdown("**📨 إشعارات المراقبة**")
        st.caption("مراقبة + احتياط + مداوم — بدون ذكر القاعة")
        st.download_button(
            "⬇️  تحميل الإشعارات",
            data=build_convocations_excel(data, schedule),
            file_name=f"ish3arat_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
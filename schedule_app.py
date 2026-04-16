"""
Exam Scheduler — v3
• Welcome page explaining both input modes
• Mode A : Excel upload
• Mode B : Quick entry (count → auto-names → editable matrix)
• Results : distribution grid + workload chart
• Downloads : distribution Excel, teacher→room+session mapping, convocations
"""
from __future__ import annotations
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import pandas as pd
import streamlit as st
from excel_io import InputData, ScheduleOutput, read_input_workbook
from scheduler import build_schedule

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="جدول المراقبة", page_icon="🎓",
                   layout="wide", initial_sidebar_state="collapsed")

# ─── CSS ─────────────────────────────────────────────────────────────────────
_CSS = """
:root{--navy:#0f1e3c;--navy2:#162447;--gold:#c9963a;--gold2:#e8b84b;
--cream:#f2ede4;--white:#fff;--text:#1a1a2e;--muted:#6b7280;
--success:#16a34a;--error:#dc2626;--radius:14px;
--shadow:0 4px 24px rgba(15,30,60,.09);}

html,body,[class*="css"]{font-family:'Tajawal',sans-serif!important;}
.stApp{background:var(--cream);direction:rtl;}

/* ── header ── */
.app-header{background:linear-gradient(135deg,var(--navy) 0%,#1e3a8a 60%,var(--navy2) 100%);
border-radius:var(--radius);padding:1.8rem 2.5rem;margin-bottom:1.8rem;
box-shadow:var(--shadow);position:relative;overflow:hidden;}
.app-header::before{content:'';position:absolute;top:-50px;left:-50px;width:200px;height:200px;
border-radius:50%;background:rgba(201,150,58,.12);pointer-events:none;}
.app-header h1{color:#fff!important;font-size:1.8rem!important;font-weight:900!important;
margin:0!important;padding:0!important;}
.app-header .sub{color:var(--gold2);font-size:.9rem;margin-top:.3rem;}

/* ── card ── */
.card{background:var(--white);border-radius:var(--radius);padding:1.6rem 1.8rem;
margin-bottom:1.2rem;box-shadow:var(--shadow);direction:rtl;}
.card-title{font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:1rem;
padding-bottom:.6rem;border-bottom:2px solid var(--gold);}

/* ── option cards (welcome page) ── */
.opt-card{background:var(--white);border-radius:var(--radius);padding:2rem;
box-shadow:var(--shadow);border-top:5px solid var(--gold);text-align:right;
transition:transform .2s,box-shadow .2s;}
.opt-card:hover{transform:translateY(-4px);box-shadow:0 8px 32px rgba(15,30,60,.14);}
.opt-card .oc-icon{font-size:2.6rem;margin-bottom:.8rem;}
.opt-card h3{font-size:1.15rem;font-weight:800;color:var(--navy);margin:0 0 .5rem;}
.opt-card p{font-size:.88rem;color:var(--muted);line-height:1.6;}
.opt-card ul{font-size:.85rem;color:var(--text);padding-right:1.2rem;margin-top:.5rem;line-height:1.8;}

/* ── block title ── */
.blk{display:flex;align-items:center;gap:.6rem;font-size:1rem;font-weight:700;
color:var(--navy);margin-bottom:.9rem;padding-bottom:.6rem;border-bottom:1px solid #e5e7eb;}
.blk .bn{width:26px;height:26px;border-radius:50%;background:var(--navy);color:#fff;
display:flex;align-items:center;justify-content:center;font-size:.78rem;font-weight:700;flex-shrink:0;}

/* ── metric cards ── */
.mrow{display:flex;gap:1rem;margin-bottom:1.6rem;direction:rtl;}
.mc{flex:1;background:var(--white);border-radius:var(--radius);padding:1rem 1.3rem;
box-shadow:var(--shadow);border-top:4px solid var(--gold);text-align:right;}
.mc .mv{font-size:1.9rem;font-weight:900;color:var(--navy);line-height:1;}
.mc .ml{font-size:.8rem;color:var(--muted);margin-top:.2rem;}

/* ── result title ── */
.rtitle{font-size:1rem;font-weight:700;color:var(--navy);
margin-bottom:.8rem;margin-top:1.4rem;padding-bottom:.5rem;border-bottom:2px solid var(--gold);}

/* ── buttons ── */
.stButton>button{border-radius:10px!important;font-family:'Tajawal',sans-serif!important;
font-weight:600!important;transition:all .18s!important;}
.stButton>button:hover{transform:translateY(-1px)!important;}
.stDownloadButton>button{background:var(--navy)!important;color:#fff!important;border:none!important;
border-radius:10px!important;font-family:'Tajawal',sans-serif!important;font-weight:600!important;}

/* ── inputs ── */
.stTextInput input,.stNumberInput input{border-radius:10px!important;direction:rtl!important;
font-family:'Tajawal',sans-serif!important;}
.stTextInput input:focus,.stNumberInput input:focus{
border-color:var(--gold)!important;box-shadow:0 0 0 3px rgba(201,150,58,.15)!important;}

/* ── checkbox ── */
.stCheckbox label{direction:rtl!important;font-family:'Tajawal',sans-serif!important;}

/* ── alerts ── */
[data-testid="stAlert"]{border-radius:var(--radius)!important;
font-family:'Tajawal',sans-serif!important;direction:rtl;}

/* ── dataframe ── */
.stDataFrame{border-radius:var(--radius)!important;overflow:hidden!important;}

/* hide branding */
#MainMenu,footer,header{visibility:hidden;}
"""
st.markdown('<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&display=swap" rel="stylesheet">',
            unsafe_allow_html=True)
st.markdown(f"<style>{_CSS}</style>", unsafe_allow_html=True)

# ─── Session state defaults ───────────────────────────────────────────────────
_DEFAULTS = dict(
    page="welcome",       # welcome | input_excel | input_quick | results
    input_mode=None,      # "excel" | "quick"
    teachers=[], rooms=[], sessions=[],
    matrix={},            # {(room, session): bool}
    result=None,          # (InputData, ScheduleOutput)
)
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def go(page): st.session_state.page = page; st.rerun()

def sync_matrix():
    valid = {(r, s) for r in st.session_state.rooms for s in st.session_state.sessions}
    for k in list(st.session_state.matrix):
        if k not in valid: del st.session_state.matrix[k]
    for k in valid:
        if k not in st.session_state.matrix: st.session_state.matrix[k] = False


# ─── Excel helpers ────────────────────────────────────────────────────────────
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_THIN   = Border(left=Side(style="thin"), right=Side(style="thin"),
                 top=Side(style="thin"),  bottom=Side(style="thin"))
_HDR_F  = PatternFill("solid", fgColor="0F1E3C")
_HDR_FT = Font(color="FFFFFF", bold=True)
_GOLD_F = PatternFill("solid", fgColor="C9963A")
_BUP_F  = PatternFill("solid", fgColor="FEF9EE")
_ROW_F  = PatternFill("solid", fgColor="F8FAFC")

def _hdr(ws, row, col, val):
    c = ws.cell(row, col, val)
    c.fill=_HDR_F; c.font=_HDR_FT; c.alignment=_CENTER; c.border=_THIN
    return c

def _cell(ws, row, col, val=""):
    c = ws.cell(row, col, val)
    c.alignment=_CENTER; c.border=_THIN
    return c

def build_distribution_excel(data: InputData, out: ScheduleOutput) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "التوزيع"
    _hdr(ws,1,1,"القاعة / الحصة")
    for j,s in enumerate(data.sessions,2): _hdr(ws,1,j,s)
    # backups
    c=ws.cell(2,1,"الاحتياطيون"); c.font=Font(italic=True); c.alignment=_CENTER; c.border=_THIN; c.fill=_BUP_F
    for j,s in enumerate(data.sessions,2):
        c=ws.cell(2,j," / ".join(out.backups.get(s,[]))); c.font=Font(italic=True)
        c.alignment=_CENTER; c.border=_THIN; c.fill=_BUP_F
    for i,room in enumerate(data.rooms,3):
        c=ws.cell(i,1,room); c.font=Font(bold=True); c.alignment=_CENTER; c.border=_THIN; c.fill=_ROW_F
        for j,session in enumerate(data.sessions,2):
            v=" / ".join(out.supervisors.get((room,session),[])) if (room,session) in data.active_pairs else "—"
            _cell(ws,i,j,v)
    ws.freeze_panes="B3"
    ws.column_dimensions["A"].width=22
    for j in range(2,len(data.sessions)+2):
        ws.column_dimensions[ws.cell(1,j).column_letter].width=24
    # workload sheet
    ws2=wb.create_sheet("الحمل الوظيفي")
    _hdr(ws2,1,1,"الأستاذ"); _hdr(ws2,1,2,"عدد الحصص")
    for i,(t,load) in enumerate(sorted(out.load.items(),key=lambda x:-x[1]),2):
        _cell(ws2,i,1,t); _cell(ws2,i,2,load)
    ws2.column_dimensions["A"].width=26; ws2.column_dimensions["B"].width=16
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_mapping_excel(data: InputData, out: ScheduleOutput) -> bytes:
    """One row per (teacher, room, session) assignment."""
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="الخريطة التفصيلية"
    _hdr(ws,1,1,"الأستاذ"); _hdr(ws,1,2,"القاعة"); _hdr(ws,1,3,"الحصة")
    row=2
    for t in sorted(data.teachers):
        for (room,session),sups in sorted(out.supervisors.items(),key=lambda x:x[0][1]):
            if t in sups:
                _cell(ws,row,1,t); _cell(ws,row,2,room); _cell(ws,row,3,session)
                row+=1
    ws.column_dimensions["A"].width=26; ws.column_dimensions["B"].width=22; ws.column_dimensions["C"].width=22
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_convocations_excel(data: InputData, out: ScheduleOutput) -> bytes:
    """One sheet per teacher — lists their sessions (NO room info)."""
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    date_str=datetime.now().strftime("%d/%m/%Y")
    for t in sorted(data.teachers):
        sessions_assigned = sorted({session for (room,session),sups in out.supervisors.items() if t in sups})
        ws=wb.create_sheet(t[:28])  # sheet name ≤ 31 chars
        # Title band
        ws.merge_cells("A1:C1")
        c=ws["A1"]; c.value="إشعار بمهمة المراقبة"; c.font=Font(bold=True,size=14,color="FFFFFF")
        c.fill=_HDR_F; c.alignment=_CENTER
        ws.row_dimensions[1].height=32
        # Teacher name
        ws.merge_cells("A2:C2")
        c=ws["A2"]; c.value=f"الأستاذ(ة): {t}"
        c.font=Font(bold=True,size=12); c.alignment=_CENTER; c.fill=_GOLD_F; c.border=_THIN
        ws.row_dimensions[2].height=26
        # Date
        ws.merge_cells("A3:C3")
        c=ws["A3"]; c.value=f"التاريخ: {date_str}"
        c.font=Font(size=10,italic=True); c.alignment=_CENTER; c.border=_THIN
        # Sessions header
        ws.merge_cells("A4:C4")
        c=ws["A4"]; c.value="قائمة الحصص المُسنَدة"
        c.font=Font(bold=True,size=11,color="FFFFFF"); c.fill=_HDR_F
        c.alignment=_CENTER; c.border=_THIN
        ws.row_dimensions[4].height=24
        # Column headers
        _hdr(ws,5,1,"الرقم"); _hdr(ws,5,2,"الحصة"); _hdr(ws,5,3,"الملاحظات")
        ws.column_dimensions["A"].width=8; ws.column_dimensions["B"].width=28; ws.column_dimensions["C"].width=22
        if sessions_assigned:
            for i,s in enumerate(sessions_assigned,1):
                _cell(ws,5+i,1,i); _cell(ws,5+i,2,s); _cell(ws,5+i,3,"")
        else:
            ws.merge_cells(f"A6:C6")
            c=ws["A6"]; c.value="لا توجد حصص مُسنَدة"; c.alignment=_CENTER
            c.font=Font(italic=True,color="6B7280")
        # Signature line
        last=6+max(len(sessions_assigned),1)+1
        ws.merge_cells(f"A{last}:C{last}")
        c=ws.cell(last,1,"توقيع الأستاذ(ة): ___________________________")
        c.alignment=Alignment(horizontal="right"); c.font=Font(size=10)
        ws.row_dimensions[last].height=30
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


# ─── Scheduler runner (shared between both input modes) ──────────────────────
def _run_scheduler():
    teachers     = st.session_state.teachers
    rooms        = st.session_state.rooms
    sessions_    = st.session_state.sessions
    active_pairs = {k for k,v in st.session_state.matrix.items() if v}

    errors=[]
    if len(teachers)<2:  errors.append("يجب إدخال أستاذين على الأقل")
    if not rooms:         errors.append("يجب إدخال قاعة واحدة على الأقل")
    if not sessions_:     errors.append("يجب إدخال حصة واحدة على الأقل")
    if not active_pairs:  errors.append("يجب تفعيل زوج واحد على الأقل في المصفوفة")

    if not errors:
        # Real minimum: for each session, you need enough distinct teachers to
        # cover all active rooms (×2 supervisors) + backup slots in that session.
        # A teacher can only be in ONE room per session, so this is a hard floor.
        from collections import Counter
        BACKUP = 0.2
        rooms_per_sess = Counter(s for (r,s) in active_pairs)
        min_per_sess   = {s: int(2*cnt + 2*cnt*BACKUP + 0.999) for s,cnt in rooms_per_sess.items()}
        bottleneck_sess, bottleneck_val = max(min_per_sess.items(), key=lambda x: x[1])
        if len(teachers) < bottleneck_val:
            errors.append(
                f"عدد الأساتذة غير كافٍ — الحصة «{bottleneck_sess}» تحتاج على الأقل "
                f"{bottleneck_val} أستاذاً ({rooms_per_sess[bottleneck_sess]} قاعة × 2 مراقب + احتياطيون)"
            )
        # Upper bound: need at least 1 real supervision slot per teacher
        max_usable = 2 * len(active_pairs)
        if len(teachers) > max_usable:
            errors.append(
                f"عدد الأساتذة كبير جداً — لا يمكن إسناد حصة مراقبة فعلية لكل أستاذ "
                f"(الحد الأقصى {max_usable} أستاذ لـ {len(active_pairs)} زوج نشط)"
            )

    if errors:
        for e in errors: st.error(f"❌  {e}")
        return

    data=InputData(teachers=teachers,rooms=rooms,sessions=sessions_,active_pairs=active_pairs)
    with st.spinner("⏳  جارٍ حساب التوزيع الأمثل…"):
        try:
            schedule=build_schedule(data.teachers,data.sessions,data.rooms,data.active_pairs,time_limit=20)
            st.session_state.result=(data,schedule)
            st.session_state.input_mode = "quick"
            go("results")
        except Exception as err:
            st.error(f"❌  {err}")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : WELCOME
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "welcome":
    st.markdown("""
    <div class="app-header">
      <h1>🎓 جدول توزيع المراقبين</h1>
      <div class="sub">نظام ذكي لإعداد جداول الامتحانات وتوزيع الأساتذة على القاعات وإصدار الإشعارات</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("### اختر طريقة الإدخال", unsafe_allow_html=False)
    st.markdown("<div style='margin-bottom:.6rem'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown("""
        <div class="opt-card">
          <div class="oc-icon">📂</div>
          <h3>رفع ملف Excel</h3>
          <p>الخيار الأسرع إذا كانت بياناتك جاهزة في ملف Excel.</p>
          <ul>
            <li>يقبل القالب بأوراق TEACHERS, ROOMS, SESSIONS, ROOM_SESSION_MATRIX</li>
            <li>يُحمِّل البيانات دفعةً واحدة</li>
            <li>مناسب للأعداد الكبيرة</li>
          </ul>
        </div>""", unsafe_allow_html=True)
        st.markdown("<div style='margin-top:.8rem'></div>", unsafe_allow_html=True)
        if st.button("📂  ابدأ برفع ملف Excel", use_container_width=True, key="go_excel"):
            go("input_excel")

    with col2:
        st.markdown("""
        <div class="opt-card">
          <div class="oc-icon">⚡</div>
          <h3>إدخال سريع</h3>
          <p>أدخل فقط عدد الأساتذة والقاعات والحصص — تُنشأ الأسماء تلقائياً.</p>
          <ul>
            <li>تُولَّد الأسماء (أستاذ 1، قاعة 1...) فورياً</li>
            <li>يمكن تعديل الأسماء والمصفوفة بحرية</li>
            <li>لا حاجة لأي ملف خارجي</li>
          </ul>
        </div>""", unsafe_allow_html=True)
        st.markdown("<div style='margin-top:.8rem'></div>", unsafe_allow_html=True)
        if st.button("⚡  ابدأ بالإدخال السريع", use_container_width=True, key="go_quick"):
            go("input_quick")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : INPUT — EXCEL MODE
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "input_excel":
    st.markdown("""
    <div class="app-header">
      <h1>📂 رفع ملف Excel</h1>
      <div class="sub">ارفع ملف القالب الخاص بك ثم ولّد الجدول</div>
    </div>""", unsafe_allow_html=True)

    if st.button("← العودة إلى الرئيسية"): go("welcome")
    st.markdown("<div style='margin-top:1rem'></div>", unsafe_allow_html=True)

    st.markdown("""<div class="card">
    <div class="card-title">📋 هيكل الملف المطلوب</div>
    يجب أن يحتوي الملف على أربع أوراق بالضبط:
    <ul style="direction:rtl;line-height:2;margin-top:.5rem">
      <li><b>TEACHERS</b> — عمود A: أسماء الأساتذة</li>
      <li><b>ROOMS</b> — عمود A: أسماء القاعات</li>
      <li><b>SESSIONS</b> — عمود A: أسماء الحصص</li>
      <li><b>ROOM_SESSION_MATRIX</b> — مصفوفة (صفوف=قاعات، أعمدة=حصص)، ضع X أو 1 في الخلايا النشطة</li>
    </ul>
    </div>""", unsafe_allow_html=True)

    uploaded = st.file_uploader("ارفع ملف .xlsx", type=["xlsx"])

    if uploaded:
        import tempfile, pathlib
        tmp = pathlib.Path(tempfile.gettempdir()) / uploaded.name
        tmp.write_bytes(uploaded.getbuffer())
        try:
            data = read_input_workbook(tmp)
            st.session_state.teachers = list(data.teachers)
            st.session_state.rooms    = list(data.rooms)
            st.session_state.sessions = list(data.sessions)
            st.session_state.matrix   = {k: True for k in data.active_pairs}
            st.success(f"✅  تم تحميل الملف — {len(data.teachers)} أستاذ · {len(data.rooms)} قاعة · {len(data.sessions)} حصة · {len(data.active_pairs)} زوج نشط")
            if st.button("⚡  توليد الجدول", use_container_width=False, type="primary"):
                _run_scheduler()
        except Exception as e:
            st.error(f"❌  خطأ في قراءة الملف: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : INPUT — QUICK MODE
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "input_quick":
    st.markdown("""
    <div class="app-header">
      <h1>⚡ الإدخال السريع</h1>
      <div class="sub">حدّد الأعداد، عدّل الأسماء، ثم اضبط المصفوفة</div>
    </div>""", unsafe_allow_html=True)

    if st.button("← العودة إلى الرئيسية"): go("welcome")
    st.markdown("<div style='margin-top:1rem'></div>", unsafe_allow_html=True)

    # ── Step 1: counts ────────────────────────────────────────────────────────
    st.markdown('<div class="blk"><span class="bn">١</span> أدخل الأعداد</div>', unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    n_t = c1.number_input("عدد الأساتذة",  min_value=2, max_value=200, value=max(len(st.session_state.teachers),6),  step=1)
    n_r = c2.number_input("عدد القاعات",   min_value=1, max_value=100, value=max(len(st.session_state.rooms),3),    step=1)
    n_s = c3.number_input("عدد الحصص",    min_value=1, max_value=50,  value=max(len(st.session_state.sessions),4), step=1)

    if c4.button("🔄  إنشاء القوائم", use_container_width=True):
        st.session_state.teachers = [f"أستاذ {i+1}" for i in range(n_t)]
        st.session_state.rooms    = [f"قاعة {i+1}"  for i in range(n_r)]
        st.session_state.sessions = [f"حصة {i+1}"   for i in range(n_s)]
        st.session_state.matrix   = {(r, s): True
                                     for r in st.session_state.rooms
                                     for s in st.session_state.sessions}
        st.rerun()

    if not st.session_state.teachers:
        st.info("👆  أدخل الأعداد ثم اضغط «إنشاء القوائم» للمتابعة")
        st.stop()

    # ── Step 2: edit names ────────────────────────────────────────────────────
    st.markdown("<div style='margin-top:1.4rem'></div>", unsafe_allow_html=True)
    st.markdown('<div class="blk"><span class="bn">٢</span> عدّل الأسماء إذا أردت</div>', unsafe_allow_html=True)

    tab_t, tab_r, tab_s = st.tabs(["👨‍🏫 الأساتذة", "🚪 القاعات", "📅 الحصص"])

    with tab_t:
        cols = st.columns(4)
        new_teachers = []
        for i, t in enumerate(st.session_state.teachers):
            new_teachers.append(cols[i%4].text_input(f"#{i+1}", value=t, key=f"et_{i}", label_visibility="collapsed"))
        if st.button("💾 حفظ أسماء الأساتذة", key="save_t"):
            st.session_state.teachers = [v.strip() or f"أستاذ {i+1}" for i,v in enumerate(new_teachers)]
            sync_matrix(); st.rerun()

    with tab_r:
        cols = st.columns(4)
        new_rooms = []
        for i, r in enumerate(st.session_state.rooms):
            new_rooms.append(cols[i%4].text_input(f"#{i+1}", value=r, key=f"er_{i}", label_visibility="collapsed"))
        if st.button("💾 حفظ أسماء القاعات", key="save_r"):
            st.session_state.rooms = [v.strip() or f"قاعة {i+1}" for i,v in enumerate(new_rooms)]
            sync_matrix(); st.rerun()

    with tab_s:
        cols = st.columns(4)
        new_sessions = []
        for i, s in enumerate(st.session_state.sessions):
            new_sessions.append(cols[i%4].text_input(f"#{i+1}", value=s, key=f"es_{i}", label_visibility="collapsed"))
        if st.button("💾 حفظ أسماء الحصص", key="save_s"):
            st.session_state.sessions = [v.strip() or f"حصة {i+1}" for i,v in enumerate(new_sessions)]
            sync_matrix(); st.rerun()

    # ── Step 3: matrix ────────────────────────────────────────────────────────
    st.markdown("<div style='margin-top:1.6rem'></div>", unsafe_allow_html=True)
    st.markdown('<div class="blk"><span class="bn">٣</span> مصفوفة النشاط — كل القاعات مُفعَّلة بالكامل. ألغِ تحديد ما لا تريده</div>',
                unsafe_allow_html=True)

    sync_matrix()
    rooms    = st.session_state.rooms
    sessions = st.session_state.sessions

    h_cols = st.columns([2]+[1]*len(sessions))
    h_cols[0].markdown("<b style='font-size:.85rem'>القاعة</b>", unsafe_allow_html=True)
    for j,s in enumerate(sessions):
        h_cols[j+1].markdown(f"<div style='text-align:center;font-size:.75rem;font-weight:600;color:var(--navy)'>{s}</div>",
                             unsafe_allow_html=True)
    st.markdown("<hr style='margin:.3rem 0;border-color:#e5e7eb'>", unsafe_allow_html=True)

    for room in rooms:
        r_cols = st.columns([2]+[1]*len(sessions))
        r_cols[0].markdown(f"<div style='font-weight:600;color:#374151;font-size:.88rem;padding-top:.3rem'>{room}</div>",
                          unsafe_allow_html=True)
        for j,session in enumerate(sessions):
            cur = st.session_state.matrix.get((room,session), False)
            checked = r_cols[j+1].checkbox("", value=cur, key=f"mx_{room}_{session}",
                                           label_visibility="collapsed")
            st.session_state.matrix[(room,session)] = checked

    # ── Generate ──────────────────────────────────────────────────────────────
    st.markdown("<div style='margin-top:1.6rem'></div>", unsafe_allow_html=True)
    _, btn_col, _ = st.columns([1,2,1])
    with btn_col:
        if st.button("⚡  توليد جدول المراقبة", use_container_width=True, type="primary"):
            _run_scheduler()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : RESULTS
# ═══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "results":
    st.markdown("""
    <div class="app-header">
      <h1>🎉 الجدول جاهز</h1>
      <div class="sub">يمكنك تحميل الجدول والخريطة والإشعارات من الأزرار أدناه</div>
    </div>""", unsafe_allow_html=True)

    data, schedule = st.session_state.result

    col_back, col_edit = st.columns([1,5])
    with col_back:
        if st.button("← عودة"):
            mode = st.session_state.get("input_mode","quick")
            go("input_excel" if mode=="excel" else "input_quick")

    st.markdown("<div style='margin-top:1rem'></div>", unsafe_allow_html=True)
    st.success("✅  تم توليد الجدول بنجاح بدون أي تعارض!")

    # Metrics
    avg_load = round(sum(schedule.load.values())/max(len(schedule.load),1),1)
    st.markdown(f"""
    <div class="mrow">
      <div class="mc"><div class="mv">{len(data.teachers)}</div><div class="ml">👨‍🏫 أستاذ مشارك</div></div>
      <div class="mc"><div class="mv">{len(data.active_pairs)}</div><div class="ml">✅ زوج قاعة/حصة</div></div>
      <div class="mc"><div class="mv">{avg_load}</div><div class="ml">📊 متوسط الحصص/أستاذ</div></div>
      <div class="mc"><div class="mv">{len(data.sessions)}</div><div class="ml">📅 حصة</div></div>
    </div>""", unsafe_allow_html=True)

    # Distribution grid
    st.markdown('<div class="rtitle">🗺️ شبكة التوزيع</div>', unsafe_allow_html=True)
    grid={room:[" / ".join(schedule.supervisors.get((room,s),["—"])) if (room,s) in data.active_pairs else "—"
               for s in data.sessions] for room in data.rooms}
    df_grid=pd.DataFrame(grid,index=data.sessions).T
    df_grid.index.name="القاعة"
    st.dataframe(df_grid, use_container_width=True)

    # Workload
    st.markdown('<div class="rtitle">📊 عبء العمل</div>', unsafe_allow_html=True)
    df_load=pd.DataFrame(list(schedule.load.items()),columns=["الأستاذ","الحصص"]).sort_values("الحصص",ascending=False)
    st.bar_chart(df_load.set_index("الأستاذ"), use_container_width=True)

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.markdown('<div class="rtitle">⬇️ التحميلات</div>', unsafe_allow_html=True)

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.markdown("**📋 جدول التوزيع الكامل**")
        st.caption("شبكة القاعات × الحصص مع أسماء المراقبين والاحتياطيين")
        st.download_button(
            "⬇️  تحميل جدول التوزيع",
            data=build_distribution_excel(data, schedule),
            file_name=f"jadwal_tawzi3_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dl2:
        st.markdown("**🗂️ خريطة الأستاذ ← القاعة + الحصة**")
        st.caption("صف لكل مهمة: الأستاذ، القاعة، الحصة")
        st.download_button(
            "⬇️  تحميل الخريطة التفصيلية",
            data=build_mapping_excel(data, schedule),
            file_name=f"kharitat_tawzi3_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dl3:
        st.markdown("**📨 إشعارات المراقبة (الحصص فقط)**")
        st.caption("ورقة منفصلة لكل أستاذ — بدون ذكر القاعة")
        st.download_button(
            "⬇️  تحميل الإشعارات",
            data=build_convocations_excel(data, schedule),
            file_name=f"ish3arat_muraqa3a_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
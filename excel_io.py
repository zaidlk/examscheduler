"""Excel I/O — v6.

New template structure (one sheet per concern):
  الإعدادات  : settings (n_rooms in C3, n_sessions in C4)
  الحصص      : sessions  (A=name, B=date, C=time, D=subj1, E=subj2, F=n_active_rooms)
  الأساتذة   : teachers  (A=auto#, B=name, C=subject, D=ID)
  المصفوفة   : room×session matrix (formula-generated, X = active)

Backward compatible: also reads old TEACHERS/ROOMS/SESSIONS/ROOM_SESSION_MATRIX format.
"""
from __future__ import annotations

import openpyxl
from pathlib import Path
from typing import Dict, List, Set, Tuple
from dataclasses import dataclass, field

CENTER      = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = openpyxl.styles.Border(
    left=openpyxl.styles.Side(style="thin"),  right=openpyxl.styles.Side(style="thin"),
    top=openpyxl.styles.Side(style="thin"),   bottom=openpyxl.styles.Side(style="thin"),
)
ITALIC = openpyxl.styles.Font(italic=True)


@dataclass
class InputData:
    teachers:         List[str]
    rooms:            List[str]
    sessions:         List[str]
    active_pairs:     Set[Tuple[str, str]]
    teacher_subjects: Dict[str, str]        = field(default_factory=dict)
    teacher_ids:      Dict[str, str]        = field(default_factory=dict)
    session_dates:    Dict[str, str]        = field(default_factory=dict)
    session_times:    Dict[str, str]        = field(default_factory=dict)
    session_subjects: Dict[str, List[str]]  = field(default_factory=dict)
    moudawim_fixed:   Dict[str, List[str]]  = field(default_factory=dict)  # session -> [teachers]


@dataclass
class ScheduleOutput:
    supervisors: Dict[Tuple[str, str], List[str]]
    backups:     Dict[str, List[str]]
    load:        Dict[str, int]
    moudawim:    Dict[str, List[str]] = None

    def __post_init__(self):
        if self.moudawim is None:
            self.moudawim = {}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _str(v) -> str:
    if v is None:
        return ""
    if hasattr(v, "strftime"):          # datetime / date object
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def _parse_subjects(raw: str) -> List[str]:
    parts = []
    for sep in (",", "،", ";"):
        if sep in raw:
            parts = [s.strip() for s in raw.split(sep)]
            break
    if not parts:
        parts = [raw.strip()]
    return [p for p in parts if p]


# ── New template reader ───────────────────────────────────────────────────────

def _read_new_format(wb) -> InputData:
    """Read الإعدادات / الحصص / الأساتذة / المصفوفة."""

    # Settings
    ws_cfg = _sheet(wb, "الإعدادات", "Settings")
    n_rooms = int(ws_cfg["C3"].value or 0)
    rooms   = [f"القاعة {i+1}" for i in range(n_rooms)]

    # Teachers
    ws_t = _sheet(wb, "الأساتذة", "TEACHERS")
    teachers, teacher_subjects, teacher_ids = [], {}, {}
    for row in ws_t.iter_rows(min_row=2, values_only=True):
        name = _str(row[1]) if len(row) > 1 else ""  # col B
        subj = _str(row[2]) if len(row) > 2 else ""  # col C
        tid  = _str(row[3]) if len(row) > 3 else ""  # col D
        if name:
            teachers.append(name)
            teacher_subjects[name] = subj
            teacher_ids[name]      = tid

    # Sessions — read each row; col A may be a formula (→ None if not recalculated)
    ws_s = _sheet(wb, "الحصص", "SESSIONS")
    n_sess = int(ws_cfg["C4"].value or 0)  # from الإعدادات

    sessions, session_dates, session_times, session_subjects = [], {}, {}, {}
    session_active_rooms_raw = {}  # {index: n_rooms} for fallback
    row_idx = 0
    for row in ws_s.iter_rows(min_row=2, values_only=True):
        sname = _str(row[0]) if len(row) > 0 else ""
        date  = _str(row[1]) if len(row) > 1 else ""
        time_ = _str(row[2]) if len(row) > 2 else ""
        s1    = _str(row[3]) if len(row) > 3 else ""
        s2    = _str(row[4]) if len(row) > 4 else ""
        try:
            n_act = int(row[5]) if len(row) > 5 and row[5] is not None else 0
        except (TypeError, ValueError):
            n_act = 0

        # If formula not evaluated, generate name from row index
        if not sname and n_sess > 0 and row_idx < n_sess:
            sname = f"حصة {row_idx + 1}"

        if sname and row_idx < n_sess:
            sessions.append(sname)
            session_dates[sname]    = date
            session_times[sname]    = time_
            session_subjects[sname] = [x for x in [s1, s2] if x]
            session_active_rooms_raw[sname] = n_act
        row_idx += 1
        if row_idx >= n_sess and not sname:
            break

    # Matrix — try formula sheet first, fall back to computing from الحصص
    ws_m = _sheet(wb, "المصفوفة", "ROOM_SESSION_MATRIX")
    active_pairs = _read_matrix(ws_m) if ws_m is not None else set()

    if not active_pairs:
        # Formulas not evaluated — use the active rooms count we already parsed
        active_pairs = _compute_matrix_from_sessions(rooms, sessions, session_active_rooms_raw)

    if not active_pairs:
        raise ValueError(
            "لا توجد أزواج نشطة — تحقق من عمود 'عدد القاعات العاملة' في ورقة الحصص."
        )

    moudawim_fixed = _read_moudawim(wb)

    return InputData(
        teachers=teachers, rooms=rooms, sessions=sessions,
        active_pairs=active_pairs,
        teacher_subjects=teacher_subjects,
        teacher_ids=teacher_ids,
        session_dates=session_dates,
        session_times=session_times,
        session_subjects=session_subjects,
        moudawim_fixed=moudawim_fixed,
    )


def _read_matrix(ws) -> Set[Tuple[str, str]]:
    """Read a room×session matrix where headers are in row 1 / col 1."""
    session_names = []
    for cell in ws[1]:
        if cell.column == 1:
            continue
        v = _str(cell.value)
        session_names.append(v)

    active: Set[Tuple[str, str]] = set()
    for row in ws.iter_rows(min_row=2):
        room_id = _str(row[0].value)
        if not room_id:
            continue
        for idx, cell in enumerate(row[1:]):
            v = _str(cell.value).upper()
            if v in ("X", "1", "TRUE", "✔", "OUI", "YES"):
                if idx < len(session_names) and session_names[idx]:
                    active.add((room_id, session_names[idx]))
    return active


def _compute_matrix_from_sessions(rooms, sessions, session_active_rooms) -> Set[Tuple[str, str]]:
    """Fallback: build matrix from الحصص data when formula values aren't computed."""
    active: Set[Tuple[str, str]] = set()
    for s in sessions:
        n = session_active_rooms.get(s, 0)
        for room in rooms[:n]:
            active.add((room, s))
    return active


# ── Old format reader (backward compat) ──────────────────────────────────────

def _read_old_format(wb) -> InputData:
    """Read legacy TEACHERS / ROOMS / SESSIONS / ROOM_SESSION_MATRIX sheets."""

    # Teachers
    teachers, teacher_subjects = [], {}
    for row in wb["TEACHERS"].iter_rows(min_row=2, values_only=True):
        name = _str(row[0]) if len(row) > 0 else ""
        subj = _str(row[1]) if len(row) > 1 else ""
        if name:
            teachers.append(name)
            teacher_subjects[name] = subj

    rooms = [_str(r[0]) for r in wb["ROOMS"].iter_rows(min_row=2, values_only=True)
             if r[0]]

    # Sessions
    sessions, session_dates, session_subjects = [], {}, {}
    for row in wb["SESSIONS"].iter_rows(min_row=2, values_only=True):
        name = _str(row[0]) if len(row) > 0 else ""
        if not name:
            continue
        date = _str(row[1]) if len(row) > 1 else ""
        raw_subjs = [_str(v) for v in row[2:] if v]
        subjs = []
        for r in raw_subjs:
            subjs.extend(_parse_subjects(r))
        sessions.append(name)
        session_dates[name]    = date
        session_subjects[name] = list(dict.fromkeys(subjs))

    active_pairs = _read_matrix(wb["ROOM_SESSION_MATRIX"])
    if not active_pairs:
        raise ValueError("ROOM_SESSION_MATRIX est vide.")

    moudawim_fixed = _read_moudawim(wb)

    return InputData(
        teachers=teachers, rooms=rooms, sessions=sessions,
        active_pairs=active_pairs,
        teacher_subjects=teacher_subjects,
        session_dates=session_dates,
        session_subjects=session_subjects,
        moudawim_fixed=moudawim_fixed,
    )


# ── Public API ────────────────────────────────────────────────────────────────

def _read_moudawim(wb) -> "Dict[str, List[str]]":
    """Read المداومون sheet: {session -> [teacher, ...]} (one per subject row)."""
    ws = _sheet(wb, "المداومون")
    if ws is None:
        return {}
    result: dict = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        session = _str(row[0]) if len(row) > 0 else ""
        teacher = _str(row[3]) if len(row) > 3 else ""
        if session and teacher:
            result.setdefault(session, [])
            if teacher not in result[session]:
                result[session].append(teacher)
    return result


def _sheet(wb, *candidates):
    """Return sheet by trying candidate names (handles unicode normalization)."""
    import unicodedata
    sheets_nfc = {unicodedata.normalize("NFC", s): s for s in wb.sheetnames}
    for name in candidates:
        nfc = unicodedata.normalize("NFC", name)
        if nfc in sheets_nfc:
            return wb[sheets_nfc[nfc]]
    return None


def read_input_workbook(path: Path) -> InputData:
    wb = openpyxl.load_workbook(path, data_only=True)

    # Detect format: try new Arabic sheets first, fall back to legacy English
    ws_cfg  = _sheet(wb, "الإعدادات", "Settings")
    ws_sess = _sheet(wb, "الحصص",     "SESSIONS")
    ws_teac = _sheet(wb, "الأساتذة",  "TEACHERS")

    if ws_cfg is not None and ws_sess is not None and ws_teac is not None:
        # Check which format based on presence of الإعدادات vs legacy ROOMS sheet
        ws_rooms = _sheet(wb, "ROOMS")
        if ws_rooms is not None and ws_cfg is None:
            return _read_old_format(wb)
        return _read_new_format(wb)

    # Legacy fallback: TEACHERS + ROOMS + SESSIONS + ROOM_SESSION_MATRIX
    ws_t_legacy = _sheet(wb, "TEACHERS")
    ws_s_legacy = _sheet(wb, "SESSIONS")
    if ws_t_legacy is not None and ws_s_legacy is not None:
        return _read_old_format(wb)

    raise ValueError(
        "تنسيق الملف غير معروف — يجب أن يحتوي على:\n"
        "• الإعدادات + الحصص + الأساتذة + المصفوفة  (القالب الجديد)\n"
        "• أو TEACHERS + ROOMS + SESSIONS + ROOM_SESSION_MATRIX  (القالب القديم)\n\n"
        f"الأوراق الموجودة: {', '.join(wb.sheetnames)}"
    )


def write_schedule(wb_path: Path, data: InputData, out: ScheduleOutput, save_to: Path):
    """Kept for backward compat."""
    wb = openpyxl.load_workbook(wb_path)
    for sheet in ("DISTRIBUTION",):
        if sheet in wb:
            del wb[sheet]
    ws = wb.create_sheet("DISTRIBUTION", 0)
    ws.append(["القاعة / الحصة"] + data.sessions)
    for c in ws[1]:
        c.alignment = CENTER; c.border = THIN_BORDER
    ws.append(["الاحتياطيون"] + [" / ".join(out.backups.get(s, [])) for s in data.sessions])
    for c in ws[2]:
        c.font = ITALIC; c.alignment = CENTER; c.border = THIN_BORDER
    for room in data.rooms:
        row_vals = [room]
        for session in data.sessions:
            if (room, session) in data.active_pairs:
                row_vals.append(" / ".join(out.supervisors.get((room, session), [])))
            else:
                row_vals.append("")
        ws.append(row_vals)
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.alignment = CENTER; cell.border = THIN_BORDER
    ws.freeze_panes = "B3"
    wb.save(save_to)
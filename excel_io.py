"""Excel I/O — v7.

Sheets (new template):
  الإعدادات  : settings (C3=n_rooms, C4=n_sessions)
  الحصص      : sessions  (A=name, B=date, C=time, D=subj1, E=subj2, F=n_active_rooms)
  الأساتذة   : teachers  (A=auto#, B=name, C=subject, D=ID)
  المصفوفة   : room×session matrix (formula-generated, X = active)
  المداومون  : manual moudawim (A=session, B=date, C=subject, D=teacher_name)

Also reads legacy TEACHERS/ROOMS/SESSIONS/ROOM_SESSION_MATRIX format.
"""
from __future__ import annotations

import unicodedata
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


# ── Data classes ──────────────────────────────────────────────────────────────

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
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def _sheet(wb, *candidates):
    """Return sheet by name, trying NFC-normalised variants."""
    sheets_nfc = {_nfc(s): s for s in wb.sheetnames}
    for name in candidates:
        key = _nfc(name)
        if key in sheets_nfc:
            return wb[sheets_nfc[key]]
    return None


def _parse_subjects(raw: str) -> List[str]:
    parts = []
    for sep in (",", "،", ";"):
        if sep in raw:
            parts = [s.strip() for s in raw.split(sep)]
            break
    if not parts:
        parts = [raw.strip()]
    return [p for p in parts if p]


# ── Reader: المداومون sheet ───────────────────────────────────────────────────

def _read_moudawim(wb) -> Dict[str, List[str]]:
    """Read المداومون sheet → {session: [teacher, ...]}."""
    ws = _sheet(wb, "المداومون")
    if ws is None:
        return {}
    result: Dict[str, List[str]] = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        session = _str(row[0]) if len(row) > 0 else ""
        teacher = _str(row[3]) if len(row) > 3 else ""
        if session and teacher:
            result.setdefault(session, [])
            if teacher not in result[session]:
                result[session].append(teacher)
    return result


# ── Reader: matrix sheet ──────────────────────────────────────────────────────

def _read_matrix(ws) -> Set[Tuple[str, str]]:
    session_names = [_str(c.value) for c in ws[1] if c.column > 1]
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


def _compute_matrix_from_sessions(
    rooms: List[str],
    sessions: List[str],
    session_active_rooms: Dict[str, int],
) -> Set[Tuple[str, str]]:
    active: Set[Tuple[str, str]] = set()
    for s in sessions:
        n = session_active_rooms.get(s, 0)
        for room in rooms[:n]:
            active.add((room, s))
    return active


# ── Reader: new Arabic template ───────────────────────────────────────────────

def _read_new_format(wb) -> InputData:
    # Settings
    ws_cfg  = _sheet(wb, "الإعدادات")
    n_rooms = int(ws_cfg["C3"].value or 0)
    rooms   = [f"القاعة {i+1}" for i in range(n_rooms)]

    # Teachers
    ws_t = _sheet(wb, "الأساتذة")
    teachers, teacher_subjects, teacher_ids = [], {}, {}
    for row in ws_t.iter_rows(min_row=2, values_only=True):
        name = _str(row[1]) if len(row) > 1 else ""
        subj = _str(row[2]) if len(row) > 2 else ""
        tid  = _str(row[3]) if len(row) > 3 else ""
        if name:
            teachers.append(name)
            teacher_subjects[name] = subj
            teacher_ids[name]      = tid

    # Sessions
    ws_s = _sheet(wb, "الحصص")
    n_sess = int(ws_cfg["C4"].value or 0)
    sessions, session_dates, session_times = [], {}, {}
    session_subjects: Dict[str, List[str]] = {}
    session_active_rooms_raw: Dict[str, int] = {}

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

    # Matrix
    ws_m = _sheet(wb, "المصفوفة")
    active_pairs = _read_matrix(ws_m) if ws_m is not None else set()
    if not active_pairs:
        active_pairs = _compute_matrix_from_sessions(rooms, sessions, session_active_rooms_raw)
    if not active_pairs:
        raise ValueError("لا توجد أزواج نشطة — تحقق من عمود 'عدد القاعات العاملة' في ورقة الحصص.")

    # Moudawim
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


# ── Reader: legacy English template ──────────────────────────────────────────

def _read_old_format(wb) -> InputData:
    teachers, teacher_subjects = [], {}
    for row in wb["TEACHERS"].iter_rows(min_row=2, values_only=True):
        name = _str(row[0]) if len(row) > 0 else ""
        subj = _str(row[1]) if len(row) > 1 else ""
        if name:
            teachers.append(name)
            teacher_subjects[name] = subj

    rooms = [_str(r[0]) for r in wb["ROOMS"].iter_rows(min_row=2, values_only=True) if r[0]]

    sessions, session_dates, session_subjects = [], {}, {}
    for row in wb["SESSIONS"].iter_rows(min_row=2, values_only=True):
        name = _str(row[0]) if len(row) > 0 else ""
        if not name:
            continue
        date = _str(row[1]) if len(row) > 1 else ""
        raw_subjs = [_str(v) for v in row[2:] if v]
        subjs: List[str] = []
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

def read_input_workbook(path: Path) -> InputData:
    wb    = openpyxl.load_workbook(path, data_only=True)
    names = set(wb.sheetnames)

    has_arabic = (
        _sheet(wb, "الإعدادات") is not None
        and _sheet(wb, "الحصص")    is not None
        and _sheet(wb, "الأساتذة") is not None
    )
    has_legacy = "TEACHERS" in names and "SESSIONS" in names

    if has_arabic:
        return _read_new_format(wb)
    elif has_legacy:
        return _read_old_format(wb)
    else:
        raise ValueError(
            "تنسيق الملف غير معروف.\n"
            f"الأوراق الموجودة: {', '.join(wb.sheetnames)}"
        )


def write_schedule(wb_path: Path, data: InputData, out: ScheduleOutput, save_to: Path):
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
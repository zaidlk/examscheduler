"""Excel I/O — v5.

Changes from v4:
  • session_subjects is now Dict[str, List[str]]  (multiple subjects per session)
  • SESSIONS sheet: col A = name, col B = date, col C onwards = subjects
    (or a single col C with comma-separated subjects)
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
    teacher_subjects: Dict[str, str]        = field(default_factory=dict)  # teacher -> subject
    session_dates:    Dict[str, str]        = field(default_factory=dict)  # session -> date
    session_subjects: Dict[str, List[str]]  = field(default_factory=dict)  # session -> [subj, ...]


@dataclass
class ScheduleOutput:
    supervisors: Dict[Tuple[str, str], List[str]]
    backups:     Dict[str, List[str]]
    load:        Dict[str, int]
    moudawim:    Dict[str, List[str]] = None   # session -> [teachers assigned as مداوم]

    def __post_init__(self):
        if self.moudawim is None:
            self.moudawim = {}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _str(v) -> str:
    return str(v).strip() if v is not None else ""


def _parse_subjects(raw: str) -> List[str]:
    """Split comma/semicolon-separated subjects, strip blanks."""
    parts = []
    for sep in (",", "،", ";"):
        if sep in raw:
            parts = [s.strip() for s in raw.split(sep)]
            break
    if not parts:
        parts = [raw.strip()]
    return [p for p in parts if p]


def _read_teachers(ws) -> Tuple[List[str], Dict[str, str]]:
    teachers, subjects = [], {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = _str(row[0]) if len(row) > 0 else ""
        subj = _str(row[1]) if len(row) > 1 else ""
        if name:
            teachers.append(name)
            subjects[name] = subj
    return teachers, subjects


def _read_sessions(ws) -> Tuple[List[str], Dict[str, str], Dict[str, List[str]]]:
    """
    Supports two formats for subjects:
      • Single col C with comma-separated subjects: "رياضيات, فيزياء"
      • Multiple cols C, D, E…: one subject per column
    """
    sessions, dates, subjects = [], {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = _str(row[0]) if len(row) > 0 else ""
        if not name:
            continue
        date = _str(row[1]) if len(row) > 1 else ""
        # collect subjects from col C onwards
        raw_subjects: List[str] = []
        for cell_val in row[2:]:
            v = _str(cell_val)
            if v:
                raw_subjects.extend(_parse_subjects(v))
        sessions.append(name)
        dates[name]    = date
        subjects[name] = list(dict.fromkeys(raw_subjects))  # deduplicate, preserve order
    return sessions, dates, subjects


def _read_col_a(ws) -> List[str]:
    return [_str(row[0]) for row in ws.iter_rows(min_row=2, values_only=True) if row[0]]


def read_room_session_matrix(ws) -> Set[Tuple[str, str]]:
    sessions = [_str(c.value) for c in ws[1] if c.column > 1 and c.value]
    active: Set[Tuple[str, str]] = set()
    for r in ws.iter_rows(min_row=2, min_col=1, max_col=len(sessions) + 1):
        room_id = _str(r[0].value)
        if not room_id:
            continue
        for idx, cell in enumerate(r[1:]):
            if _str(cell.value).upper() in ("1", "X", "TRUE", "✔", "OUI", "YES"):
                if idx < len(sessions):
                    active.add((room_id, sessions[idx]))
    return active


def read_input_workbook(path: Path) -> InputData:
    wb = openpyxl.load_workbook(path, data_only=True)
    teachers, teacher_subjects             = _read_teachers(wb["TEACHERS"])
    rooms                                  = _read_col_a(wb["ROOMS"])
    sessions, session_dates, session_subjs = _read_sessions(wb["SESSIONS"])
    active_pairs                           = read_room_session_matrix(wb["ROOM_SESSION_MATRIX"])
    if not active_pairs:
        raise ValueError("ROOM_SESSION_MATRIX est vide — marquez au moins une paire.")
    return InputData(
        teachers=teachers, rooms=rooms, sessions=sessions,
        active_pairs=active_pairs,
        teacher_subjects=teacher_subjects,
        session_dates=session_dates,
        session_subjects=session_subjs,
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
"""Scheduler — CP-SAT v6.

Three roles per session:
  x[(t, r, s)]  : مراقب  — supervises room r in session s
  b[(t, s)]     : احتياطي — backup for session s
  m[(t, s)]     : مداوم  — subject monitor for session s

Moudawim constraints:
  • For each session s, for each examined subject subj:
      – Exactly 1 teacher teaching `subj` is assigned as مداوم
      – NO other teacher teaching `subj` may be مراقب or احتياطي in s
  • مداوم is exclusive: a teacher cannot also be مراقب or احتياطي in the same session.

Backups: ceil(22% × 2 × active_rooms_in_session)  (varies per session)
Every teacher must have ≥1 real duty of any kind across all sessions.
"""
from __future__ import annotations

import math
from typing import Dict, List, Set, Tuple

from ortools.sat.python import cp_model
from excel_io import ScheduleOutput

BACKUP_RATIO = 0.22


def build_schedule(
    teachers:         List[str],
    sessions:         List[str],
    rooms:            List[str],
    active_pairs:     Set[Tuple[str, str]],
    teacher_subjects: Dict[str, str]       | None = None,
    session_subjects: Dict[str, List[str]] | None = None,
    time_limit:       int = 20,
) -> ScheduleOutput:

    teacher_subjects = teacher_subjects or {}
    session_subjects = session_subjects or {}

    model = cp_model.CpModel()

    # Pre-compute per-session info
    rooms_in_session: Dict[str, List[str]] = {
        s: [r for r in rooms if (r, s) in active_pairs]
        for s in sessions
    }
    # For each session, which teachers are "subject-locked"
    # (teach a subject that is examined in that session)?
    subject_locked: Dict[str, Set[str]] = {}
    for s in sessions:
        locked = set()
        for subj in session_subjects.get(s, []):
            for t in teachers:
                if teacher_subjects.get(t, "") == subj:
                    locked.add(t)
        subject_locked[s] = locked

    # ── Variables ─────────────────────────────────────────────────────────────

    # x[(t, r, s)] = 1  →  t is مراقب in room r during session s
    x = {
        (t, r, s): model.NewBoolVar(f"x_{t}_{r}_{s}")
        for t in teachers
        for (r, s) in active_pairs
    }

    # b[(t, s)] = 1  →  t is احتياطي for session s
    b = {
        (t, s): model.NewBoolVar(f"b_{t}_{s}")
        for t in teachers
        for s in sessions
    }

    # m[(t, s)] = 1  →  t is مداوم for session s
    m = {
        (t, s): model.NewBoolVar(f"m_{t}_{s}")
        for t in teachers
        for s in sessions
    }

    # ── 1. Exactly 2 مراقب per (room, session) ───────────────────────────────
    for (r, s) in active_pairs:
        model.Add(sum(x[(t, r, s)] for t in teachers) == 2)

    # ── 2. A teacher is in at most 1 room per session ─────────────────────────
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms_in_session[s]]
            if in_rooms:
                model.Add(sum(in_rooms) <= 1)

    # ── 3. Exclusive roles: at most one duty per session per teacher ──────────
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms_in_session[s]]
            model.Add(sum(in_rooms) + b[(t, s)] + m[(t, s)] <= 1)

    # ── 4. Each teacher is احتياطي exactly once (priority) ─────────────────────
    # Distribute N backups proportionally to active rooms per session.
    # target_s = round(N × rooms_s / Σ rooms_all), floor ≤ b_s ≤ ceil
    import math as _math
    total_rooms_all = sum(len(v) for v in rooms_in_session.values())
    N = len(teachers)

    if total_rooms_all > 0:
        # compute floor/ceil targets
        raw = {s: N * len(rooms_in_session[s]) / total_rooms_all for s in sessions}
        floors = {s: int(_math.floor(raw[s])) for s in sessions}
        ceils  = {s: int(_math.ceil(raw[s]))  for s in sessions}

        # adjust so floors sum == N (distribute remainders to largest fractional parts)
        remainder = N - sum(floors.values())
        sorted_sess = sorted(sessions, key=lambda s: -(raw[s] - floors[s]))
        for s in sorted_sess[:remainder]:
            floors[s] += 1
        ceils = {s: max(floors[s], ceils[s]) for s in sessions}

        for s in sessions:
            if rooms_in_session[s]:   # active session
                model.Add(sum(b[(t, s)] for t in teachers) >= floors[s])
                model.Add(sum(b[(t, s)] for t in teachers) <= ceils[s])
            else:
                model.Add(sum(b[(t, s)] for t in teachers) == 0)
    else:
        for s in sessions:
            model.Add(sum(b[(t, s)] for t in teachers) >= 1 if rooms_in_session[s] else
                      model.Add(sum(b[(t, s)] for t in teachers) == 0))

    # ── 5. مداوم constraints ──────────────────────────────────────────────────
    for s in sessions:
        subjs = session_subjects.get(s, [])
        for subj in subjs:
            if not subj:
                continue
            subj_teachers = [t for t in teachers if teacher_subjects.get(t, "") == subj]
            if not subj_teachers:
                continue

            # 5a. Exactly 1 مداوم per subject per session
            model.Add(sum(m[(t, s)] for t in subj_teachers) == 1)

            # 5b. No subject teacher can be مراقب or احتياطي in this session
            for t in subj_teachers:
                # cannot be مراقب in any room
                for r in rooms_in_session[s]:
                    if (t, r, s) in x:
                        model.Add(x[(t, r, s)] == 0)
                # cannot be احتياطي
                model.Add(b[(t, s)] == 0)

        # 5c. Teachers NOT teaching any examined subject cannot be مداوم
        for t in teachers:
            if t not in subject_locked[s]:
                model.Add(m[(t, s)] == 0)

    # ── 6. Each teacher must be احتياطي exactly once across all sessions ────────
    for t in teachers:
        model.Add(sum(b[(t, s)] for s in sessions) == 1)

    # ── 7. Every teacher must have ≥1 duty (of any kind) across all sessions ─
    for t in teachers:
        sup_total = sum(x[(t, r, s)] for (r, s) in active_pairs)
        bup_total = sum(b[(t, s)]    for s in sessions)
        mou_total = sum(m[(t, s)]    for s in sessions)
        model.Add(sup_total + bup_total + mou_total >= 1)

    # ── 8. Balanced load ──────────────────────────────────────────────────────
    # Count مداوم slots: one per (session × subject) where a teacher is assigned
    mou_total_slots = sum(
        len([subj for subj in session_subjects.get(s, []) if subj and
             any(teacher_subjects.get(t, "") == subj for t in teachers)])
        for s in sessions
    )
    total_slots = (
        2 * len(active_pairs)
        + len(teachers)        # exactly 1 backup duty per teacher
        + mou_total_slots
    )
    avg = (total_slots + len(teachers) - 1) // len(teachers)

    load_vars = {}
    sq_diffs  = []
    for t in teachers:
        sup_load = sum(x[(t, r, s)] for (r, s) in active_pairs)
        bup_load = sum(b[(t, s)]    for s in sessions)
        mou_load = sum(m[(t, s)]    for s in sessions)
        load     = sup_load + bup_load + mou_load
        load_vars[t] = load

        model.Add(load <= avg + 1)
        model.Add(load >= max(0, avg - 1))

        diff = model.NewIntVar(-1, 1, f"diff_{t}")
        sq   = model.NewIntVar(0,  1, f"sq_{t}")
        model.Add(diff == load - avg)
        model.AddMultiplicationEquality(sq, diff, diff)
        sq_diffs.append(sq)

    model.Minimize(sum(sq_diffs))

    # ── Solve ─────────────────────────────────────────────────────────────────
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(
            "لم يتمكن السولفر من إيجاد حل — تحقق من عدد الأساتذة ومواد الحصص."
        )

    # ── Extract ───────────────────────────────────────────────────────────────
    supervisors: dict = {}
    for (t, r, s), var in x.items():
        if solver.Value(var):
            supervisors.setdefault((r, s), []).append(t)

    backups: dict = {}
    for (t, s), var in b.items():
        if solver.Value(var):
            backups.setdefault(s, []).append(t)

    moudawim_out: dict = {}
    for (t, s), var in m.items():
        if solver.Value(var):
            moudawim_out.setdefault(s, []).append(t)

    load_out = {t: int(solver.Value(load_vars[t])) for t in teachers}

    return ScheduleOutput(
        supervisors=supervisors,
        backups=backups,
        load=load_out,
        moudawim=moudawim_out,
    )
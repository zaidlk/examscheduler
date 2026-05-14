"""Scheduler — CP-SAT v6.

Three roles per session:
  x[(t, r, s)]  : مراقب  — supervises room r in session s
  b[(t, s)]     : احتياطي — backup for session s
  مداوم are pre-fixed by user input (no solver variable)

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

    # ── 2a. A teacher supervises at most 1 room per session ──────────────────
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms_in_session[s]]
            if in_rooms:
                model.Add(sum(in_rooms) <= 1)

    # ── 2b. A teacher supervises each room at most once across ALL sessions ──
    for t in teachers:
        for r in rooms:
            sessions_for_room = [s for s in sessions if (r, s) in active_pairs]
            if len(sessions_for_room) > 1:
                model.Add(sum(x[(t, r, s)] for s in sessions_for_room) <= 1)

    # ── 3. Exclusive roles: at most one duty per session per teacher ──────────
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms_in_session[s]]
            model.Add(sum(in_rooms) + b[(t, s)] <= 1)

    # ── 4. Each teacher is احتياطي exactly once — distributed proportionally ────
    # Free pool per session = teachers not locked as مداوم in that session.
    # target_s = N × rooms_s / Σ rooms_all, clamped to free pool size.
    # Floors are adjusted so Σ floors == N exactly.

    N             = len(teachers)
    total_rooms_all = sum(len(v) for v in rooms_in_session.values())

    # No active rooms at all → no backups anywhere
    if total_rooms_all == 0:
        for s in sessions:
            model.Add(sum(b[(t, s)] for t in teachers) == 0)
    else:
        # Proportional raw targets
        raw = {
            s: N * len(rooms_in_session[s]) / total_rooms_all
            for s in sessions
        }
        floors = {s: int(math.floor(raw[s])) for s in sessions}
        ceils  = {s: int(math.ceil(raw[s]))  for s in sessions}

        # Distribute remainder to sessions with largest fractional parts
        remainder = N - sum(floors.values())
        for s in sorted(sessions, key=lambda s: -(raw[s] - floors[s]))[:remainder]:
            floors[s] += 1

        # Clamp floor/ceil to the free pool of each session
        # (teachers who are مداوم in session s cannot be backup there)
        for s in sessions:
            free_pool = len([t for t in teachers if t not in subject_locked[s]])
            floors[s] = min(floors[s], free_pool)
            ceils[s]  = max(floors[s], min(ceils[s], free_pool))

        # Hard per-session bounds
        for s in sessions:
            b_sum = sum(b[(t, s)] for t in teachers)
            if rooms_in_session[s]:          # active session
                model.Add(b_sum >= floors[s])
                model.Add(b_sum <= ceils[s])
            else:                            # inactive session → no backups
                model.Add(b_sum == 0)

    # ── 5. مداوم already blocked in step 3 — nothing more needed ──────────────

    # ── 6. Each teacher must be احتياطي exactly once across all sessions ────────
    for t in teachers:
        model.Add(sum(b[(t, s)] for s in sessions) == 1)

    # ── 7. Every teacher must have ≥1 duty across all sessions ─────────────────
    for t in teachers:
        is_moudawim_anywhere = any(t in moudawim_fixed.get(s, []) for s in sessions)
        sup_total = sum(x[(t, r, s)] for (r, s) in active_pairs)
        bup_total = sum(b[(t, s)] for s in sessions)
        if not is_moudawim_anywhere:
            model.Add(sup_total + bup_total >= 1)

    # ── 8. Balanced load ──────────────────────────────────────────────────────
    # Count مداوم slots: one per (session × subject) where a teacher is assigned
    mou_total_slots = sum(len(ts) for ts in moudawim_fixed.values())
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
        bup_load = sum(b[(t, s)] for s in sessions)
        # Fixed مداوم assignments: Python int (not a CP-SAT var)
        mou_fixed = sum(1 for s in sessions if t in moudawim_fixed.get(s, []))
        load_vars[t] = (sup_load + bup_load, mou_fixed)

        # Balance only the solver-assigned portion, accounting for fixed مداوم load
        effective_avg = max(0, avg - mou_fixed)
        solver_load   = sup_load + bup_load
        model.Add(solver_load <= effective_avg + 1)
        model.Add(solver_load >= max(0, effective_avg - 1))

        diff = model.NewIntVar(-1, 1, f"diff_{t}")
        sq   = model.NewIntVar(0,  1, f"sq_{t}")
        model.Add(diff == solver_load - effective_avg)
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

    moudawim_out: dict = {s: list(ts) for s, ts in moudawim_fixed.items() if ts}

    load_out = {
        t: int(solver.Value(load_vars[t][0])) + load_vars[t][1]
        for t in teachers
    }

    return ScheduleOutput(
        supervisors=supervisors,
        backups=backups,
        load=load_out,
        moudawim=moudawim_out,
    )
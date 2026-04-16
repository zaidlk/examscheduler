"""Scheduler – CP-SAT model using ROOM_SESSION_MATRIX active pairs.

Key fix: every teacher must have ≥1 real supervision slot (not backup-only).
"""
from __future__ import annotations
from typing import List, Set, Tuple
from ortools.sat.python import cp_model
from excel_io import ScheduleOutput


def build_schedule(
    teachers: List[str],
    sessions: List[str],
    rooms: List[str],
    active_pairs: Set[Tuple[str, str]],
    backup_ratio: float = 0.2,
    time_limit: int = 20,
) -> ScheduleOutput:

    model = cp_model.CpModel()

    # x[(t, r, s)] = 1 if teacher t supervises room r during session s
    x = {
        (t, r, s): model.NewBoolVar(f"x_{t}_{r}_{s}")
        for t in teachers
        for (r, s) in active_pairs
    }

    # b[(t, s)] = 1 if teacher t is a backup for session s
    b = {
        (t, s): model.NewBoolVar(f"b_{t}_{s}")
        for t in teachers
        for s in sessions
    }

    # ── Core constraints ──────────────────────────────────────────────────────

    # 1. Exactly 2 supervisors per active (room, session)
    for (r, s) in active_pairs:
        model.Add(sum(x[(t, r, s)] for t in teachers) == 2)

    # 2. A teacher supervises each room at most once across all sessions
    for t in teachers:
        for r in rooms:
            room_sessions = [s for (rr, s) in active_pairs if rr == r]
            if room_sessions:
                model.Add(sum(x[(t, r, s)] for s in room_sessions) <= 1)

    # 3. At most one duty per session per teacher (supervise OR backup, not both)
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms if (r, s) in active_pairs]
            model.Add(sum(in_rooms) + b[(t, s)] <= 1)

    # 4. Required number of backups per session
    rooms_per_session = {
        s: sum(1 for r in rooms if (r, s) in active_pairs)
        for s in sessions
    }
    for s in sessions:
        backup_needed = int(2 * rooms_per_session[s] * backup_ratio + 0.999)
        model.Add(sum(b[(t, s)] for t in teachers) == backup_needed)

    # ── KEY FIX: every teacher must have ≥1 real supervision slot ────────────
    # This prevents teachers being assigned only to backup slots.
    for t in teachers:
        model.Add(sum(x[(t, r, s)] for (r, s) in active_pairs) >= 1)

    # ── Balanced load ─────────────────────────────────────────────────────────
    total_slots = (
        2 * len(active_pairs)
        + sum(int(2 * rooms_per_session[s] * backup_ratio + 0.999) for s in sessions)
    )
    avg = (total_slots + len(teachers) - 1) // len(teachers)

    load_vars = {}
    sq_diffs  = []

    for t in teachers:
        sup_load = sum(x[(t, r, s)] for (r, s) in active_pairs)
        bup_load = sum(b[(t, s)] for s in sessions)
        load     = sup_load + bup_load
        load_vars[t] = load

        # Load stays within [avg-1, avg+1]
        model.Add(load <= avg + 1)
        model.Add(load >= avg - 1)

        # Minimise variance: sum of squared deviations
        diff = model.NewIntVar(-1, 1, f"diff_{t}")
        sq   = model.NewIntVar(0, 1, f"sq_{t}")
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
            "لم يتمكن السولفر من إيجاد حل — تحقق من عدد الأساتذة والأزواج النشطة."
        )

    # ── Extract results ───────────────────────────────────────────────────────
    supervisors: dict = {}
    for (t, r, s), var in x.items():
        if solver.Value(var):
            supervisors.setdefault((r, s), []).append(t)

    backups: dict = {}
    for (t, s), var in b.items():
        if solver.Value(var):
            backups.setdefault(s, []).append(t)

    load_out = {t: int(solver.Value(load_vars[t])) for t in teachers}

    return ScheduleOutput(supervisors=supervisors, backups=backups, load=load_out)
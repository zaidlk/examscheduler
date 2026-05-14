"""Scheduler — CP-SAT v7.

Roles:
  x[(t, r, s)]  : مراقب  — supervises room r during session s
  b[(t, s)]     : احتياطي — backup for session s
  مداوم are pre-fixed by the user in the Excel sheet المداومون.
    → excluded from x and b in their session
    → counted as 1 load unit (Python int, not a CP-SAT variable)

Constraints:
  1. Exactly 2 مراقب per active (room, session)
  2a. A teacher is in at most 1 room per session
  2b. A teacher supervises each room at most once across ALL sessions
  3. Exclusive roles per session: مراقب + احتياطي ≤ 1
     (مداوم are blocked from both x and b in their session)
  4. احتياطي distributed proportionally to active rooms per session
     (floor ≤ b_sum_s ≤ ceil), with Σ floors == N exactly
  5. Each teacher is احتياطي exactly once
  6. Every teacher has ≥1 duty (مداوم counts; others need sup+bup ≥ 1)
  7. Balanced load
"""
from __future__ import annotations

import math
from typing import Dict, List, Set, Tuple

from ortools.sat.python import cp_model
from excel_io import ScheduleOutput


def build_schedule(
    teachers:         List[str],
    sessions:         List[str],
    rooms:            List[str],
    active_pairs:     Set[Tuple[str, str]],
    teacher_subjects: Dict[str, str]        | None = None,
    session_subjects: Dict[str, List[str]]  | None = None,
    moudawim_fixed:   Dict[str, List[str]]  | None = None,
    time_limit:       int = 20,
) -> ScheduleOutput:

    teacher_subjects = teacher_subjects or {}
    session_subjects = session_subjects or {}
    moudawim_fixed   = moudawim_fixed   or {}

    model = cp_model.CpModel()

    # Pre-compute useful sets
    rooms_in_session: Dict[str, List[str]] = {
        s: [r for r in rooms if (r, s) in active_pairs]
        for s in sessions
    }

    # moudawim_set: set of (teacher, session) pairs that are pre-fixed
    moudawim_set: Set[Tuple[str, str]] = {
        (t, s)
        for s, ts in moudawim_fixed.items()
        for t in ts
    }

    # ── Variables ─────────────────────────────────────────────────────────────

    x = {
        (t, r, s): model.NewBoolVar(f"x_{t}_{r}_{s}")
        for t in teachers
        for (r, s) in active_pairs
    }

    b = {
        (t, s): model.NewBoolVar(f"b_{t}_{s}")
        for t in teachers
        for s in sessions
    }

    # ── 1. Exactly 2 مراقب per (room, session) ───────────────────────────────
    for (r, s) in active_pairs:
        model.Add(sum(x[(t, r, s)] for t in teachers) == 2)

    # ── 2a. At most 1 room per teacher per session ───────────────────────────
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms_in_session[s]]
            if in_rooms:
                model.Add(sum(in_rooms) <= 1)

    # ── 2b. Each teacher supervises each room at most once across all sessions
    for t in teachers:
        for r in rooms:
            sessions_for_room = [s for s in sessions if (r, s) in active_pairs]
            if len(sessions_for_room) > 1:
                model.Add(sum(x[(t, r, s)] for s in sessions_for_room) <= 1)

    # ── 3. Exclusive roles + مداوم blocking ──────────────────────────────────
    for t in teachers:
        for s in sessions:
            in_rooms = [x[(t, r, s)] for r in rooms_in_session[s]]
            if (t, s) in moudawim_set:
                # مداوم: blocked from supervision AND backup this session
                if in_rooms:
                    model.Add(sum(in_rooms) == 0)
                model.Add(b[(t, s)] == 0)
            else:
                if in_rooms:
                    model.Add(sum(in_rooms) + b[(t, s)] <= 1)
                else:
                    model.Add(b[(t, s)] <= 1)

    # ── 4. Proportional backup distribution ──────────────────────────────────
    N               = len(teachers)
    total_rooms_all = sum(len(v) for v in rooms_in_session.values())

    if total_rooms_all == 0:
        for s in sessions:
            model.Add(sum(b[(t, s)] for t in teachers) == 0)
    else:
        raw    = {s: N * len(rooms_in_session[s]) / total_rooms_all for s in sessions}
        floors = {s: int(math.floor(raw[s])) for s in sessions}
        ceils  = {s: int(math.ceil(raw[s]))  for s in sessions}

        # Adjust so Σ floors == N
        remainder = N - sum(floors.values())
        for s in sorted(sessions, key=lambda s: -(raw[s] - floors[s]))[:remainder]:
            floors[s] += 1

        # Clamp to free pool (teachers not blocked as مداوم in this session)
        for s in sessions:
            free_pool = len([t for t in teachers if (t, s) not in moudawim_set])
            floors[s] = min(floors[s], free_pool)
            ceils[s]  = max(floors[s], min(ceils[s], free_pool))

        for s in sessions:
            b_sum = sum(b[(t, s)] for t in teachers)
            if rooms_in_session[s]:
                model.Add(b_sum >= floors[s])
                model.Add(b_sum <= ceils[s])
            else:
                model.Add(b_sum == 0)

    # ── 5. Each teacher is احتياطي exactly once ───────────────────────────────
    for t in teachers:
        model.Add(sum(b[(t, s)] for s in sessions) == 1)

    # ── 6. Every teacher must have ≥1 duty ───────────────────────────────────
    for t in teachers:
        is_moudawim = any((t, s) in moudawim_set for s in sessions)
        if not is_moudawim:
            sup = sum(x[(t, r, s)] for (r, s) in active_pairs)
            bup = sum(b[(t, s)] for s in sessions)
            model.Add(sup + bup >= 1)

    # ── 7. Balanced load ──────────────────────────────────────────────────────
    # Fixed مداوم count as load (Python int, not CP-SAT variable)
    mou_total_slots = sum(len(ts) for ts in moudawim_fixed.values())
    total_slots     = 2 * len(active_pairs) + N + mou_total_slots
    avg             = (total_slots + N - 1) // N

    load_vars: Dict[str, Tuple] = {}
    sq_diffs = []

    for t in teachers:
        sup_load  = sum(x[(t, r, s)] for (r, s) in active_pairs)
        bup_load  = sum(b[(t, s)] for s in sessions)
        mou_fixed = sum(1 for s in sessions if (t, s) in moudawim_set)

        # Solver only balances the non-fixed portion
        effective_avg = max(0, avg - mou_fixed)
        solver_load   = sup_load + bup_load

        model.Add(solver_load <= effective_avg + 1)
        model.Add(solver_load >= max(0, effective_avg - 1))

        diff = model.NewIntVar(-1, 1, f"diff_{t}")
        sq   = model.NewIntVar(0,  1, f"sq_{t}")
        model.Add(diff == solver_load - effective_avg)
        model.AddMultiplicationEquality(sq, diff, diff)
        sq_diffs.append(sq)

        load_vars[t] = (solver_load, mou_fixed)

    model.Minimize(sum(sq_diffs))

    # ── Solve ─────────────────────────────────────────────────────────────────
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(
            "لم يتمكن السولفر من إيجاد حل — تحقق من عدد الأساتذة والمداومين والأزواج النشطة."
        )

    # ── Extract results ───────────────────────────────────────────────────────
    supervisors: Dict[Tuple[str, str], List[str]] = {}
    for (t, r, s), var in x.items():
        if solver.Value(var):
            supervisors.setdefault((r, s), []).append(t)

    backups: Dict[str, List[str]] = {}
    for (t, s), var in b.items():
        if solver.Value(var):
            backups.setdefault(s, []).append(t)

    moudawim_out = {s: list(ts) for s, ts in moudawim_fixed.items() if ts}

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
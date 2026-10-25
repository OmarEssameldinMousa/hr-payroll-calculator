# Build spec — HR payroll calculation engine (Streamlit prototype)

Hand this file to Claude Code as the brief. The code in this repo already
implements it; use the spec to extend or rebuild.

## Goal
Validate the payroll **calculation engine** in Python before porting it to the
Express/TypeScript HR system. One Streamlit page, five steps:

1. Create a **payroll group** and its **global config**.
2. Create / edit **payroll policies** (seeded: C04 office + bonus + overtime, C05 office).
3. Upload the device **.xls** export (one employee) and parse only:
   date, clock-in datetime, clock-out datetime, day status.
4. Enter **employee details** not in the xls, and apply **HR day-by-day overrides**.
5. Calculate and **export** an .xlsx identical in layout to `hr_report_V1`.

## Architecture (keep the engine pure)
```
hr_engine/
  config.py   # group/global/payroll defaults, seed, JSON persistence (data/config.json)
  parser.py   # .xls → RawDay[]        (I/O only, no rules)
  engine.py   # run_period(days, employee, global, payroll) → PeriodResult   (pure)
  report.py   # PeriodResult → xlsx bytes, fills assets/hr_report_template.xlsx
app.py        # Streamlit UI only — no business rules
tests/        # pytest per rule + end-to-end on the sample xls
```
- Day amounts are `fractions.Fraction` (never float). Time is integer minutes.
- Every deduction/bonus is a `LedgerEntry(date, rule, amount, reason)`. KPIs are aggregates of the ledger.
- `run_period` returns a `carry_over` dict that becomes the next period's inputs.

## Parser rules (device export, Sheet1)
| col | meaning |
|---|---|
| 20 | date (Excel serial) |
| 16 | clock-in time |
| 14 | clock-out time |
| 18 | status text |
| 25 / 31 | name / code (first data row) |

- `00:00:00` or empty = no punch.
- Clock-out ≤ clock-in → clock-out is the next day (ignore device date columns 13/15, they are unreliable).
- Status: `""`→attended, `غياب`→absent_unexcused, `غياب - بإذن`→absent_excused, `عطلة`→off_day, `يوم إضافي`→additional_day.
- Period from the header text `من dd/mm/yyyy الى dd/mm/yyyy`.

## Domain model
**PayrollGroup** `{name, global, payrolls[]}`

**Global config** (per group): `grace_minutes`, `max_permissions_per_period` (3),
`exception_overtime_threshold_minutes` (360), `exception_minutes_per_hour` (15),
`exception_exceeded_deduct` (1/3), `dwr_next_day_cutoff` (12:00).

**Payroll**: `code, name, overtime_enabled, bonus_enabled, bonus_hour_multiplier (2),
lateness_tiers [{after_minutes:10, deduct:1/4},{30, 1/3}], lateness_escalation_threshold (6),
lateness_escalation_multiplier (2), lateness_streak_multiplier (3),
permission_schedule [0,1/3,2/3], permission_reward_schedule [0,0,2/3],
permission_max_minutes (120), permission_exceeded_deduct (1/3),
absence_unexcused_deduct (2), absence_excused_deduct (0),
weekly_off_allowed_per_week (1), weekly_off_extra_deduct (2), weekly_off_min_gap_workdays (3, warning only),
dwr_late_after_minutes (60), dwr_late_deduct (1/4), dwr_next_day_deduct (1/3), dwr_missing_deduct (1/3)`.

**Day** (engine input): `date, status, clock_in, clock_out, permission (none|late_arrival|early_leave),
dwr_state (on_time|submitted|not_submitted|pending), dwr_time, bonus_days, deduction_days, adjustment_reason, overridden{fields}`.

**Employee**: `official_in, official_out, official_off_weekday, permissions_used_last_period,
lates_last_period, prev_last_day_overtime_minutes, carried_dwr (none|late|next_day|missing),
hr_name, bonus_hours, bonus_hours_reason`.

## Business rules (confirmed with HR)
1. **Lateness**: late = clock_in − (official_in + grace). Highest tier passed applies: >10m → 1/4, >30m → 1/3.
2. **Morning exception from overtime**: overtime on a present day earns
   `floor((overtime − threshold)/60) × minutes_per_hour` of allowed lateness for the **next workday**.
   If late ≤ exception → no deduction (counted as morning permission). If late > exception → 1/3.
   An absence breaks the chain. Day 1 uses `prev_last_day_overtime_minutes`.
3. **Permissions**: max 3 per payroll period. Cost by order: `[0, 1/3, 2/3]`; if 0 were used
   last period → `[0, 0, 2/3]`. A 4th is ignored → normal lateness rules. A permission day is
   **not** counted as a late arrival. Lateness above `permission_max_minutes` → at least 1/3.
4. **Escalation**: if lates in the period ≥ 6 → all lateness deductions ×2; ×3 if the
   previous period was also ≥ 6.
5. **Absence**: excused → 0. Unexcused → −2 days (configurable).
6. **Weekly off**: 1 per 7-day block from the period start; each extra → −2 days.
   Fewer than 3 workdays between offs → warning.
7. **DWR**: not submitted → 1/3; submitted after next day 12:00 → 1/3; submitted > 60 min after
   clock-out → 1/4. Last day may be `pending` → carried into the next period.
8. **Hours**: official = daily official × workdays (attended + absences). Overtime = worked − official
   (all worked time on an additional day), only if `overtime_enabled`.
   Balance = regular + overtime − official. Bonus hours are ×2, then added; show raw → ×2 → balance after.
9. **Manual**: per-day bonus/deduction (days) with a reason → actions sheet.
10. The payroll period is ~26 days and depends on the employee; take it from the file.

## Export (must match hr_report_V1)
Fill a copy of `assets/hr_report_template.xlsx`. Keep header rows 1–2, delete the sample rows,
clone styles from row 3 (data) and row 5 (totals).

Sheet `attendance`, one row per day:
| col | meaning |
|---|---|
| A | total day deduction |
| B | DWR deduction (خصم_3) |
| C | DWR delay hrs |
| D | exception earned `"+30min"` |
| E | overtime |
| F | lateness deduction (خصم_2) |
| G | morning lateness hrs |
| H | permission deduction (خصم_1) |
| I | `late arrival N` |
| J | worked hrs |
| K | done by (HR name if overridden, yellow) |
| L | clock-out `05:32am, 13/aug` |
| M | clock-in `09:16am,` |
| N | date `d-mmm` |
| O | day `wed` |

- Totals row: A, B, E, F, H, J.
- Formats: `-1/4 day`, `-1,1/4 days`, `_` for zero, `0 day` in totals, `HH:MM hrs`.
- Sheet `actions`: reason | value | type (خصم / مكافأة) | date | day.
- Optional 3rd `summary` sheet with the PRD KPIs (off by default).

## Acceptance
- `pytest` passes: tiers, escalation ×2/×3, permission schedules, 4th permission, exception chain,
  C05 no overtime, absence/off, DWR tiers, hours balance, end-to-end sample.
- Real HR months (golden cases) reproduce HR's hand-calculated totals exactly.

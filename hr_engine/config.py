"""The single payroll group and its payroll policy (fixed, not user-editable).

The only per-run inputs are the employee's official check-in, check-out and
weekly day off (see engine.Employee). Day amounts are stored as strings like
"1/4" and turned into exact Fractions by the engine. Never use floats for day
units.
"""
from __future__ import annotations

from fractions import Fraction


def F(x) -> Fraction:
    return Fraction(str(x).strip()) if str(x).strip() else Fraction(0)


PAYROLL = {
    "code": "C04",
    "name": "Office",
    # late arrival on a normal day: the highest tier passed applies
    "lateness_tiers": [
        {"after_minutes": 10, "deduct": "1/4"},
        {"after_minutes": 30, "deduct": "1/3"},
    ],
    # late arrival on a Friday
    "friday_lateness_tiers": [
        {"after_minutes": 30, "deduct": "1/3"},
    ],
    # overtime above this earns a flat exception on the next attended day,
    # no matter how much overtime was done
    "exception_overtime_after_minutes": 60,
    "exception_minutes": 30,
    "exception_exceeded_deduct": "1/3",
    "absence_unexcused_deduct": "2",
}

GROUP = {"name": "Office", "payrolls": [PAYROLL]}

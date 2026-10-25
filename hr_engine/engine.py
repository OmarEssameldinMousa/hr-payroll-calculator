"""Calculation engine — pure functions, no I/O, no Streamlit.

run_period(days, employee, payroll) -> PeriodResult

Rules evaluated (all amounts from the payroll dict, see config.py):
  * late arrival:      <=10m nothing, >10m..30m -1/4, >30m -1/3
  * friday:            <=30m nothing, >30m -1/3
  * overtime exception: overtime (time after official clock-out) > 60m earns a
                       flat 30m lateness allowance on the next day; late beyond
                       it -> -1/3. A day off or absence in between cancels it.
  * unexcused absence: -2 days
  * official day off:  weekly vacation — no deductions, no overtime

All day amounts are exact Fractions (negative = deduction). Time is in
integer minutes. Every deduction is also written to a ledger with a rule id
and reason, so any number in the report can be explained.
"""
from __future__ import annotations  # allows "int | None" style type hints on older Pythons

from dataclasses import dataclass, field  # dataclass = simple record classes; field = per-instance defaults
from datetime import date, datetime, time, timedelta  # date/time types used for punches and schedules
from fractions import Fraction  # exact day amounts like 1/3 (floats would give rounding errors)

from .config import F  # helper that turns "1/4" (string) into Fraction(1, 4)

# every status the parser can produce from the device export
STATUSES = ["attended", "absent_unexcused", "absent_excused", "off_day", "additional_day"]
# statuses where the employee was expected at work (they count toward official hours)
WORKDAY_STATUSES = {"attended", "absent_unexcused", "absent_excused"}
# statuses where the employee actually came in (they have clock-in/out punches)
PRESENT_STATUSES = {"attended", "additional_day"}
# Python's weekday() numbering: Mon=0 … Thu=3, Fri=4 … Sun=6
FRIDAY = 4


# ---------------------------------------------------------------- inputs
@dataclass  # auto-generates __init__ etc. from the fields below
class Day:  # one row of the attendance export = one calendar day
    date: date  # the calendar day
    status: str = "attended"  # one of STATUSES
    clock_in: datetime | None = None  # check-in punch, None if missing
    clock_out: datetime | None = None  # check-out punch (may be the next calendar day), None if missing
    raw_status: str = ""        # status text exactly as in the device export (goes to the ملحوظة column)
    edited: set[str] = field(default_factory=set)  # fields HR fixed in the UI: "status", "clock_in", "clock_out"

    @property  # read like an attribute: day.day_name
    def day_name(self) -> str:  # short lowercase weekday name for the sheet
        return self.date.strftime("%a").lower()  # e.g. "sat", "sun"


@dataclass  # auto-generates __init__ etc.
class Employee:  # the employee's schedule — the only user-configurable inputs
    code: str  # employee code from the export
    name: str  # employee name from the export (also used for the file name)
    official_in: time = time(9, 0)  # official check-in time
    official_out: time = time(17, 30)  # official check-out time
    official_off_weekday: int = 3              # weekly vacation day: 0=Mon … 3=Thu … 6=Sun


# ---------------------------------------------------------------- outputs
@dataclass  # auto-generates __init__ etc.
class LedgerEntry:  # one line explaining one deduction
    date: date | None  # the day the deduction belongs to
    rule: str           # LATENESS | FRIDAY_LATENESS | EXCEPTION_EXCEEDED | ABSENCE
    amount: Fraction    # negative = deduction
    reason: str  # human-readable explanation, e.g. "late 35m"


@dataclass  # auto-generates __init__ etc.
class DayResult:  # everything calculated for one day (one row of the sheet)
    day: Day  # the input day this result is for
    official_minutes: int = 0  # minutes the employee was supposed to work this day
    worked_minutes: int = 0  # clock-out minus clock-in
    overtime_minutes: int = 0  # minutes stayed after the official check-out
    late_minutes: int = 0  # minutes arrived after the official check-in
    exception_available: int = 0  # exception minutes carried in from yesterday (0 or 30)
    exception_earned: int = 0  # exception minutes earned today for tomorrow (0 or 30)
    exception_used: bool = False  # True if today's lateness was forgiven by the exception
    is_official_off: bool = False  # True if today is the employee's weekly vacation day
    lateness_deduct: Fraction = Fraction(0)  # lateness deduction in days (negative)
    absence_deduct: Fraction = Fraction(0)  # absence deduction in days (negative)
    counted_late: bool = False  # True if a lateness deduction was applied today
    notes: list[str] = field(default_factory=list)  # free-text notes, e.g. "missing punch"

    @property  # read like an attribute: result.total_deduct
    def total_deduct(self) -> Fraction:  # the day's total deduction (column A of the sheet)
        return self.lateness_deduct + self.absence_deduct  # both are negative or zero


@dataclass  # auto-generates __init__ etc.
class PeriodResult:  # everything calculated for the whole period
    days: list[DayResult]  # one result per day, sorted by date
    ledger: list[LedgerEntry]  # every deduction with its reason
    kpis: dict  # period totals and counts for the summary
    warnings: list[str]  # data problems HR should look at (e.g. missing punches)


# ---------------------------------------------------------------- helpers
def _mins(a: datetime, b: datetime) -> int:  # whole minutes from a to b (negative if b is before a)
    return int((b - a).total_seconds() // 60)  # subtract, convert to seconds, floor-divide to minutes


def _tier_deduct(late: int, tiers: list[dict]) -> Fraction:  # deduction for `late` minutes under a tier table
    best = Fraction(0)  # default: no deduction
    for t in sorted(tiers, key=lambda t: int(t["after_minutes"])):  # walk tiers from smallest threshold up
        if late > int(t["after_minutes"]):  # strictly past this tier's threshold?
            best = F(t["deduct"])  # yes → this tier applies (a higher tier later overrides it)
    return best  # deduction of the highest tier passed, as a positive Fraction


# ---------------------------------------------------------------- engine
def run_period(days: list[Day], emp: Employee, p: dict) -> PeriodResult:  # p = payroll rules from config.py
    # sort days by date, so that the exception carry logic works correctly
    days = sorted(days, key=lambda d: d.date)
    # initialize outputs
    ledger: list[LedgerEntry] = []  # every deduction with its reason
    warnings: list[str] = []  # data problems found while calculating
    results: list[DayResult] = []  # one DayResult per day
    # compute official daily minutes (official check-out minus official check-in)
    official_daily = _mins(datetime.combine(date.min, emp.official_in),
                           datetime.combine(date.min, emp.official_out))
    # handle overnight shifts (official_out < official_in) by adding a full day
    if official_daily <= 0:
        official_daily += 24 * 60
    # overtime must be strictly more than this many minutes to earn the exception (60)
    ex_after = int(p["exception_overtime_after_minutes"])
    # how many minutes of lateness the exception forgives next day (30, flat)
    ex_minutes = int(p["exception_minutes"])
    # exception minutes carried from the previous day (0 = no exception today)
    exception_carry = 0
    # iterate over days and compute results
    for d in days:
        # initialize day result
        r = DayResult(day=d)
        # official day off: today's weekday is the employee's weekly vacation day
        r.is_official_off = d.date.weekday() == emp.official_off_weekday
        # workday: the employee was expected at work (and it isn't their vacation day)
        workday = d.status in WORKDAY_STATUSES and not r.is_official_off
        if workday:  # only workdays count toward official hours
            r.official_minutes = official_daily  # expected minutes for this day

        # ---- hours / overtime
        if d.status in PRESENT_STATUSES and d.clock_in and d.clock_out:  # came in and both punches exist
            r.worked_minutes = max(0, _mins(d.clock_in, d.clock_out))  # time between punches, never negative
            if not r.is_official_off:   # overtime = time stayed after official clock-out (none on vacation day)
                end = datetime.combine(d.date, emp.official_out)  # official check-out on this date
                if emp.official_out <= emp.official_in:  # overnight shift: check-out is on the next day
                    end += timedelta(days=1)  # move official check-out to tomorrow
                r.overtime_minutes = max(0, _mins(end, d.clock_out))  # minutes after official check-out, or 0
        elif d.status in PRESENT_STATUSES:  # came in but a punch is missing
            r.notes.append("missing punch")  # note it on the day
            warnings.append(f"{d.date}: missing clock-in/out, hours not computed")  # and warn HR

        # ---- lateness: overtime exception > friday rule > normal tiers
        if d.status == "attended" and d.clock_in and not r.is_official_off:  # normal attended day with a check-in
            start = datetime.combine(d.date, emp.official_in)  # official check-in on this date
            r.late_minutes = max(0, _mins(start, d.clock_in))  # minutes after official check-in, or 0
            r.exception_available = exception_carry  # record whether yesterday's overtime gave an exception
            rule, amt = "LATENESS", Fraction(0)  # default: normal lateness rule, no deduction yet
            if exception_carry:  # 1) yesterday earned the 30-minute exception
                if r.late_minutes > exception_carry:  # late more than 30m → exception exceeded
                    rule, amt = "EXCEPTION_EXCEEDED", F(p["exception_exceeded_deduct"])  # deduct 1/3
                elif r.late_minutes > 0:  # late but within 30m → forgiven
                    r.exception_used = True  # mark the exception as used (no deduction)
            elif d.date.weekday() == FRIDAY:  # 2) Friday: only >30m counts
                rule, amt = "FRIDAY_LATENESS", _tier_deduct(r.late_minutes, p["friday_lateness_tiers"])
            else:  # 3) any other day: normal tiers (>10m 1/4, >30m 1/3)
                amt = _tier_deduct(r.late_minutes, p["lateness_tiers"])
            if amt:  # a deduction applies
                r.lateness_deduct = -amt  # store it as a negative amount
                r.counted_late = True  # count this day as a late day
                extra = f" > exception {exception_carry}m" if rule == "EXCEPTION_EXCEEDED" else ""  # reason suffix
                ledger.append(LedgerEntry(d.date, rule, -amt, f"late {r.late_minutes}m{extra}"))  # explain it

        # ---- absence
        if d.status == "absent_unexcused" and not r.is_official_off:  # absent without permission on a workday
            amt = F(p["absence_unexcused_deduct"])  # 2 days
            r.absence_deduct = -amt  # store it as a negative amount
            ledger.append(LedgerEntry(d.date, "ABSENCE", -amt, "absence without permission"))  # explain it

        # ---- exception for tomorrow
        # exception earned today feeds the next day only (flat, not per hour);
        # any day off or absence in between cancels it
        if d.status == "attended" and not r.is_official_off:   # additional days never earn it
            r.exception_earned = ex_minutes if r.overtime_minutes > ex_after else 0  # 30 if overtime > 60m
            exception_carry = r.exception_earned  # hand it to tomorrow (0 clears any old one)
        else:  # off day, official day off, absence or additional day
            exception_carry = 0  # cancel any exception
        results.append(r)  # save today's result

    # ---- aggregate
    def s(attr):  # sum a Fraction field (e.g. "lateness_deduct") over all days
        return sum((getattr(r, attr) for r in results), Fraction(0))  # start at Fraction(0) to keep it exact

    official = sum(r.official_minutes for r in results)  # total expected minutes
    worked = sum(r.worked_minutes for r in results)  # total minutes between punches
    overtime = sum(r.overtime_minutes for r in results)  # total minutes after official check-out
    st = lambda x: sum(1 for r in results if r.day.status == x)  # count days with status x
    count = lambda rule, amt: sum(1 for e in ledger if e.rule == rule and e.amount == amt)  # count ledger hits

    kpis = {  # period totals shown in the KPI tab and the optional summary sheet
        "official_minutes": official,  # expected minutes in the period
        "worked_regular_minutes": worked - overtime,  # worked minutes excluding overtime
        "worked_total_minutes": worked,  # all worked minutes
        "overtime_minutes": overtime,  # all overtime minutes
        "balance_minutes": worked - official,  # surplus (+) or shortfall (−) against official hours
        "attendance_days": st("attended") + st("additional_day"),  # days the employee came in
        "additional_days": st("additional_day"),  # days marked يوم إضافي
        "late_count": sum(r.counted_late for r in results),  # days with any lateness deduction
        "late_quarter_count": count("LATENESS", Fraction(-1, 4)),  # normal days late >10m..30m
        "late_third_count": count("LATENESS", Fraction(-1, 3)),  # normal days late >30m
        "friday_late_count": sum(1 for e in ledger if e.rule == "FRIDAY_LATENESS"),  # Fridays late >30m
        "exception_exceeded_count": sum(1 for e in ledger if e.rule == "EXCEPTION_EXCEEDED"),  # late >30m on exception days
        "morning_exceptions_earned": sum(1 for r in results if r.exception_earned),  # days that earned an exception
        "morning_exceptions_used": sum(r.exception_used for r in results),  # days lateness was forgiven
        "absence_unexcused": sum(1 for r in results if r.absence_deduct),  # deducted absences (excludes vacation day)
        "absence_excused": st("absent_excused"),  # absences with permission (no deduction)
        "off_days": st("off_day"),  # days marked عطلة
        "deduct_lateness": s("lateness_deduct"),  # total lateness deduction in days
        "deduct_absence": s("absence_deduct"),  # total absence deduction in days
        "deductions_total": s("lateness_deduct") + s("absence_deduct"),  # grand total deduction in days
    }
    return PeriodResult(results, ledger, kpis, warnings)  # hand everything to the UI and the report

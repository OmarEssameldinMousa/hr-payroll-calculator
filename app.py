"""HR payroll calculation engine — all-in-one Streamlit prototype.

Run:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import re
from collections import Counter
from datetime import datetime, time, timedelta

import pandas as pd
import streamlit as st

from hr_engine.config import GROUP, PAYROLL
from hr_engine.engine import PRESENT_STATUSES, STATUSES, Day, Employee, run_period
from hr_engine.parser import STATUS_MAP, parse_attendance_xls
from hr_engine.report import build_pdf, build_report, fmt_days, fmt_hm, summary_rows

st.set_page_config(page_title="HR payroll engine", layout="wide")

st.title("HR payroll calculation engine — prototype")
st.caption("1 Payroll group → 2 Upload attendance → 3 Employee schedule → 4 Fix data problems → 5 Calculate & export")


# =========================================================== 1. payroll group
st.header(f"1 · Payroll group: {GROUP['name']}")
p = PAYROLL
st.markdown(f"**Payroll {p['code']} — {p['name']}** (fixed rules)")
st.dataframe(pd.DataFrame([  # not st.table: it renders markdown, so "> 30 min" became a blockquote
    ("Late arrival", "≤ 10 min", "nothing"),
    ("Late arrival", "> 10 min and ≤ 30 min", "-1/4 day"),
    ("Late arrival", "> 30 min", "-1/3 day"),
    ("Late arrival on Friday", "≤ 30 min", "nothing"),
    ("Late arrival on Friday", "> 30 min", "-1/3 day"),
    ("Overtime exception", "stayed > 1:00 after official check-out → next day may be up to 30 min late "
                           "(always 30 min, however much overtime)", "nothing"),
    ("Overtime exception", "late > 30 min on the exception day", "-1/3 day"),
    ("Overtime exception", "a day off or absence before the next attended day cancels it", "—"),
    ("Absence without permission", "غياب", "-2 days"),
    ("Official day off", "weekly vacation — no deductions, no overtime", "nothing"),
], columns=["Rule", "Condition", "Deduction"]), hide_index=True, width="stretch")

# =========================================================== 2. upload
st.header("2 · Upload attendance export (.xls)")
up = st.file_uploader("Device export (one employee)", type=["xls"])
if up is None:
    st.info("Upload the .xls to continue.")
    st.stop()

parsed = parse_attendance_xls(up.getvalue())
st.write(f"**{parsed.employee_name}** (code {parsed.employee_code}) · period "
         f"{parsed.period_start} → {parsed.period_end} · {len(parsed.days)} days")
with st.expander("Parsed raw days", expanded=False):
    st.dataframe(pd.DataFrame([{
        "date": d.date, "day": d.date.strftime("%a").lower(), "status": d.status,
        "raw status": d.raw_status_text, "clock_in": d.clock_in, "clock_out": d.clock_out,
    } for d in parsed.days]), width="stretch", hide_index=True)

# =========================================================== 3. employee schedule
st.header("3 · Employee schedule")
wk = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
# default day off = the weekday the device marks as عطلة most often (Thu if none)
off_counts = Counter(d.date.weekday() for d in parsed.days if d.status == "off_day")
detected_wd = off_counts.most_common(1)[0][0] if off_counts else 3
a, b, c = st.columns(3)
off_in = a.time_input("Official check-in", time(9, 0))
off_out = b.time_input("Official check-out", time(17, 30))
off_wd = c.selectbox("Official day off", range(7), index=detected_wd, format_func=lambda i: wk[i],
                     key=f"offwd_{up.name}",
                     help=f"Detected from the file: {wk[detected_wd]}" if off_counts else "No عطلة days in the file")
worked_on_off = [d.date for d in parsed.days if d.status == "attended" and d.date.weekday() == off_wd]
if worked_on_off:
    st.warning(f"The employee attended a normal day on {len(worked_on_off)} {wk[off_wd]}(s) "
               f"({', '.join(map(str, worked_on_off))}). Those days get no overtime and no lateness check. "
               "Check that the official day off is right.")

days = [Day(d.date, d.status, d.clock_in, d.clock_out, d.raw_status_text) for d in parsed.days]


def problem(d: Day) -> str:
    """Why a day needs HR attention ('' if it's fine)."""
    if d.raw_status not in STATUS_MAP:
        return f"unknown status '{d.raw_status}'"
    if d.status in PRESENT_STATUSES and (d.clock_in is None or d.clock_out is None):
        return "missing " + " and ".join(f for f, v in (("clock-in", d.clock_in), ("clock-out", d.clock_out)) if v is None)
    return ""


def to_time(v):
    if v is None or (not isinstance(v, time) and pd.isna(v)):
        return None
    return v if isinstance(v, time) else pd.to_datetime(str(v)).time()


# =========================================================== 4. fix data problems
st.header("4 · Fix data problems")
show_all = st.toggle("Show all days (not only the ones with problems)", key=f"showall_{up.name}")
rows = [d for d in days if show_all or problem(d)]
if not rows:
    st.success("No missing punches or unknown statuses.")
else:
    st.caption("Fill in the missing times or change the status. A clock-out earlier than the clock-in counts as "
               "the next day. Fixed cells are highlighted yellow in the exported sheet.")
    edited = st.data_editor(pd.DataFrame([{
        "date": d.date, "day": d.day_name, "problem": problem(d), "ملحوظة": d.raw_status, "status": d.status,
        "clock_in": d.clock_in.time() if d.clock_in else None,
        "clock_out": d.clock_out.time() if d.clock_out else None,
    } for d in rows]), key=f"fix_{up.name}_{show_all}", width="stretch", hide_index=True, num_rows="fixed",
        column_config={
            "date": st.column_config.DateColumn("Date", disabled=True),
            "day": st.column_config.TextColumn("Day", disabled=True),
            "problem": st.column_config.TextColumn("Problem (from the file)", disabled=True),
            "ملحوظة": st.column_config.TextColumn("ملحوظة", disabled=True),
            "status": st.column_config.SelectboxColumn("Status", options=STATUSES, required=True),
            "clock_in": st.column_config.TimeColumn("Clock in", format="HH:mm"),
            "clock_out": st.column_config.TimeColumn("Clock out", format="HH:mm"),
        })
    for d, (_, e) in zip(rows, edited.iterrows()):
        if e.status != d.status:
            d.status = e.status
            d.edited.add("status")
        t_in, t_out = to_time(e.clock_in), to_time(e.clock_out)
        in_changed = t_in != (d.clock_in.time() if d.clock_in else None)
        out_changed = t_out != (d.clock_out.time() if d.clock_out else None)
        if in_changed:
            d.clock_in = datetime.combine(d.date, t_in) if t_in else None
            d.edited.add("clock_in")
        if in_changed or out_changed:   # rebuild clock-out so the overnight rule follows the new clock-in
            d.clock_out = datetime.combine(d.date, t_out) if t_out else None
            if d.clock_in and d.clock_out and d.clock_out <= d.clock_in:   # overnight shift
                d.clock_out += timedelta(days=1)
        if out_changed:
            d.edited.add("clock_out")
    left = [d for d in days if problem(d)]
    if left:
        st.warning(f"{len(left)} day(s) still need fixing: " + ", ".join(f"{d.date} ({problem(d)})" for d in left))
    else:
        st.success("All problems fixed.")

emp = Employee(code=parsed.employee_code, name=parsed.employee_name, official_in=off_in,
               official_out=off_out, official_off_weekday=off_wd)

# =========================================================== 5. calculate + export
st.header("5 · Results & export")
res = run_period(days, emp, p)
k = res.kpis

m = st.columns(4)
m[0].metric("Worked / official", f"{fmt_hm(k['worked_regular_minutes'], '0')}",
            f"of {fmt_hm(k['official_minutes'], '0')}", delta_color="off")
m[1].metric("Overtime", fmt_hm(k["overtime_minutes"], "00:00 hrs"))
m[2].metric("Hour balance", fmt_hm(k["balance_minutes"], "00:00 hrs"))
m[3].metric("Total deductions", fmt_days(k["deductions_total"], "0 day"))
if res.warnings:
    with st.expander(f"⚠ {len(res.warnings)} engine warnings", expanded=False):
        for w in res.warnings:
            st.warning(w)

t1, t2, t3 = st.tabs(["Per day", "KPIs", "Ledger (why)"])
with t1:
    st.dataframe(pd.DataFrame([{
        "date": r.day.date, "day": r.day.day_name, "status": r.day.status,
        "ملحوظة": r.day.raw_status, "official off": "✓" if r.is_official_off else "",
        "worked": fmt_hm(r.worked_minutes), "overtime": fmt_hm(r.overtime_minutes),
        "late": fmt_hm(r.late_minutes), "exception avail": f"{r.exception_available}min" if r.exception_available else "",
        "exception earned": f"+{r.exception_earned}min" if r.exception_earned else "",
        "lateness": fmt_days(r.lateness_deduct), "absence": fmt_days(r.absence_deduct),
        "total": fmt_days(r.total_deduct), "notes": "; ".join(r.notes),
    } for r in res.days]), width="stretch", hide_index=True, height=38 + 35 * len(res.days))
with t2:
    st.dataframe(pd.DataFrame(summary_rows(res), columns=["Section", "KPI", "Value"]),
                 width="stretch", hide_index=True)
with t3:
    st.dataframe(pd.DataFrame([{"date": e.date, "rule": e.rule, "amount": fmt_days(e.amount), "reason": e.reason}
                               for e in res.ledger]), width="stretch", hide_index=True)

inc = st.checkbox("Add a KPI summary sheet (extra 3rd sheet, not in hr_report_V1)", value=False)
safe_name = re.sub(r'[\\/:*?"<>|]+', "_", emp.name).strip() or emp.code
base_name = f"{safe_name}_{parsed.period_start}_{parsed.period_end}"
xlsx = build_report(res, emp, inc)
a, b = st.columns(2)
a.download_button("⬇ Export HR report (.xlsx)", xlsx, f"{base_name}.xlsx",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
# PDF = the same Excel report converted by LibreOffice; slow, so only on request.
# Any change to the data/inputs changes the xlsx bytes, which invalidates the prepared PDF.
key = hashlib.sha1(xlsx).hexdigest()
if b.button("Prepare PDF"):
    with st.spinner("Converting the Excel report to PDF…"):
        try:
            st.session_state.pdf = (key, build_pdf(res, emp, inc))
        except Exception as e:
            st.error(str(e))
if st.session_state.get("pdf", (None,))[0] == key:
    b.download_button("⬇ Export HR report (.pdf)", st.session_state.pdf[1], f"{base_name}.pdf",
                      "application/pdf", type="primary")

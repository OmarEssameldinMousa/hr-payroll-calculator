"""Parse the device's Crystal-Reports .xls export into RawDay rows.

Only extracts: date, clock-in datetime, clock-out datetime, day status
(+ employee code/name and period as metadata).

Layout (Sheet1, fixed columns): 20 = date serial, 14 = out time,
16 = in time, 18 = status text, 25 = name, 31 = code.
Clock-out earlier than clock-in => next day (the device's date columns
13/15 are unreliable and ignored). "00:00:00" means no punch.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta

import xlrd

COL_DATE, COL_OUT, COL_IN, COL_STATUS, COL_NAME, COL_CODE = 20, 14, 16, 18, 25, 31

STATUS_MAP = {
    "": "attended",
    "غياب": "absent_unexcused",
    "غياب - بإذن": "absent_excused",
    "عطلة": "off_day",
    "يوم إضافي": "additional_day",
}


@dataclass
class RawDay:
    date: date
    status: str
    clock_in: datetime | None
    clock_out: datetime | None
    raw_status_text: str = ""


@dataclass
class ParsedSheet:
    employee_code: str
    employee_name: str
    period_start: date | None
    period_end: date | None
    days: list[RawDay]
    warnings: list[str]


def _serial_to_date(v: float, datemode: int) -> date:
    return xlrd.xldate_as_datetime(v, datemode).date()


def _parse_time(v) -> time | None:
    s = str(v).strip()
    if not s:
        return None
    if isinstance(v, float):
        if v == 0:
            return None
        secs = round(v * 86400) % 86400
        return time(secs // 3600, (secs % 3600) // 60)
    m = re.match(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", s)
    if not m:
        return None
    t = time(int(m.group(1)) % 24, int(m.group(2)))
    return None if t == time(0, 0) else t


def parse_attendance_xls(file_bytes: bytes) -> ParsedSheet:
    book = xlrd.open_workbook(file_contents=file_bytes)
    sh = book.sheet_by_index(0)
    warnings: list[str] = []

    def cell(r, c):
        return sh.cell_value(r, c) if c < sh.ncols else ""

    code = name = ""
    p_start = p_end = None
    for r in range(min(sh.nrows, 12)):
        for c in range(sh.ncols):
            s = str(cell(r, c))
            m = re.search(r"من\s*(\d{1,2}/\d{1,2}/\d{4})\s*الى\s*(\d{1,2}/\d{1,2}/\d{4})", s)
            if m:
                p_start = datetime.strptime(m.group(1), "%d/%m/%Y").date()
                p_end = datetime.strptime(m.group(2), "%d/%m/%Y").date()

    days: list[RawDay] = []
    for r in range(sh.nrows):
        dv = cell(r, COL_DATE)
        if not isinstance(dv, float) or dv < 40000:
            continue
        d = _serial_to_date(dv, book.datemode)
        if not code and str(cell(r, COL_CODE)).strip():
            cv = cell(r, COL_CODE)
            code = str(int(cv)) if isinstance(cv, float) else str(cv).strip()
            name = str(cell(r, COL_NAME)).strip()
        status_txt = str(cell(r, COL_STATUS)).strip()
        status = STATUS_MAP.get(status_txt)
        if status is None:
            warnings.append(f"{d}: unknown status '{status_txt}', treated as attended")
            status = "attended"
        t_in, t_out = _parse_time(cell(r, COL_IN)), _parse_time(cell(r, COL_OUT))
        cin = datetime.combine(d, t_in) if t_in else None
        cout = None
        if t_out:
            cout = datetime.combine(d, t_out)
            if cin and cout <= cin:
                cout += timedelta(days=1)
        if status in ("attended", "additional_day") and (cin is None or cout is None):
            warnings.append(f"{d}: missing {'clock-in' if cin is None else ''}"
                            f"{' and ' if cin is None and cout is None else ''}"
                            f"{'clock-out' if cout is None else ''} — HR should override")
        days.append(RawDay(d, status, cin, cout, status_txt))

    days.sort(key=lambda x: x.date)
    if days and p_start is None:
        p_start, p_end = days[0].date, days[-1].date
    return ParsedSheet(code, name, p_start, p_end, days, warnings)

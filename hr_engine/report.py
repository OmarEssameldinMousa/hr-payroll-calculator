"""Export a PeriodResult into the exact hr_report_V1 layout.

Opens the bundled template (assets/hr_report_template.xlsx), keeps the
two header rows of each sheet, removes the sample rows, and writes data
rows by cloning the sample row styles (row 3 = data, row 5 = totals).
Column K (ملحوظة) holds the raw status text from the device export.
"""
from __future__ import annotations

import io
import shutil
import subprocess
import tempfile
from copy import copy
from datetime import datetime
from fractions import Fraction
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .engine import Employee, PeriodResult

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "hr_report_template.xlsx"
NOFILL = PatternFill(fill_type=None)
YELLOW = PatternFill("solid", fgColor="FFFFE599")  # template colour for cells HR fixed
EDITED_COL = {"clock_out": 12, "clock_in": 13, "status": 13}
STATUS_LABEL = {
    "absent_unexcused": "absent",
    "absent_excused": "absent (excused)",
    "off_day": "off day",
}


# ---------------------------------------------------------------- formatting
def fmt_days(x: Fraction, zero: str = "_") -> str:
    if not x:
        return zero
    sign = "-" if x < 0 else ""
    a = abs(x)
    whole, rem = a.numerator // a.denominator, a - a.numerator // a.denominator
    body = f"{whole},{rem}" if whole and rem else (f"{whole}" if whole else f"{rem}")
    return f"{sign}{body} {'day' if a <= 1 else 'days'}"


def fmt_hm(minutes: int | None, zero: str = "_") -> str:
    if not minutes:
        return zero
    sign = "-" if minutes < 0 else ""
    m = abs(int(minutes))
    return f"{sign}{m // 60:02d}:{m % 60:02d} hrs"


def fmt_clock(dt: datetime | None, work_date) -> str:
    if dt is None:
        return "_"
    s = dt.strftime("%I:%M") + dt.strftime("%p").lower() + ","
    if dt.date() != work_date:
        s += f" {dt.day}/{dt.strftime('%b').lower()}"
    return s


# ---------------------------------------------------------------- writer
def _clear_below(ws, first_row: int):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= first_row:
            ws.unmerge_cells(str(rng))
    if ws.max_row >= first_row:
        ws.delete_rows(first_row, ws.max_row - first_row + 1)


def _fit_to_page(ws, ncol: int, last_row: int) -> None:
    """Shrink columns to their content and stretch rows so one A4 landscape page is filled.

    Widths are in Excel character units (~7px each); heights in points. Approximate is
    fine: fit-to-page still guarantees everything lands on one page.
    """
    # one readable font everywhere: the template's Aptos Narrow isn't on Linux/LibreOffice,
    # which then substitutes something else
    for r in range(1, last_row + 1):
        for c in range(1, ncol + 1):
            f = copy(ws.cell(r, c).font)
            f.name = "Arial"
            ws.cell(r, c).font = f

    def text_w(cell, text: str) -> float:   # width in Excel char units
        big = (cell.font.sz or 11) > 11 or cell.font.b   # 12pt / bold text is wider
        return len(text) * (1.15 if big else 1)

    # Size each column by its data. Long headers wrap onto 2 lines instead of widening the
    # column, and so do rare long values like "09:00am, (additional)" (they break at " (").
    # Row-1 group headers merged across 2 columns don't size a single column.
    wide = {(m.min_row, m.min_col) for m in ws.merged_cells.ranges if m.max_col > m.min_col}
    widths, wrap_rows = [], set()
    for c in range(1, ncol + 1):
        need = 0.0
        for r in (1, 2):
            cell = ws.cell(r, c)
            if cell.value and (r, c) not in wide:
                words = str(cell.value).split()
                need = max(need, text_w(cell, max(words, key=len)), text_w(cell, str(cell.value)) / 2)
        for r in range(3, last_row + 1):
            cell = ws.cell(r, c)
            if cell.value is None:
                continue
            text = cell.value.strftime("%d-%b") if isinstance(cell.value, datetime) else str(cell.value)
            need = max(need, max(text_w(cell, part) for part in text.split(" (")))
        widths.append(max(6, round(need + 1.5, 1)))
        ws.column_dimensions[get_column_letter(c)].width = widths[-1]
        for r in range(1, last_row + 1):   # anything still wider than the column wraps
            cell = ws.cell(r, c)
            if isinstance(cell.value, str) and text_w(cell, cell.value) > widths[-1] - 1:
                al = copy(cell.alignment)
                al.wrap_text = True
                cell.alignment = al
                if r >= 3:
                    wrap_rows.add(r)

    margin = 0.3   # inches
    ws.page_margins.left = ws.page_margins.right = margin
    ws.page_margins.top, ws.page_margins.bottom = 0.5, margin   # room for the name header
    ws.page_margins.header = 0.2
    page_w, page_h = 842 - 2 * margin * 72, 595 - (0.5 + margin) * 72   # A4 landscape, points
    table_w = sum(w * 7 + 5 for w in widths) * 0.75                     # chars → px → points
    two_lines, header_h = 32, 16 + 32
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 16, 32
    # rows tall enough that, at the scale set by the width, the table also fills the height
    target = 1.15 * table_w * page_h / page_w - header_h   # 1.15: px estimate runs short
    plain = [r for r in range(3, last_row + 1) if r not in wrap_rows]
    row_h = (target - two_lines * len(wrap_rows)) / max(len(plain), 1)
    for r in range(3, last_row + 1):
        ws.row_dimensions[r].height = max(two_lines, row_h) if r in wrap_rows else min(max(row_h, 16), 40)
        for c in range(1, ncol + 1):   # taller rows → keep every value vertically centred
            al = copy(ws.cell(r, c).alignment)
            al.vertical = "center"
            ws.cell(r, c).alignment = al

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.print_options.horizontalCentered = True
    ws.print_area = f"A1:{get_column_letter(ncol)}{last_row}"
    ws.print_title_rows = "1:2"


def build_report(res: PeriodResult, emp: Employee, include_summary: bool = False) -> bytes:
    wb = load_workbook(TEMPLATE)
    ws = wb["attendance"]
    ncol = 15  # A..O
    data_style = [copy(ws.cell(3, c)._style) for c in range(1, ncol + 1)]
    total_style = [copy(ws.cell(5, c)._style) for c in range(1, ncol + 1)]
    date_fmt = ws.cell(3, 14).number_format
    ws.cell(2, 11).value = "ملحوظة"
    _clear_below(ws, 3)

    row = 3
    for r in res.days:
        d = r.day
        label = STATUS_LABEL.get(d.status)
        cin = label if label else fmt_clock(d.clock_in, d.date)
        cout = label if label else fmt_clock(d.clock_out, d.date)
        if d.status == "additional_day":
            cin = f"{cin} (additional)"
        vals = {
            1: fmt_days(r.total_deduct),
            2: "_",
            3: "_",
            4: f'"+{r.exception_earned}min"' if r.exception_earned else "_",
            5: fmt_hm(r.overtime_minutes),
            6: fmt_days(r.lateness_deduct),
            7: fmt_hm(r.late_minutes),
            8: "_",
            9: "_",
            10: fmt_hm(r.worked_minutes),
            11: d.raw_status or "_",
            12: cout,
            13: cin,
            14: datetime.combine(d.date, datetime.min.time()),
            15: d.day_name,
        }
        for c in range(1, ncol + 1):
            cell = ws.cell(row, c, vals[c])
            cell._style = copy(data_style[c - 1])
            cell.fill = NOFILL
        ws.cell(row, 14).number_format = date_fmt
        for f in d.edited:
            ws.cell(row, EDITED_COL[f]).fill = YELLOW
        row += 1

    k = res.kpis
    totals = {1: fmt_days(k["deductions_total"], "0 day"),
              2: "0 day",
              5: fmt_hm(k["overtime_minutes"], "00:00 hrs"),
              6: fmt_days(k["deduct_lateness"], "0 day"),
              8: "0 day",
              10: fmt_hm(k["worked_total_minutes"], "00:00 hrs")}
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c, totals.get(c))
        cell._style = copy(total_style[c - 1])

    # print setup (used by Excel printing and the PDF): one full A4 landscape page,
    # employee + period in the page header
    _fit_to_page(ws, ncol, row)
    period = f"{res.days[0].day.date} → {res.days[-1].day.date}" if res.days else ""
    ws.oddHeader.center.text = f"{emp.name} ({emp.code})   {period}".replace("&", "&&")  # & is a header code

    # actions sheet is kept for layout parity; no rule produces manual actions
    _clear_below(wb["actions"], 3)

    if include_summary:
        _summary_sheet(wb, res, emp)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_pdf(res: PeriodResult, emp: Employee, include_summary: bool = False) -> bytes:
    """Build the Excel report, then convert it to PDF with LibreOffice (headless)."""
    wb = load_workbook(io.BytesIO(build_report(res, emp, include_summary)))
    del wb["actions"]   # always empty → would only add a blank page
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "report.xlsx"
        wb.save(src)
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise RuntimeError("LibreOffice is not installed (needed for PDF export)")
        # private profile dir so it works even while LibreOffice is open on the desktop
        proc = subprocess.run([soffice, f"-env:UserInstallation=file://{tmp}/profile", "--headless",
                               "--convert-to", "pdf", "--outdir", tmp, str(src)],
                              capture_output=True, text=True, timeout=120)
        pdf = src.with_suffix(".pdf")
        if not pdf.exists():
            raise RuntimeError(f"PDF conversion failed: {proc.stderr or proc.stdout}")
        return pdf.read_bytes()


def summary_rows(res: PeriodResult) -> list[tuple[str, str, str]]:
    k = res.kpis
    return [
        ("Time & attendance", "اجمالي عدد ساعات العمل الرسمية",
         f"{fmt_hm(k['worked_regular_minutes'], '00:00 hrs')} from {fmt_hm(k['official_minutes'], '00:00 hrs')}"),
        ("Time & attendance", "اجمالي الاوفر تايم", fmt_hm(k["overtime_minutes"], "00:00 hrs")),
        ("Time & attendance", "اجمالي رصيد ساعات العمل", fmt_hm(k["balance_minutes"], "00:00 hrs")),
        ("Attendance", "اجمالي عدد ايام الحضور", str(k["attendance_days"])),
        ("Attendance", "عدد حالات خصم ربع يوم (تأخير أكثر من 10 دقائق)", str(k["late_quarter_count"])),
        ("Attendance", "عدد حالات خصم ثلث يوم (تأخير أكثر من 30 دقيقة)", str(k["late_third_count"])),
        ("Attendance", "عدد حالات تأخير يوم الجمعة (أكثر من 30 دقيقة)", str(k["friday_late_count"])),
        ("Attendance", "عدد الاستثناءات المكتسبة من الإضافي / المستخدمة",
         f"{k['morning_exceptions_earned']} / {k['morning_exceptions_used']}"),
        ("Attendance", "عدد مرات تجاوز الاستثناء (أكثر من 30 دقيقة)", str(k["exception_exceeded_count"])),
        ("Attendance", "عدد أيام الغياب بدون إذن", f"{k['absence_unexcused']} → {fmt_days(k['deduct_absence'], '0 day')}"),
        ("Attendance", "عدد أيام الغياب بإذن", str(k["absence_excused"])),
        ("Deductible", "إجمالي الخصومات", fmt_days(k["deductions_total"], "0 day")),
    ]


def _summary_sheet(wb, res, emp):
    ws = wb.create_sheet("summary")
    hdr = Font(name="Arial", size=11, bold=True)
    grey = PatternFill("solid", fgColor="FFB7B7B7")
    ws.append(["Section", "KPI", "Value"])
    for c in ws[1]:
        c.font, c.fill = hdr, grey
    for row in summary_rows(res):
        ws.append(list(row))
    for col, w in zip("ABC", (20, 48, 28)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(horizontal="center", vertical="center")

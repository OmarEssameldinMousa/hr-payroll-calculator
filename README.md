# HR payroll engine — Streamlit prototype

```bash
pip install -r requirements.txt
streamlit run app.py      # opens http://localhost:8501
pytest -q                 # rule tests
```

Flow: one payroll group (Office) with one payroll whose rules are fixed and shown
read-only. Upload the device .xls, then set the employee's official check-in,
check-out and day off. Fix any missing punches or unknown statuses in the
"Fix data problems" table (fixed cells are yellow in the export). Review the results (per day, KPIs, ledger) and export the
HR report. The file is named after the employee.

Rules: late ≤10m nothing, >10m–30m −1/4, >30m −1/3. On Fridays only >30m
counts (−1/3). Overtime is time stayed after the official check-out; more than
1:00 earns a flat 30-minute allowance on the next day, and being late beyond it
costs −1/3. A day off or absence in between cancels the allowance. Absence
without permission costs −2 days. The official day off is a weekly vacation: no
deductions and no overtime. The ملحوظة column holds the raw status
text from the device.

- Engine rules: `hr_engine/engine.py`. Full spec: `BUILD_SPEC.md`.
- Report layout comes from `assets/hr_report_template.xlsx`. Replace it to change
  the template, as long as the column positions stay the same.

## Deploy (Streamlit Community Cloud)

Main file: `app.py`. Python packages come from `requirements.txt`; `packages.txt`
installs LibreOffice and an Arabic font, which the PDF export needs.

# Attendance

`fill_attendance.py` fills the SEMIS bulk **attendance** template (`SEMIS - Learners Attendance - ... .xlsx`)
with `present` / `absent` for every school day, using the **Attendance** sheet of each tester workbook.
General setup, tester selection and common options are in the [root README](../README.md).

```
python attendance/fill_attendance.py --tester 1
```

Downloaded attendance templates go in `attendance/templates/`; results are written to `attendance/output/`.
Only the Standard 1 template is included as an example: for other standards press Enter after saving the
downloaded template, or type `s` to skip. `--report-only` shows the change reports for every standard without templates.

## Extra option

| Option | Meaning |
|---|---|
| `--report-only` | Do not fill templates; just build the change reports from the calendar |

## Tester workbook: Attendance sheet

Headers in row 4, data from row 5: ID, Learner, Academic year, Term, Standard, School, From, To, School days,
Present, Late, Absent, Absent on, Late on. "Absent on" holds dates like `04 Oct, 09 Oct` with no year; the year
comes from the academic year (August to December = first year, otherwise second year).

## Rules the script applies

1. **The calendar wins.** Term start/end, weekends and public holidays come from the calendar, not the test data.
2. **Only school days are filled.** Blank template cells are school days; `Non School Day` cells are never touched.
3. **Absent dates that are not school days are moved** (weekend, holiday, outside the term, or outside the
   period the learner was at the school) to the nearest free school day in the same term. The number of
   absences per learner is preserved, and every move is recorded in the report.
4. **Late counts as present**, because SEMIS only offers `present` / `absent`.
5. **Transfers and dropouts**: a learner whose record covers only part of a term is filled only for those days.
6. **Matching** is by the last 4 digits of the LIN (the test data `ID`), standard and academic year. School or
   class differences are reported, not blocked.
7. Dropdown values are written in lower case: `present`, `absent`.

## Output

`output/<tester selection>/`:

- `<template name> - filled.xlsx`
- `Attendance report - Standard N of YYYY.xlsx` with sheets **Summary**, **Date changes** (every moved date and
  why), **Term dates check** (test data versus calendar) and **Issues and notes**.

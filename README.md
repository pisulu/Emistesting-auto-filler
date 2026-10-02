# SEMIS attendance filler

A guided script that fills the bulk **attendance templates** downloaded from SEMIS, using the
per-tester test workbooks (`SEMIS_Test_Tester_<n>.xlsx`) and the school calendar.

For each standard it tells you what to download, fills the template with `present` / `absent`,
writes a report of every date it had to change, and asks whether to continue with the next standard.

```
========================================================================
 Download attendance for Standard 1 of 2019  (academic year 2018-2019)
   Starting date : 10-09-2018   (Term 1 opens)
   Closing date  : 19-07-2019   (Term 3 closes)
   Learners      : 21 (...)
========================================================================
 Filled <template>.xlsx: 21 learners, 3803 present, 271 absent -> ... - filled.xlsx
 Dates adjusted: 26   Notes: 2   Report: Attendance report - Standard 1 of 2019.xlsx

 Proceed with Standard 2 of 2020? [Y/n]:
```

## Requirements

- Python 3.9+
- `pip install -r requirements.txt` (only `openpyxl`)

## Quick start

Run the script **from the folder that holds your data**; it looks for everything in the current folder.

```
cd my-data-folder
python path/to/fill_attendance.py
```

Try it with the bundled data first:

```
cd usage-example
python ../fill_attendance.py --tester 1
```

## What goes in your data folder

| File | Purpose |
|---|---|
| `SEMIS_Test_Tester_<n>.xlsx` | One per tester. Needs an **Attendance** sheet (headers in row 4, data from row 5: ID, Learner, Academic year, Term, Standard, School, From, To, School days, Present, Late, Absent, Absent on, Late on) and optionally an **Enrolment plan** sheet. |
| `calendar.json` (or `calendar script.json`) | The school calendar export (see below). |
| `SEMIS - Learners Attendance - ... .xlsx` | The attendance template(s) you download from SEMIS. Several can sit in the folder; the right one is chosen by standard and academic year read from inside the file. |

### Calendar format

The script reads `schoolCalendar[]`, and for each entry:

- `academicYear.code` (e.g. `"2019"` for 2018-2019), `label`, `startDate`, `endDate`
- `classPeriods[]` with `key` (`term1`..`term3`), `startDate`, `endDate`
- `holidays[]` with `date` and `event`
- `weekDays` (`monday`: true, ... `saturday`: false)

Opening/closing dates shown to the user are the first term start and last term end.

## Choosing testers

Each tester has their own workbook, so the script asks which to use when more than one is present:

```
python fill_attendance.py                 # asks: 1, 1,3 or all
python fill_attendance.py --tester 2      # Tester 2 only
python fill_attendance.py --tester all    # every tester into the same templates
```

`--tester all` is the one to use when a downloaded class template contains learners from several
testers: a single filled file can then be uploaded once. Learners in the template with no data in the
selected workbooks are left blank and listed in the report.

## Options

| Option | Meaning |
|---|---|
| `--tester` | `1`, `1,3` or `all` (asked if omitted and several workbooks exist) |
| `--templates DIR` | Where the downloaded templates are (default: current folder) |
| `--data-dir DIR` | Where the tester workbooks are (default: current folder) |
| `--calendar FILE` | Calendar JSON (default: `calendar.json` / `calendar script.json`) |
| `--out DIR` | Output folder (default: `./output`) |
| `--start-standard N` | Begin at Standard N |
| `--report-only` | Do not fill templates; just build the change reports from the calendar |
| `--yes` | Never prompt; skip standards whose template is missing |

## Output

Written to `output/<Tester 1 | Testers 1,3 | All testers>/`:

- `<template name> - filled.xlsx` — the template with attendance filled. The original is never changed.
- `Attendance report - Standard N of YYYY.xlsx` with sheets:
  - **Summary**: counts of cells filled and dates moved.
  - **Date changes**: every absent date moved, the new date and the reason.
  - **Term dates check**: test-data term dates versus the calendar.
  - **Issues and notes**: anything that did not match (school, class, counts, ID clashes).

## Rules the script applies

1. **The calendar wins.** Term start/end, weekends and public holidays come from the calendar.
2. **Only school days are filled.** Blank template cells are school days; `Non School Day` cells are never touched.
3. **Absent dates that are not school days are moved** (weekend, holiday, outside the term, or outside the
   period the learner was at the school) to the nearest free school day in the same term. The number of
   absences per learner is preserved, and the move is recorded in the report.
4. **Late counts as present**, because SEMIS only offers `present` / `absent`.
5. **Transfers and dropouts**: a learner whose record covers only part of a term is filled only for those days.
6. **Matching** is by the last 4 digits of the LIN (the test data `ID`), standard and academic year.
   School or class differences are reported, not blocked.
7. Dropdown values are written in lower case: `present`, `absent`.
8. Test workbooks are never modified.

## Notes

- Dates in "Absent on" have no year (`04 Oct`); the year is derived from the academic year
  (August to December = first year, otherwise second year).
- The usage example uses dummy test data.

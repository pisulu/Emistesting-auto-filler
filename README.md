# SEMIS test-data loaders

Guided scripts that fill the bulk-upload templates downloaded from SEMIS, using the per-tester test
workbooks (`SEMIS_Test_Tester_<n>.xlsx`) and the school calendar.

| Folder | What it fills | Script |
|---|---|---|
| [`attendance/`](attendance/README.md) | Learner attendance (`present` / `absent`) | `fill_attendance.py` |
| [`performance/`](performance/README.md) | Term scores per subject (grade level and term remark fill themselves) | `fill_performance.py` |

Both scripts work the same way. For each standard (and academic year) in the test data they:

1. tell you what to download from SEMIS (standard, year, start and closing date from the calendar),
2. fill the template you downloaded,
3. write an Excel report of what was changed or could not be entered,
4. ask whether to continue with the next standard.

```
========================================================================
 Download attendance for Standard 1 of 2019  (academic year 2018-2019)
   Starting date : 10-09-2018   (Term 1 opens)
   Closing date  : 19-07-2019   (Term 3 closes)
   Learners      : 21 (...)
========================================================================
```

## Where things go

```
calendar.json                      shared: school calendar
SEMIS_Test_Tester_1.xlsx ... _5    shared: one workbook per tester (used by both attendance and performance)
semis_common.py                    shared code
attendance/
    templates/                     put the downloaded attendance templates here
    output/                        filled templates and reports are written here
performance/
    templates/                     put the downloaded performance templates here
    output/
```

The calendar and the tester workbooks sit in the **repo root** because both scripts use them: attendance reads
the **Attendance** sheet, performance reads the **Marks** sheet. Add `SEMIS_Test_Tester_2.xlsx`, `_3`, ... next to
the first one. The repo already contains Tester 1 and a Standard 1 template for each script as an example, with the
results in `output/`.

## Requirements

- Python 3.9+
- `pip install -r requirements.txt` (only `openpyxl`)

## Running

From anywhere (the paths above are the defaults):

```
python attendance/fill_attendance.py
python performance/fill_performance.py
```

### Calendar format

The scripts read `schoolCalendar[]`, and for each entry:

- `academicYear.code` (e.g. `"2019"` for 2018-2019), `label`, `startDate`, `endDate`
- `classPeriods[]` with `key` (`term1`..`term3`), `startDate`, `endDate`
- `holidays[]` with `date` and `event`
- `weekDays` (`monday`: true, ... `saturday`: false)

The opening and closing dates shown are the first term start and last term end.

## Choosing testers

Each tester has their own workbook, so the scripts ask which to use when more than one is present:

```
python attendance/fill_attendance.py                 # asks: 1, 1,3 or all
python attendance/fill_attendance.py --tester 2      # Tester 2 only
python attendance/fill_attendance.py --tester all    # every tester into the same templates
```

`--tester all` is the one to use when a downloaded class template contains learners from several testers:
a single filled file can then be uploaded once. Learners in the template with no data in the selected
workbooks are left blank and listed in the report.

## Options common to both scripts

| Option | Meaning |
|---|---|
| `--tester` | `1`, `1,3` or `all` (asked if omitted and several workbooks exist) |
| `--data-dir DIR` | Folder with `calendar.json` and the tester workbooks (default: repo root) |
| `--templates DIR` | Where the downloaded templates are (default: `<module>/templates`) |
| `--out DIR` | Output folder (default: `<module>/output`); results go in `<out>/<Tester 1 \| Testers 1,3 \| All testers>/` |
| `--calendar FILE` | Calendar JSON (default: `calendar.json` in `--data-dir`) |
| `--start-standard N` | Begin at Standard N |
| `--yes` | Never prompt; skip standards whose template is missing |

Test workbooks and downloaded templates are never modified: filled copies are written to the output folder.

## Notes

- The data in this repo is dummy test data.
- Anything that did not match is listed in the **Issues and notes** sheet of the report.

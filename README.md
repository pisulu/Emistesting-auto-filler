# SEMIS test-data loaders

Guided scripts that fill the bulk-upload templates downloaded from SEMIS, using the per-tester test
workbooks (`SEMIS_Test_Tester_<n>.xlsx`) and the school calendar.

| Folder | What it fills | Script |
|---|---|---|
| [`attendance/`](attendance/README.md) | Learner attendance (`present` / `absent`) | `fill_attendance.py` |
| [`performance/`](performance/README.md) | Term scores per subject (grade level and term remark fill themselves) | `fill_performance.py` |
| `sample-data/` | Dummy tester workbook and calendar used by the examples | |
| `semis_common.py` | Code shared by both scripts (calendar, tester selection, prompts) | |

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

## Requirements

- Python 3.9+
- `pip install -r requirements.txt` (only `openpyxl`)

## Running

Run a script **from the folder that holds your data**; it looks for everything in the current folder
(or use `--data-dir` and `--templates`). The bundled examples:

```
cd attendance/usage-example
python ../fill_attendance.py --data-dir ../../sample-data

cd performance/usage-example
python ../fill_performance.py --data-dir ../../sample-data
```

## Your data folder

| File | Purpose |
|---|---|
| `SEMIS_Test_Tester_<n>.xlsx` | One per tester. Attendance reads the **Attendance** sheet, performance reads the **Marks** sheet; both can use the **Enrolment plan** sheet. |
| `calendar.json` (or `calendar script.json`) | The school calendar export (format below). Looked up in `--data-dir`, then the current folder. |
| the downloaded template(s) | `SEMIS - Learners Attendance - ... .xlsx` or `SEMIS - Learners Performance - ... .xlsx`. Several can sit in one folder; each is matched to the right standard and year. |

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
python fill_attendance.py                 # asks: 1, 1,3 or all
python fill_attendance.py --tester 2      # Tester 2 only
python fill_attendance.py --tester all    # every tester into the same templates
```

`--tester all` is the one to use when a downloaded class template contains learners from several testers:
a single filled file can then be uploaded once. Learners in the template with no data in the selected
workbooks are left blank and listed in the report.

## Options common to both scripts

| Option | Meaning |
|---|---|
| `--tester` | `1`, `1,3` or `all` (asked if omitted and several workbooks exist) |
| `--templates DIR` | Where the downloaded templates are (default: current folder) |
| `--data-dir DIR` | Where the tester workbooks are (default: current folder) |
| `--calendar FILE` | Calendar JSON |
| `--out DIR` | Output folder (default: `./output`); results go in `output/<Tester 1 \| Testers 1,3 \| All testers>/` |
| `--start-standard N` | Begin at Standard N |
| `--yes` | Never prompt; skip standards whose template is missing |

Test workbooks and downloaded templates are never modified: filled copies are written to the output folder.

## Notes

- The usage examples use dummy test data.
- Anything that did not match is listed in the **Issues and notes** sheet of the report.

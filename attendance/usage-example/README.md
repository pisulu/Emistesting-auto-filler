# Attendance usage example

Contents:

- `SEMIS - Learners Attendance - Standard 1 - A (1).xlsx` – a template downloaded from SEMIS (Standard 1, 2018-2019)
- `output/` – what the script produced from it
- the tester workbook and calendar live in `../../sample-data/`

Run it (from this folder):

```
python ../fill_attendance.py --data-dir ../../sample-data --tester 1
```

Answer `Y` to move on to the next standard. Only the Standard 1 template is included, so for the other
standards press Enter after downloading the template into this folder, or type `s` to skip it.

To see the change reports for every standard without any templates:

```
python ../fill_attendance.py --data-dir ../../sample-data --tester 1 --report-only
```

# Usage example

Contents:

- `SEMIS_Test_Tester_1.xlsx` – test data for Tester 1
- `calendar script.json` – school calendar
- `SEMIS - Learners Attendance - Standard 1 - A (1).xlsx` – a template downloaded from SEMIS for Standard 1 of 2019 (stream A)
- `output/` – what the script produced from the three files above

Run it (from this folder):

```
python ../fill_attendance.py --tester 1
```

Answer `Y` to move on to the next standard. Only the Standard 1 template is included, so for the other
standards press Enter after downloading the template into this folder, or type `s` to skip it.

To see the change reports for every standard without any templates:

```
python ../fill_attendance.py --tester 1 --report-only
```

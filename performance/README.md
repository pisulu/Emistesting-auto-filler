# Performance

`fill_performance.py` fills the SEMIS bulk **performance** template (`SEMIS - Learners Performance - ... .xlsx`)
with the term scores from the **Marks** sheet of each tester workbook.
General setup, tester selection and common options are in the [root README](../README.md).

```
python performance/fill_performance.py --tester 1
```

Downloaded performance templates go in `performance/templates/`; results are written to `performance/output/`.
Only the Standard 1 template is included as an example: for other standards press Enter after saving the
downloaded template, or type `s` to skip.

## How the template works

One file covers a whole academic year: Term 1, Term 2 and Term 3 sit side by side. Each term has a score column
per subject, a **Subject Grade** column (Level 1 to Level 4) and a **Term Remarks** column. The last two are
formulas in the template and fill themselves in, so the script writes **scores only**.

The subjects in the template differ between infant (Standards 1-2), junior (3-5) and senior (6-8) classes.
The script reads the subject columns from the template, so it fills whatever that template contains.

## Extra option

| Option | Meaning |
|---|---|
| `--subject-map FILE` | JSON `{"SEMIS subject": ["test-data subject", ...]}` that overrides the default mapping |

## Subject mapping

SEMIS and the test data name subjects differently. Defaults (first listed subject that has a score is used):

| SEMIS subject | Taken from the Marks sheet |
|---|---|
| Mathematics, Chichewa, English, Expressive Arts, Agriculture | same name |
| Primary Science | Science & Technology |
| Social Studies | Social & Environmental Sciences |
| Bible Knowledge | Religious Education, otherwise Life Skills |

So Standards 1-4 (Life Skills only) feed Bible Knowledge from Life Skills, and Standards 5-8 from Religious
Education; Life Skills in those standards has no SEMIS column and is listed in the report as not entered.
Edit `SUBJECT_MAP` in the script, or pass `--subject-map`, to change this.

## Which academic year is a template for?

The performance template does not contain its academic year. The script decides, in this order:

1. a year in the **file name**, e.g. `SEMIS - Learners Performance - Standard 3 - A - 2022.xlsx` (or `2021-2022`);
2. the only year in which that standard appears in the test data;
3. the year whose learners are best covered by the learners in the file.

Name the file with the year when a standard appears in several years (repeaters). The match is printed
and written to the report.

## Rules the script applies

1. Only score cells are written; the formula columns are not touched.
2. Scores are written as whole numbers; a missing score stays blank.
3. A test-data subject with no column in the template is not entered, and a template subject with no score is left blank. Both are reported.
4. Learners are matched by the last 4 digits of the LIN; learners in the template with no marks are left blank.
5. The workbook is set to recalculate when opened. **Open the filled file in Excel and save it once before
   uploading**, so the grade levels and remarks are stored in the file.

## Output

`output/<tester selection>/`:

- `<template name> - filled.xlsx`
- `Performance report - Standard N of YYYY.xlsx` with sheets:
  - **Summary**
  - **Term remarks**: total, average, the remark SEMIS should show, and the test-data average and remark
  - **Scores entered**: every score with the grade level SEMIS should show
  - **Issues and notes**

The test data uses five remark bands (Excellent, Very good, Good, Pass, Needs improvement) and grades A-F,
while SEMIS uses four remarks and Level 1-4, so the **Remark vs test data** column shows `Differs` where the two
scales do not line up (for example "Needs improvement" with an average of 40 or more is AVERAGE in SEMIS).

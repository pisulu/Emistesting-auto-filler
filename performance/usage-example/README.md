# Performance usage example

Contents:

- `SEMIS - Learners Performance - Standard 1 - A.xlsx` – a template downloaded from SEMIS (Standard 1, 2018-2019)
- `output/` – what the script produced from it
- the tester workbook and calendar live in `../../sample-data/`

Run it (from this folder):

```
python ../fill_performance.py --data-dir ../../sample-data --tester 1
```

Answer `Y` to move on to the next standard. Only the Standard 1 template is included, so for the other
standards press Enter after downloading the template into this folder, or type `s` to skip it.

Open the filled file in Excel and save it once before uploading to SEMIS, so the grade levels and term
remarks (formulas) are stored in the file.

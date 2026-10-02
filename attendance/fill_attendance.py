#!/usr/bin/env python3
"""Guided bulk-attendance filler for the SEMIS lifecycle test.

Shared data (calendar.json, SEMIS_Test_Tester_<n>.xlsx) is read from the repo root, downloaded templates
from attendance/templates/ and results go to attendance/output/ (override with --data-dir / --templates / --out).

For each standard (and academic year) in the tester's test data the script:
  1. tells you what to download from SEMIS (standard, year, start and closing date
     taken from "calendar script.json"),
  2. fills the downloaded "SEMIS - Learners Attendance - ..." template with
     `present` / `absent` using the Attendance sheet of the tester workbook,
  3. writes a change report for every date that had to be adjusted to fit the calendar,
  4. asks whether to continue with the next standard.

Usage:
    python fill_attendance.py                    # interactive: asks which tester(s), then guides each standard
    python fill_attendance.py --tester 2         # Tester 2 only (SEMIS_Test_Tester_2.xlsx)
    python fill_attendance.py --tester all       # all testers into the same templates (one upload per class)
    python fill_attendance.py --templates "C:\\Users\\me\\Downloads"
    python fill_attendance.py --report-only      # no templates: just produce the change reports
    python fill_attendance.py --yes              # never ask, skip standards whose template is missing

Rules applied (see the report for what changed):
  * The calendar decides term start/end, weekends and public holidays.
  * Template cells that are blank are school days; "Non School Day" cells are never touched.
  * A listed absent date that is not a school day (weekend, holiday, outside the term or
    outside the period the learner was at the school) is moved to the nearest free
    school day of the same term. The number of absences per learner is preserved.
  * "Late" is not a SEMIS option, so late days are recorded as `present`.
  * The tester workbook itself is never modified.
"""
import argparse
import datetime as dt
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))   # shared helpers live in the repo root

import openpyxl
from openpyxl.styles import Font

from semis_common import (SKIP_SHEETS, add_common_args, ask, code_of, dmy, explain_no_template, iso,
                          list_workbooks, prepare, sheet, std_num)

PRESENT, ABSENT, NON_SCHOOL = "present", "absent", "Non School Day"
MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}
DATE_HEADER = re.compile(r"^\d{4}-\d{2}-\d{2}$")


# --------------------------------------------------------------------------- test data
@dataclass
class Record:
    tester: int
    row: int
    id: str
    name: str
    label: str
    term: int
    standard: str
    school: str
    start: dt.date
    end: dt.date
    days: int
    late: int
    absent: int
    absent_on: list = field(default_factory=list)   # parsed dates
    late_on: list = field(default_factory=list)
    problems: list = field(default_factory=list)


def parse_days(text, label, problems, what):
    """'04 Oct, 09 Oct' -> dates; the year comes from the academic year label."""
    y1, y2 = int(label[:4]), int(label[5:])
    out = []
    for part in (text or "").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            day, mon = part.split()
            month = MONTHS[mon[:3].title()]
            out.append(dt.date(y1 if month >= 8 else y2, month, int(day)))
        except (ValueError, KeyError):
            problems.append(f"Cannot read {what} date '{part}'")
    return out


def load_records(path, tester):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Attendance"]
    out = []
    for n, r in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
        if not r or not r[0]:
            continue
        rec = Record(tester=tester, row=n, id=str(r[0]), name=r[1], label=r[2], term=int(r[3]), standard=r[4],
                     school=r[5], start=r[6].date(), end=r[7].date(), days=int(r[8] or 0),
                     late=int(r[10] or 0), absent=int(r[11] or 0))
        rec.absent_on = parse_days(r[12], rec.label, rec.problems, "absent")
        rec.late_on = parse_days(r[13], rec.label, rec.problems, "late")
        out.append(rec)
    classes = {}
    if "Enrolment plan" in wb.sheetnames:
        for r in wb["Enrolment plan"].iter_rows(min_row=5, values_only=True):
            if r and r[0]:
                classes[(tester, str(r[0]), r[2])] = r[5]
    wb.close()
    return out, classes


@dataclass
class Job:
    std: int
    standard: str
    label: str
    code: str
    records: list

    @property
    def learners(self):
        return sorted({r.id for r in self.records})


def build_jobs(records):
    groups = defaultdict(list)
    for r in records:
        groups[(std_num(r.standard), code_of(r.label), r.standard, r.label)].append(r)
    return [Job(k[0], k[2], k[3], k[1], v) for k, v in sorted(groups.items())]


# --------------------------------------------------------------------------- planning
@dataclass
class Segment:
    rec: Record
    start: dt.date
    end: dt.date
    absent: set


def term_reference(records):
    """Most common (from, to) per (tester, year, term): what a learner present all term has."""
    c = defaultdict(Counter)
    for r in records:
        c[(r.tester, r.label, r.term)][(r.start, r.end)] += 1
    return {k: v.most_common(1)[0][0] for k, v in c.items()}


def why_moved(d, cal, seen, start, end, term_range):
    if d in seen:
        return "Duplicate date in test data"
    if d.weekday() not in cal.weekdays:
        return f"{d.strftime('%A')} (weekend)"
    if d in cal.holidays:
        return f"Public holiday: {cal.holidays[d]}"
    if not (term_range[0] <= d <= term_range[1]):
        return f"Outside calendar term ({dmy(term_range[0])} to {dmy(term_range[1])})"
    if not (start <= d <= end):
        return f"Outside the period learner was at the school ({dmy(start)} to {dmy(end)})"
    return "Not a school day in the template"


def plan_records(records, cal, school_days, ref, adjustments, issues):
    """Turn test records into Segments (who is present/absent on which day)."""
    segments = defaultdict(list)
    for rec in records:
        t_start, t_end = cal.terms[rec.term]
        r_from, r_to = ref[(rec.tester, rec.label, rec.term)]
        # Full-term learners follow the calendar; transfers/dropouts keep their own dates.
        start = max(rec.start if rec.start > r_from else t_start, t_start)
        end = min(rec.end if rec.end < r_to else t_end, t_end)
        window = sorted(d for d in school_days if start <= d <= end)
        wset, taken, pending = set(window), set(), []
        for d in sorted(rec.absent_on):
            (taken.add(d) if d in wset and d not in taken else pending.append(d))
        seen = set(taken)
        for d in pending:
            reason = why_moved(d, cal, seen, start, end, (t_start, t_end))
            free = [x for x in window if x not in taken]
            if not free:
                issues.append((rec.id, rec.name, rec.label, rec.term,
                               f"No free school day to move absence of {dmy(d)} ({reason})"))
                continue
            late = set(rec.late_on)
            new = min(free, key=lambda x: (x in late, abs((x - d).days), x))
            taken.add(new)
            adjustments.append(dict(tester=rec.tester, id=rec.id, name=rec.name, label=rec.label, term=rec.term,
                                    school=rec.school, old=d, new=new, reason=reason))
        for p in rec.problems:
            issues.append((rec.id, rec.name, rec.label, rec.term, p))
        if rec.absent != len(rec.absent_on):
            issues.append((rec.id, rec.name, rec.label, rec.term,
                           f"Absent count is {rec.absent} but {len(rec.absent_on)} dates are listed "
                           "(dates were used)"))
        segments[rec.id].append(Segment(rec, start, end, taken))
    return segments


# --------------------------------------------------------------------------- templates
@dataclass
class Template:
    path: Path
    code: str            # academic year code from the file, "" when the template has no year column
    standard: str
    stream: str
    school: str
    ids: set             # last 4 digits of every LIN in the template
    first: dt.date       # first and last date column
    last: dt.date
    days: set            # dates that are blank in the first learner row (= school days)


def month_sheets(wb):
    return [ws for ws in wb.worksheets if ws.title not in SKIP_SHEETS]


def peek_template(path):
    """Read standard / year / stream / learners / dates. Raises ValueError with the reason.

    Columns are found by header name: templates downloaded with different options have a different layout.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = [n for n in wb.sheetnames if n not in SKIP_SHEETS]
        if not sheets:
            raise ValueError("no data sheet (this does not look like an attendance template)")
        rows = list(wb[sheets[0]].iter_rows(min_row=2, values_only=True))
        dates, days = set(), set()
        for n in sheets:
            top = list(wb[n].iter_rows(min_row=2, max_row=4, values_only=True))
            for i, h in enumerate(top[0] if top else ()):
                if isinstance(h, str) and DATE_HEADER.match(h.strip()):
                    d = iso(h.strip())
                    dates.add(d)
                    if len(top) >= 3 and (i >= len(top[2]) or top[2][i] != NON_SCHOOL):
                        days.add(d)
    finally:
        wb.close()
    if not dates:
        raise ValueError("no date columns in row 2 (this does not look like an attendance template)")
    head = [str(h or "").strip().lower() for h in rows[0]]
    if "lin" not in head:
        raise ValueError("no LIN column in row 2")
    i_lin, i_std = head.index("lin"), head.index("grade") if "grade" in head else None
    if i_std is None:
        raise ValueError("no Grade column in row 2")

    def col(name):
        return head.index(name) if name in head else None

    def cell(row, i):
        return str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else ""
    learners = [r for r in rows[2:] if len(r) > i_lin and r[i_lin]]
    if not learners:
        raise ValueError("no learner rows found (was the class empty when it was downloaded?)")
    year = re.search(r"\d{4}", cell(learners[0], col("academic year_package")))
    return Template(Path(path), year.group() if year else "", cell(learners[0], i_std),
                    cell(learners[0], col("stream")), cell(learners[0], col("school")),
                    {str(r[i_lin])[-4:] for r in learners}, min(dates), max(dates), days)


def find_templates(folder, out_root):
    """(templates, skipped) for every attendance template in the folder."""
    cands, skipped = list_workbooks(folder, out_root)
    found = []
    for p in cands:
        try:
            found.append(peek_template(p))
        except Exception as e:
            skipped.append((p.name, str(e) or type(e).__name__))
    return found, skipped


def best_job(tpl, all_jobs, calendars):
    """The standard/year job a template belongs to, or None.

    With a year column the year is known. Without one it is read from the template's school days, and when
    the template has none (wrong dates at download) from the learners it contains.
    """
    m = re.search(r"\d+", tpl.standard)
    cands = [j for j in all_jobs if m and j.std == int(m.group())]
    if tpl.code:
        cands = [j for j in cands if j.code == tpl.code]
    elif tpl.days:
        cands = [j for j in cands if tpl.days & set(calendars[j.code].school_days())]
    if len(cands) > 1:
        cands.sort(key=lambda j: (len(set(j.learners) & tpl.ids) / len(j.learners),
                                  len(set(j.learners) & tpl.ids)), reverse=True)
    return cands[0] if cands else None


def describe(t):
    return (f"{t.path.name}: {t.standard} {t.stream}, "
            f"{'academic year ' + t.code if t.code else 'no year column'}, dates {dmy(t.first)} to {dmy(t.last)}, "
            f"{len(t.days)} school day(s), {len(t.ids)} learners")


def lookup(job, all_jobs, calendars, args, out_root):
    found, skipped = find_templates(args.templates, out_root)
    return [t for t in found if best_job(t, all_jobs, calendars) is job], found, skipped


def template_days(wb):
    """Dates that are not marked 'Non School Day' in the first learner row (= school days; they may already
    hold present/absent from an earlier upload), plus date->column maps."""
    days, cols = set(), {}
    for ws in month_sheets(wb):
        cols[ws.title] = {c.column: iso(c.value) for c in ws[2]
                          if isinstance(c.value, str) and DATE_HEADER.match(c.value)}
        for col, d in cols[ws.title].items():
            if ws.cell(4, col).value != NON_SCHOOL:
                days.add(d)
    return days, cols


def refusal(tpl, job, cal, records):
    """Why a template cannot be filled for this job, or "" when it can."""
    if not tpl.days & set(cal.school_days()):
        span = f"{dmy(tpl.first)} to {dmy(tpl.last)}"
        if not tpl.days:
            why = f"every date in it ({span}) is marked 'Non School Day', so there is nothing to fill"
        else:
            why = f"its school days ({span}) are outside {job.label}"
        return (f"{why}. Download it again for {job.standard} with Starting date {dmy(cal.opens)} "
                f"and Closing date {dmy(cal.closes)}.")
    if not tpl.ids & {r.id for r in records}:
        return (f"none of its {len(tpl.ids)} learners are in the selected tester data "
                "(probably another tester's class: select that tester, or use --tester all)")
    return ""


def fill_template(tpl, job, cal, records, classes, out_dir, adjustments, issues):
    why = refusal(tpl, job, cal, records)
    if why:
        return None, why
    wb = openpyxl.load_workbook(tpl.path)
    days, cols = template_days(wb)
    cal_days = set(cal.school_days())
    diff = sorted(days ^ cal_days)
    if diff:
        issues.append(("", "", job.label, "", f"Template school days differ from the calendar on "
                       f"{len(diff)} date(s), e.g. {', '.join(dmy(d) for d in diff[:5])}. "
                       "The template was followed."))
    in_tpl = tpl.ids
    mine = [r for r in records if r.id in in_tpl]
    owners = defaultdict(set)
    for r in mine:
        owners[r.id].add(r.tester)
    for lid, t in owners.items():
        if len(t) > 1:
            issues.append((lid, "", job.label, "", "Learner ID appears in more than one tester workbook "
                           f"({', '.join(f'Tester {x}' for x in sorted(t))}); check for a clash"))
    ref = term_reference(records)
    segments = plan_records(mine, cal, days, ref, adjustments, issues)

    for i in sorted(in_tpl - {r.id for r in records}):
        issues.append((i, "", job.label, "", f"Learner is in the template but has no {job.standard} "
                       f"{job.label} attendance in the selected tester data (left blank; may belong to another tester)"))
    stats = Counter()
    for ws in month_sheets(wb):
        lin_col = next((c.column for c in ws[2] if str(c.value or '').strip().lower() == 'lin'), 7)
        for r in range(4, ws.max_row + 1):
            lin = ws.cell(r, lin_col).value
            segs = segments.get(str(lin)[-4:]) if lin else None
            if not segs:
                continue
            for col, d in cols[ws.title].items():
                cell = ws.cell(r, col)
                if cell.value == NON_SCHOOL:
                    continue
                seg = next((s for s in segs if s.start <= d <= s.end), None)
                if seg is None:
                    continue
                new = ABSENT if d in seg.absent else PRESENT
                if cell.value in (PRESENT, ABSENT):
                    stats['existing'] += 1
                    stats['changed'] += cell.value != new
                cell.value = new
                stats[new] += 1
    other_school, other_class = defaultdict(set), {}
    for lid, segs in segments.items():
        for s in segs:
            if tpl.school and tpl.school.lower().find(s.rec.school.lower()) < 0:
                other_school[(lid, s.rec.name, s.rec.school)].add(s.rec.term)
            cls = classes.get((s.rec.tester, lid, s.rec.label))
            if cls and cls != tpl.stream:
                other_class[lid] = cls
    for (lid, name, school), terms in other_school.items():
        issues.append((lid, name, job.label, ", ".join(map(str, sorted(terms))),
                       f"Test data says {school} but the template is '{tpl.school}'; "
                       "filled anyway (check the learner's enrolment/transfer in SEMIS)"))
    if other_class:
        issues.append(("", "", job.label, "", f"Template stream is {tpl.stream}; the test data class differs for "
                       f"{len(other_class)} learner(s): {', '.join(sorted(other_class))}. Matched by learner ID."))
    late_days = sum(len([d for d in s.rec.late_on if s.start <= d <= s.end])
                    for ss in segments.values() for s in ss)
    stats["late_as_present"] = late_days
    if stats["existing"]:
        issues.append(("", "", job.label, "", f"{tpl.path.name} already held attendance in "
                       f"{stats['existing']} cell(s); they were overwritten ({stats['changed']} with a different value)"))
    stats["learners"] = len(segments)
    out = out_dir / f"{tpl.path.stem} - filled.xlsx"
    wb.save(out)
    return out, stats


# --------------------------------------------------------------------------- report
def write_report(job, cal, adjustments, issues, summary_lines, all_records, out_dir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.column_dimensions["A"].width = 120
    for i, line in enumerate(summary_lines, 1):
        ws.cell(i, 1, line)
    ws["A1"].font = Font(bold=True, size=14)

    sheet(wb, "Date changes",
          ["Tester", "ID", "Learner", "Academic year", "Term", "School", "Original absent date", "Day",
           "New absent date", "Day", "Days moved", "Reason"],
          [(f"Tester {a['tester']}", a["id"], a["name"], a["label"], a["term"], a["school"], dmy(a["old"]),
            a["old"].strftime("%a"), dmy(a["new"]), a["new"].strftime("%a"),
            (a["new"] - a["old"]).days, a["reason"])
           for a in sorted(adjustments, key=lambda a: (a["tester"], a["id"], a["term"], a["old"]))])

    ref = term_reference(all_records)
    per = []
    for (tester, label, term), (f, t) in sorted(ref.items()):
        if label != job.label:
            continue
        s, e = cal.terms[term]
        reported = [r.days for r in all_records if r.tester == tester and r.label == label
                    and r.term == term and (r.start, r.end) == (f, t)][0]
        n = len(list(cal.school_days(s, e)))
        per.append((f"Tester {tester}", label, term, dmy(f), dmy(t), reported, dmy(s), dmy(e), n,
                    "OK" if (f, t, reported) == (s, e, n) else "Differs - calendar used"))
    sheet(wb, "Term dates check",
          ["Tester", "Academic year", "Term", "Test data from", "Test data to", "Test data school days",
           "Calendar from", "Calendar to", "Calendar school days", "Status"], per)

    sheet(wb, "Issues and notes", ["ID", "Learner", "Academic year", "Term", "Note"],
          sorted(set(issues), key=lambda x: tuple(str(v) for v in x)))
    name = f"Attendance report - {job.standard} of {job.code}.xlsx"
    wb.save(out_dir / name)
    return out_dir / name


# --------------------------------------------------------------------------- console
def show_job(job, cal):
    schools = Counter(r.school for r in {(r.id, r.school): r for r in job.records}.values())
    print()
    print("=" * 72)
    print(f" Download attendance for {job.standard} of {job.code}  (academic year {job.label})")
    print(f"   Starting date : {dmy(cal.opens)}   (Term 1 opens)")
    print(f"   Closing date  : {dmy(cal.closes)}   (Term 3 closes)")
    if (cal.year_start, cal.year_end) != (cal.opens, cal.closes):
        print(f"   Calendar academic-year range is {dmy(cal.year_start)} to {dmy(cal.year_end)}")
    print(f"   Learners      : {len(job.learners)} ("
          + ", ".join(f"{s} {n}" for s, n in sorted(schools.items())) + ")")
    by_tester = Counter(t for t, _ in {(r.tester, r.id) for r in job.records})
    if len(by_tester) > 1:
        print("   Per tester    : " + ", ".join(f"Tester {t}: {n}" for t, n in sorted(by_tester.items())))
    print("=" * 72)


def process_job(job, all_jobs, calendars, args, classes, all_records, out_dir, out_root):
    cal = calendars[job.code]
    adjustments, issues, lines = [], [], []
    lines += [f"Attendance report: {job.standard} of {job.code} ({job.label})",
              f"Testers: {', '.join(f'Tester {t}' for t in sorted({r.tester for r in job.records}))}", "",
              f"Calendar: opens {dmy(cal.opens)}, closes {dmy(cal.closes)}; "
              f"public holidays and weekends are non-school days."]
    templates = []
    if not args.report_only:
        templates, found, skipped = lookup(job, all_jobs, calendars, args, out_root)
        if not templates:
            explain_no_template(args.templates, f"{job.standard} / {job.code}",
                                [describe(t) for t in found], skipped)
            if not args.yes:
                ans = ask(" Save the right template there, then press Enter (or type s to skip): ", "")
                if ans != "s":
                    templates = lookup(job, all_jobs, calendars, args, out_root)[0]
        if not templates:
            print(" Skipped: no template available.")
            return
    if args.report_only:
        days = set(cal.school_days())
        plan_records(job.records, cal, days, term_reference(all_records), adjustments, issues)
        lines.append("Mode: report only (school days taken from the calendar, no template filled).")
    for tpl in templates:
        out, st = fill_template(tpl, job, cal, job.records, classes, out_dir, adjustments, issues)
        if out is None:
            lines += ["", f"Template: {tpl.path.name}  ->  NOT FILLED: {st}"]
            print(f" NOT filled {tpl.path.name}: {st}")
            continue
        lines += ["", f"Template: {tpl.path.name}  ->  {out.name}",
                  f"  Learners filled : {st['learners']}",
                  f"  Cells present   : {st[PRESENT]}",
                  f"  Cells absent    : {st[ABSENT]}",
                  *([f"  Cells that already held attendance: {st['existing']} (changed: {st['changed']})"]
                    if st['existing'] else []),
                  f"  Late days recorded as present (SEMIS has no 'late'): {st['late_as_present']}"]
        print(f" Filled {tpl.path.name}: {st['learners']} learners, {st[PRESENT]} present, "
              f"{st[ABSENT]} absent -> {out.name}")
    lines += ["", f"Absent dates moved to fit the calendar: {len(adjustments)}",
              f"Issues / notes: {len(set(issues))}  (see the 'Issues and notes' sheet)"]
    rep = write_report(job, cal, adjustments, issues, lines, all_records, out_dir)
    print(f" Dates adjusted: {len(adjustments)}   Notes: {len(set(issues))}   Report: {rep.name}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    add_common_args(ap, Path(__file__).resolve().parent)
    ap.add_argument("--report-only", action="store_true", help="only build change reports")
    args = ap.parse_args()
    calendars, available, picked, out_root, out_dir = prepare(args)
    records, classes = [], {}
    for n in picked:
        recs, cls = load_records(available[n], n)
        records += recs
        classes.update(cls)
    all_jobs = build_jobs(records)
    jobs = [j for j in all_jobs if j.std >= args.start_standard]
    missing = sorted({j.code for j in jobs} - set(calendars))
    if missing:
        sys.exit(f"Calendar has no academic year for code(s): {', '.join(missing)}")

    for i, job in enumerate(jobs):
        cal = calendars[job.code]
        show_job(job, cal)
        process_job(job, all_jobs, calendars, args, classes, records, out_dir, out_root)
        if i + 1 < len(jobs):
            nxt = jobs[i + 1]
            if not args.yes and not args.report_only and ask(
                    f"\n Proceed with {nxt.standard} of {nxt.code}? [Y/n]: ") not in ("y", "yes"):
                print(" Stopped.")
                return
    print("\n All standards done. Output folder:", out_dir)


if __name__ == "__main__":
    main()

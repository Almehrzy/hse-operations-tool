# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Weekly HSE Statistics Report (HSE-FRM-00015) from the tool's records.

    python weekly_hse.py <export dir> [--week YYYY-MM-DD] [--out-dir DIR]

<export dir> holds the collections exported with ArtifactData (`list` with `out_dir`): observations, violations, permits,
consultant, equipmentInspections, manpower, dailyLog, incidents. The week runs Saturday to Friday and belongs to the month of its
Wednesday (as the trend report). The script finds the previous week's report, works out this week's numbers, and writes
values.json for weekly_hse.ps1, which copies the previous report through Excel and carries month and project-to-date
totals forward. It prints what it found and what it could not (e.g. rates, which stay as they were).
"""
import datetime as dt, glob, json, os, re, sys
import openpyxl
from config import CFG
import trend_report as TR

DAYS = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
INC_ROWS = {"ftl": 20, "lti": 21, "mtc": 22, "fac": 23, "fire": 24, "mvi": 25, "illness": 26, "env": 27, "damage": 28, "nearmiss": 29}
MON_ROWS = {"leadership": 40, "joint": 41, "observation": 42, "obsRaised": 43, "obsClosed": 44, "obsGood": 45, "stopwork": 46, "meetings": 47,
            "audit": 48, "ncIssued": 49, "ncClosed": 50, "ntp": 51, "permits": 52, "starrt": 53, "ra": 54, "inspections": 55, "rr": 56,
            "svrIssued": 57, "svrClosed": 58, "drills": 59}
TR_ROWS = {"induction": 63, "tbt": 64, "trInt": 65, "trExt": 66}


def load(export, coll):
    out = []
    for f in glob.glob(os.path.join(export, coll, "*.json")):
        d = json.load(open(f, encoding="utf-8-sig"))
        d = d.get("data", d) if isinstance(d, dict) else d
        d["_id"] = os.path.basename(f)[:-5]
        out.append(d)
    return out


def day_of(iso, off=3):
    if not iso:
        return ""
    s = str(iso)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    try:
        return TR.local_day(s, off).isoformat()
    except ValueError:
        return s[:10]


def report_files():
    base = os.path.join(CFG.get("trend", {}).get("folder", r"D:\Claude Projects\Quick Access\Weekly Report"))
    out = []
    for f in glob.glob(os.path.join(base, "**", "HSE-FRM-00015*.xls*"), recursive=True):
        m = re.findall(r"(\d{1,2})[-\s](\d{1,2})[-\s](\d{2,4})", os.path.basename(f))
        if len(m) >= 2:
            d, mo, y = m[1]
            y = int(y) + (2000 if len(y) == 2 else 0)
            try:
                out.append((dt.date(y, int(mo), int(d)), f))
            except ValueError:
                pass
    return sorted(out)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:]
    export = a[0]
    opt = dict(zip(a[1::2], a[2::2]))
    today = dt.date.today()
    ref = dt.date.fromisoformat(opt["--week"]) if "--week" in opt else today - dt.timedelta(days=((today.weekday() - 4) % 7) or 7)
    start = ref - dt.timedelta(days=(ref.weekday() - 5) % 7)          # Saturday
    end = start + dt.timedelta(days=6)
    wed = start + dt.timedelta(days=4)
    s0, s1 = str(start), str(end)
    inweek = lambda x: bool(x) and s0 <= x <= s1
    notes = []

    obs = [o for o in load(export, "observations") if not o.get("example") and inweek(day_of(o.get("createdAt")))]
    vios = load(export, "violations")
    permits = load(export, "permits")
    cons = load(export, "consultant")
    insp = load(export, "equipmentInspections")
    mp = {d.get("date") or d["_id"]: d for d in load(export, "manpower")}
    logs = [x for x in load(export, "dailyLog") if inweek(x.get("date"))]

    # manpower per day
    days, missing = [], []
    for i in range(7):
        d = str(start + dt.timedelta(days=i))
        m = mp.get(d)
        if not m:
            missing.append(d)
        g = lambda k: int((m or {}).get(k) or 0)
        days.append([DAYS[i], d, g("company_staff"), g("company_labor"), g("company_saudi"), g("sub_staff"), g("sub_labor"), g("sub_saudi"),
                     g("spec_staff"), g("spec_labor"), g("spec_saudi"), float((m or {}).get("hours") or 11)])
    if missing:
        notes.append(f"manpower missing for {len(missing)} day(s): {', '.join(missing)}")
    contractor = sum(r[2] + r[3] for r in days)
    contractor_sa = sum(r[4] for r in days)
    sub = sum(r[5] + r[6] + r[8] + r[9] for r in days)
    sub_sa = sum(r[7] + r[10] for r in days)
    hours = round(sum((r[2] + r[3] + r[5] + r[6] + r[8] + r[9]) * r[11] for r in days))
    hours_sa = round(sum((r[4] + r[7] + r[10]) * r[11] for r in days))

    def n(kind, key=None, val=None):
        return sum(1 for x in logs if x.get("kind") == kind and (key is None or x.get(key) == val))

    def att(kind, key=None, val=None):
        return sum(int(x.get("attendees") or 0) for x in logs if x.get("kind") == kind and (key is None or x.get(key) == val))

    # incidents: the Incidents tab (collection `incidents`) and older daily-log entries; RWC and dangerous occurrences
    # have no row of their own in HSE-FRM-00015 and are counted with MTC and near misses (as the page does)
    inc_map = {"rwc": "mtc", "dangerous": "nearmiss"}
    incs = [x for x in load(export, "incidents") if not x.get("example") and inweek(x.get("date"))]
    rows = {}
    for k, r in INC_ROWS.items():
        rows[r] = n("incident", "incClass", k) + sum(1 for x in incs if inc_map.get(x.get("type"), x.get("type")) == k)
    if any(x.get("type") in inc_map for x in incs):
        notes.append("an RWC or dangerous occurrence this week is counted under MTC / near miss: the form has no row for it")
    if any(rows[INC_ROWS[k]] for k in ("lti", "ftl")):
        notes.append("an LTI or fatality this week: reset 'Worked man-hours since last LTI' (row 15) by hand")
    if any(rows[INC_ROWS[k]] for k in ("lti", "ftl", "mtc")):
        notes.append("recordable incident this week: the rates in section 3 stay as they were; recalculate them as your project does")
    # SVR from the register (the source of truth), by the date column
    svr = 0
    try:
        ws = openpyxl.load_workbook(CFG["svr"]["register"], data_only=True)[CFG["svr"]["register_sheet"]]
        for r in ws.iter_rows(min_row=5, values_only=True):
            v = r[7] if len(r) > 7 else None
            dd = v.date() if isinstance(v, dt.datetime) else v if isinstance(v, dt.date) else None
            if dd and start <= dd <= end and r[0]:
                svr += 1
    except Exception as e:                                  # register locked or moved: fall back to the tool
        notes.append(f"SVR register not read ({e}); used the tool's issued violations")
        svr = sum(1 for v in vios if v.get("status") == "issued" and inweek(day_of(v.get("issuedAt"))))
    mon = {
        "leadership": n("tour", "tourType", "leadership"), "joint": n("tour", "tourType", "joint"), "observation": n("tour", "tourType", "observation"),
        "obsRaised": len(obs), "obsClosed": sum(1 for o in obs if o.get("status") == "closed"), "obsGood": sum(1 for o in obs if o.get("type") == "positive"),
        "stopwork": n("other", "otherType", "stopwork"), "meetings": n("meeting"), "audit": n("other", "otherType", "audit"),
        "ncIssued": sum(1 for c in cons if str(c.get("kind", "")).upper() == "NCR" and not c.get("example") and inweek(str(c.get("issued", ""))[:10])),
        "ncClosed": sum(1 for c in cons if str(c.get("kind", "")).upper() == "NCR" and not c.get("example") and c.get("status") == "closed" and inweek(day_of(c.get("closedAt")))),
        "ntp": n("other", "otherType", "ntp"), "permits": sum(1 for p in permits if inweek(day_of(p.get("issuedAt")))),
        "starrt": n("other", "otherType", "starrt"), "ra": n("other", "otherType", "ra"),
        "inspections": n("inspection") + sum(1 for x in insp if x.get("kind") == "monthly" and inweek(x.get("date"))),
        "rr": n("other", "otherType", "rr"), "svrIssued": svr, "svrClosed": svr, "drills": n("drill"),
    }
    for k, r in MON_ROWS.items():
        rows[r] = mon[k]
    tr = {"induction": att("induction"), "tbt": n("tbt"), "trInt": sum(1 for x in logs if x.get("kind") == "training" and x.get("scope", "internal") != "external"),
          "trExt": n("training", "scope", "external")}
    for k, r in TR_ROWS.items():
        rows[r] = tr[k]

    # detail sheets
    by_date = lambda k: sorted([x for x in logs if x.get("kind") == k], key=lambda x: (x.get("date", ""), x.get("createdAt", "")))
    tbt = [[x.get("topic", ""), x.get("by", ""), int(x.get("attendees") or 0)] for x in by_date("tbt")]
    trainings = [[x.get("topic", ""), "External" if x.get("scope") == "external" else "Internal",
                  f"{x.get('attendees') or 0}" + (f" ({x['duration']})" if x.get("duration") else ""), x.get("passed") or ""] for x in by_date("training")]
    inspections = [[x.get("topic", ""), x.get("date"), "External" if x.get("scope") == "external" else "Internal", x.get("qty") or ""] for x in by_date("inspection")]
    eqm = [x for x in insp if x.get("kind") == "monthly" and inweek(x.get("date"))]
    if eqm:
        inspections.append([f"Equipment monthly inspection ({len(eqm)} items)", max(x["date"] for x in eqm), "Internal", len(eqm)])
    tour_name = {"leadership": "Leadership Tour", "joint": "Joint Formal Weekly Tour", "observation": "Site Observation Tour", "walkthrough": "Site Walkthrough"}
    tours = [[tour_name.get(x.get("tourType"), "Site Walkthrough"), x.get("date"), "Walk through" if x.get("tourType") == "walkthrough" else tour_name.get(x.get("tourType"), ""), 1]
             for x in by_date("tour")]
    drills = [[x.get("topic", ""), x.get("date"), "Drill", x.get("location", "")] for x in by_date("drill")]

    # observations by topic, as in the trend report
    topics = {}
    for o in obs:
        cat = "Good/Positive Observation" if o.get("type") == "positive" else TR.LEGACY.get(o.get("category"), o.get("category"))
        k = re.sub(r"\s*/\s*", "/", TR.norm(cat))         # the report writes "Good / Positive", the trend "Good/Positive"
        topics.setdefault(k, [0, 0])
        topics[k][0] += 1
        topics[k][1] += 1 if o.get("status") == "closed" else 0

    prev = [f for d, f in report_files() if d < start]
    if not prev:
        raise SystemExit("No earlier HSE-FRM-00015 report found to carry totals from.")
    prev_end, prev_file = [x for x in report_files() if x[1] == prev[-1]][0]
    prev_wed = prev_end - dt.timedelta(days=2)
    same_month = (prev_wed.year, prev_wed.month) == (wed.year, wed.month)
    if (start - prev_end).days != 1:
        notes.append(f"the previous report ends {prev_end}, not the day before this week ({start}); totals carry from it anyway")
    out_dir = opt.get("--out-dir") or TR.month_folder(CFG["trend"]["folder"], wed.year, wed.month)
    name = f"HSE-FRM-00015 - WP.03 Weekly HSE report {start:%d-%m-%Y} to {end:%d-%m-%Y}.xls"
    values = {"source": prev_file, "out": os.path.join(out_dir, name), "sameMonth": same_month,
              "dates": {"from": s0, "to": s1, "report": str(end + dt.timedelta(days=1))},
              "manpower": {"D9": contractor, "E9": contractor_sa, "D10": sub, "E10": sub_sa, "D12": hours, "E12": hours_sa},
              "rows": rows, "days": days, "tbt": tbt, "trainings": trainings, "inspections": inspections, "tours": tours, "drills": drills,
              "observations": topics}
    print(json.dumps({"week": f"{s0} to {s1}", "previous": prev_file, "same_month": same_month, "out": values["out"],
                      "manpower": values["manpower"], "monitoring": mon, "training": tr,
                      "incidents": {k: rows[r] for k, r in INC_ROWS.items() if rows[r]}, "tbt": len(tbt), "notes": notes}, ensure_ascii=False, indent=1))
    return values


if __name__ == "__main__":
    v = main()
    dest = os.environ.get("WEEKLY_VALUES") or os.path.join(os.environ.get("TEMP", "."), "weekly_values.json")
    json.dump(v, open(dest, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("values:", dest)

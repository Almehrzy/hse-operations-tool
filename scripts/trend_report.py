# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Weekly SOR Trend Analysis from the page's site observations, written into the house workbook.

    python trend_report.py <observations dir or .json> [--week YYYY-MM-DD] [--template X.xlsx] [--out-dir DIR] [--replace]

- Observations: the `observations` collection exported with ArtifactData (`list` with `out_dir`), or a JSON list.
- Week: Saturday to Friday (project.json `trend.week_starts`), counted in the month that holds its Wednesday
  (`trend.month_by`); this reproduces the weeks of every month of the 2026 workbook (user, 2026-10-15).
  The file is dated the Saturday after the week, as the house files are. Default: the last full week.
- Raised = observations recorded in the week; Closed = those of them closed by now. Demo rows (`example`) are skipped.
- The newest "*TREND ANALYSIS*.xlsx" under the Weekly Report folder is copied and only that week's cells are written;
  monthly totals and the Monthly sheet are formulas already. A month with no block yet (e.g. November) is added.
- Never overwrites a file; refuses to replace non-empty week cells unless --replace.
"""
import argparse, copy, datetime as dt, glob, io, json, os, re, sys
import openpyxl
from openpyxl.utils import get_column_letter as L
from config import CFG

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
LEGACY = {"height": "Working at Height", "scaffold": "Scaffolds / Ladders / Stairways / Platforms", "excavation": "Excavations",
          "lifting": "Hoisting and Lifting Equipment / MEWP", "electrical": "Electrical Works / LOTO", "hotwork": "Fire Protection and Hot Works",
          "ppe": "Personal Protective Equipment", "housekeeping": "Housekeeping / Sanitation", "vehicles": "Vehicles and Mobile Equipment / Plant",
          "fire": "Fire Protection and Hot Works", "confined": "Confined Spaces", "other": "Other"}
FIRST_ROW, LAST_ROW, TOTAL_ROW = 5, 45, 47


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
WEEK_START = DAYS.index(CFG["trend"].get("week_starts", "Saturday").lower())
MONTH_BY = DAYS.index(CFG["trend"].get("month_by", "Wednesday").lower())


def week_of(day):
    """(first day, last day, year, month, week number) of the week holding `day`; the week belongs to the month of its MONTH_BY day."""
    start = day - dt.timedelta(days=(day.weekday() - WEEK_START) % 7)
    key = start + dt.timedelta(days=(MONTH_BY - WEEK_START) % 7)
    return start, start + dt.timedelta(days=6), key.year, key.month, (key.day - 1) // 7 + 1


def weeks_in_month(year, month):
    d = dt.date(year, month, 1)
    return sum(1 for i in range(31) if (d + dt.timedelta(days=i)).month == month and (d + dt.timedelta(days=i)).weekday() == MONTH_BY)


def load_obs(src):
    files = glob.glob(os.path.join(src, "**", "*.json"), recursive=True) if os.path.isdir(src) else [src]
    out = []
    for f in files:
        data = json.load(open(f, encoding="utf-8-sig"))
        for d in (data if isinstance(data, list) else [data]):
            out.append(d.get("data", d) if isinstance(d, dict) and "data" in d and isinstance(d["data"], dict) else d)
    return out


def local_day(iso, offset):
    t = dt.datetime.fromisoformat(str(iso).replace("Z", "+00:00"))
    return (t + dt.timedelta(hours=offset)).date() if t.tzinfo else t.date()


def blocks(ws):
    """{(month index 1-12): {"start": col, "weeks": {k: col}, "total": col}} read from rows 2-3 of the Weekly sheet."""
    out, cur = {}, None
    for c in range(3, ws.max_column + 1):
        m, w = ws.cell(row=2, column=c).value, ws.cell(row=3, column=c).value
        w = str(w or "").strip()
        if w.lower() == "week 1" and m:
            mi = [x[:3].lower() for x in MONTHS].index(str(m).strip()[:3].lower()) + 1
            cur = out.setdefault(mi, {"start": c, "weeks": {}, "total": None})
        if cur is not None and w.lower().startswith("week "):
            cur["weeks"][int(w.split()[1])] = c
        elif cur is not None and w.lower() == "total":
            cur["total"] = c
            cur = None
    return out


def add_block(wb, month, year):
    """Append a month block after the last one, styled like the previous block, and link the Monthly sheet."""
    ws, mo = wb["Weekly"], wb["Monthly"]
    bl = blocks(ws)
    last = bl[max(bl)]
    src_week, src_total = last["weeks"][1], last["total"]
    n = weeks_in_month(year, month)
    start = last["total"] + 3
    total = start + 2 * n + 1

    def clone(src_col, dst_col):
        ws.column_dimensions[L(dst_col)].width = ws.column_dimensions[L(src_col)].width or 13
        for r in range(1, TOTAL_ROW + 1):
            s, d = ws.cell(row=r, column=src_col), ws.cell(row=r, column=dst_col)
            d._style = copy.copy(s._style)
    for k in range(n):
        clone(src_week, start + 2 * k); clone(src_week + 1, start + 2 * k + 1)
        ws.cell(row=3, column=start + 2 * k, value=f"Week {k + 1}")
        ws.cell(row=4, column=start + 2 * k, value="Raised"); ws.cell(row=4, column=start + 2 * k + 1, value="Closed")
        ws.merge_cells(start_row=3, start_column=start + 2 * k, end_row=3, end_column=start + 2 * k + 1)
    clone(src_total - 1, start + 2 * n)
    clone(src_total, total); clone(src_total + 1, total + 1)
    label = MONTHS[month - 1][:3] if month != 9 else "Sep"
    ws.cell(row=2, column=start, value=label)
    ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=start + 2 * n - 1)
    ws.cell(row=2, column=total, value=label); ws.cell(row=3, column=total, value="Total")
    ws.cell(row=4, column=total, value="Raised"); ws.cell(row=4, column=total + 1, value="Closed")
    ws.merge_cells(start_row=2, start_column=total, end_row=2, end_column=total + 1)
    ws.merge_cells(start_row=3, start_column=total, end_row=3, end_column=total + 1)
    for r in range(FIRST_ROW, LAST_ROW + 1):
        ws.cell(row=r, column=total, value="=" + "+".join(f"{L(start + 2 * k)}{r}" for k in range(n)))
        ws.cell(row=r, column=total + 1, value="=" + "+".join(f"{L(start + 2 * k + 1)}{r}" for k in range(n)))
    for c in list(range(start, start + 2 * n)) + [total, total + 1]:
        ws.cell(row=TOTAL_ROW, column=c, value=f"=SUM({L(c)}{FIRST_ROW}:{L(c)}{TOTAL_ROW - 1})")
    mcol = 3 + 2 * (month - 1)           # Monthly: C/D = January, E/F = February, ...
    for r in range(4, 45):
        mo.cell(row=r, column=mcol, value=f"=Weekly!{L(total)}{r + 1}")
        mo.cell(row=r, column=mcol + 1, value=f"=Weekly!{L(total + 1)}{r + 1}")
    return blocks(ws)[month]


def month_folder(base, year, month):
    yd = os.path.join(base, str(year))
    name = MONTHS[month - 1]
    if os.path.isdir(yd):
        for d in os.listdir(yd):
            if os.path.isdir(os.path.join(yd, d)) and d.strip().lower()[:3] == name[:3].lower():
                return os.path.join(yd, d)
    return os.path.join(yd, name)


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("observations")
    ap.add_argument("--week")
    ap.add_argument("--template")
    ap.add_argument("--out-dir")
    ap.add_argument("--replace", action="store_true")
    a = ap.parse_args()
    tcfg = CFG["trend"]
    today = dt.date.today()
    last_day = (WEEK_START - 1) % 7
    ref = dt.date.fromisoformat(a.week) if a.week else today - dt.timedelta(days=((today.weekday() - last_day) % 7) or 7)   # last full week
    start, end, year, month, idx = week_of(ref)
    tpl = a.template or max(glob.glob(os.path.join(tcfg["folder"], "**", "*TREND ANALYSIS*.xlsx"), recursive=True), key=os.path.getmtime)
    out_dir = a.out_dir or month_folder(tcfg["folder"], year, month)
    out = os.path.join(out_dir, f"{end + dt.timedelta(days=1):%d-%m-%Y}{tcfg['file_suffix']}")   # issued the day after the week
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")

    wb = openpyxl.load_workbook(tpl)
    ws = wb["Weekly"]
    bl = blocks(ws)
    fixed = []                                   # month label over a Total pair that names another month
    for mi, b in bl.items():
        cell = ws.cell(row=2, column=b["total"]) if b["total"] else None
        if cell is not None and str(cell.value or "").strip()[:3].lower() != MONTHS[mi - 1][:3].lower():
            fixed.append(f"{cell.coordinate}: {cell.value} -> {MONTHS[mi - 1][:3]}")
            cell.value = MONTHS[mi - 1][:3]
    added = False
    if month not in bl:
        if month != max(bl) + 1:
            raise SystemExit(f"Workbook has no block for {MONTHS[month - 1]} and it is not the next month after the last block.")
        blk = add_block(wb, month, year); added = True
    else:
        blk = bl[month]
    if idx not in blk["weeks"]:
        raise SystemExit(f"{MONTHS[month - 1]} block has no Week {idx}.")
    rc, cc = blk["weeks"][idx], blk["weeks"][idx] + 1

    topics = {norm(ws.cell(row=r, column=2).value): r for r in range(FIRST_ROW, LAST_ROW + 1) if ws.cell(row=r, column=2).value}
    off = tcfg.get("utc_offset_hours", 0)
    raised, closed, unmapped, used = {}, {}, [], 0
    for o in load_obs(a.observations):
        if o.get("example") or not o.get("createdAt"):
            continue
        day = local_day(o["createdAt"], off)
        if not (start <= day <= end):
            continue
        cat = "Good/Positive Observation" if o.get("type") == "positive" else LEGACY.get(o.get("category"), o.get("category"))
        r = topics.get(norm(cat))
        if r is None:
            unmapped.append(f"no. {o.get('no')}: {o.get('category')}")
            r = topics.get("other")
        used += 1
        raised[r] = raised.get(r, 0) + 1
        if o.get("status") == "closed":
            closed[r] = closed.get(r, 0) + 1
    filled = [r for r in range(FIRST_ROW, LAST_ROW + 1) if ws.cell(row=r, column=rc).value not in (None, "") or ws.cell(row=r, column=cc).value not in (None, "")]
    if filled and not a.replace:
        raise SystemExit(f"Week {idx} of {MONTHS[month - 1]} already has numbers in {len(filled)} rows of the template. Use --replace to overwrite them.")
    for r in range(FIRST_ROW, LAST_ROW + 1):
        ws.cell(row=r, column=rc, value=raised.get(r) or None)
        ws.cell(row=r, column=cc, value=closed.get(r) or None)
    os.makedirs(out_dir, exist_ok=True)
    wb.save(out)
    print(json.dumps({"written": out, "template": tpl, "week": f"{start} to {end}", "month": MONTHS[month - 1], "week_no": idx,
                      "block_added": added, "observations_counted": used,
                      "by_topic": {ws.cell(row=r, column=2).value.strip(): [raised.get(r, 0), closed.get(r, 0)] for r in sorted(raised)},
                      "unmapped": unmapped, "labels_fixed": fixed}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Where each number of the monthly report comes from (user rule, 2026-10-15: take every figure from its source).

    weekly reports (read_weekly.ps1 -> weekly.json)  manpower and man-hours per day, project-to-date hours, TBT, trainings,
                                                     inspections, HSE meetings held
    the tool (export dir)                            the same for days after the last weekly report; photos; events
    incident register (incidents.register_013)       incidents (date, type, WPS, description)
    consultant SVN / NCR / SOR registers             consultant notices issued and replied in the month
    Engineer's letters (tool, or letters.json)       letters received and replies sent, with the date and subject read
                                                     from each letter (the PDFs are scans, so a file date is not used)

Every value carries its source; anything that has no source is listed under `missing` instead of being estimated.
"""
import calendar, datetime as dt, json, os, re
import openpyxl
from config import CFG

CS, CON, ENG = CFG["contractor_short"], CFG["consultant_short"], CFG.get("pmc", "Engineer")
LC = CFG.get("letters", {})
LTR_REF = re.compile(LC.get("ref_pattern", r"LTR-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{5}"))
DAYS = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


def week_start_of_file(path):
    m = re.findall(r"(\d{1,2})[-\s](\d{1,2})[-\s](\d{2,4})", os.path.basename(path))
    y = int(m[0][2]); y += 2000 if y < 100 else 0
    return dt.date(y, int(m[0][1]), int(m[0][0]))


def month_of_week(start):
    wed = start + dt.timedelta(days=4)             # Saturday + 4 = Wednesday: the week belongs to that month
    return wed.year, wed.month


def cell_date(v):
    """Register dates are a mix of real dates and typed 'dd/mm/yyyy' text."""
    if isinstance(v, dt.datetime):
        return v.date(), "date"
    if isinstance(v, dt.date):
        return v, "date"
    if re.match(r"^\d{4}-\d{2}-\d{2}", str(v or "")):          # the tool's own dates
        return dt.date.fromisoformat(str(v)[:10]), "date"
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})", str(v or ""))
    if m:
        y = int(m[3]); y += 2000 if y < 100 else 0
        try:
            return dt.date(y, int(m[2]), int(m[1])), "text"
        except ValueError:
            return None, ""
    return None, ""


def compute(month, weekly_json, export_load, letters_json=None):
    y, m = map(int, month.split("-"))
    first, last = dt.date(y, m, 1), dt.date(y, m, calendar.monthrange(y, m)[1])
    out = {"sources": {}, "missing": [], "checks": []}
    weekly = json.load(open(weekly_json, encoding="utf-8-sig")) if weekly_json else []

    # --- daily manpower: weekly reports first (dated by the file's week, not the sheet's dates), then the tool
    daily = {}
    for e in weekly:
        s0 = week_start_of_file(e["file"])
        for i, r in enumerate(e.get("dailyManpower") or []):
            day = s0 + dt.timedelta(days=i)
            if r[0] and r[0] != str(day) and first <= day <= last:
                out["checks"].append(f"{os.path.basename(e['file'])}: Daily Manpower row {i + 1} is dated {r[0]}, the report covers {day}; counted as {day}")
            daily[str(day)] = {"manpower": r[2] or 0, "hours": r[3] or 0, "src": os.path.basename(e["file"])}
    covered_to = max(daily) if daily else ""
    tool_mp = {d.get("date") or d["_id"]: d for d in export_load("manpower")}
    for d0, rec in tool_mp.items():
        if d0 > covered_to:
            tot = sum(int(rec.get(f"{g}_{c}") or 0) for g in ("company", "sub", "spec") for c in ("staff", "labor", "saudi"))
            daily[d0] = {"manpower": tot, "hours": tot * float(rec.get("hours") or 11), "src": "tool daily log"}
    days = [str(first + dt.timedelta(days=i)) for i in range((last - first).days + 1)]
    have = [d for d in days if d in daily]
    miss = [d for d in days if d not in daily and d <= str(dt.date.today())]
    if miss:
        out["missing"].append(f"manpower: no weekly report or daily-log record for {len(miss)} day(s): {miss[0]} … {miss[-1]}")
    worked = [d for d in have if daily[d]["manpower"]]
    out["hours"] = round(sum(daily[d]["hours"] for d in have))
    out["avg"] = round(sum(daily[d]["manpower"] for d in worked) / len(worked)) if worked else 0
    out["workingDays"] = len(worked)
    out["sources"]["manpower"] = sorted({daily[d]["src"] for d in have})

    # --- project-to-date: the weekly report holding the month end, less the days after the month end
    ptd = None
    for e in sorted(weekly, key=lambda e: e["file"]):
        s0 = week_start_of_file(e["file"])
        if s0 <= last <= s0 + dt.timedelta(days=6) and e.get("ptdHours") is not None:
            after = sum((r[3] or 0) for i, r in enumerate(e.get("dailyManpower") or []) if s0 + dt.timedelta(days=i) > last)
            ptd = round(e["ptdHours"] - after)
            out["sources"]["toDate"] = f"{os.path.basename(e['file'])} (to date {e['ptdHours']:,.0f} less {after:,.0f} after {last})"
    if ptd is None and weekly:
        lastw = max(weekly, key=lambda e: week_start_of_file(e["file"]))
        s0 = week_start_of_file(lastw["file"])
        end_w = s0 + dt.timedelta(days=6)
        if end_w < last and lastw.get("ptdHours") is not None:
            extra = [d for d in days if d > str(end_w)]
            gap = [d for d in extra if d not in daily]
            ptd = round(lastw["ptdHours"] + sum(daily[d]["hours"] for d in extra if d in daily))
            out["sources"]["toDate"] = f"{os.path.basename(lastw['file'])} + daily log {extra[0]}…{extra[-1]}" if extra else os.path.basename(lastw["file"])
            if gap:
                out["missing"].append(f"to-date hours: {len(gap)} day(s) after the last weekly report have no manpower record")
    out["toDate"] = ptd

    # --- weeks of the month (Wednesday rule, as the weekly and trend reports)
    weeks = [e for e in weekly if month_of_week(week_start_of_file(e["file"])) == (y, m)]
    out["sources"]["weeks"] = [os.path.basename(e["file"]) for e in sorted(weeks, key=lambda e: e["file"])]
    tbt, trainings, inspections = [], [], []
    meetings = 0
    for e in sorted(weeks, key=lambda e: e["file"]):
        for r in e.get("tbt") or []:
            tbt.append({"topic": r[1], "by": r[2], "attendees": r[3], "src": os.path.basename(e["file"])})
        for r in e.get("trainings") or []:
            trainings.append({"topic": r[0], "scope": r[1], "attendees": re.sub(r"\D.*$", "", str(r[2]).strip()) or r[2], "src": os.path.basename(e["file"])})
        for r in e.get("inspections") or []:
            inspections.append({"topic": r[1], "date": r[2], "type": r[3]})
        meetings += int((e.get("rows") or {}).get("47") or 0)
    last_week_end = max((week_start_of_file(e["file"]) + dt.timedelta(days=6) for e in weeks), default=first - dt.timedelta(days=1))
    logs = [x for x in export_load("dailyLog") if x.get("date") and str(last_week_end) < x["date"] <= str(last)]
    for x in logs:
        if x.get("kind") == "tbt":
            tbt.append({"topic": x.get("topic", ""), "by": x.get("by", ""), "attendees": x.get("attendees", ""), "src": "tool daily log"})
        elif x.get("kind") == "training":
            trainings.append({"topic": x.get("topic", ""), "scope": "External" if x.get("scope") == "external" else "Internal", "attendees": x.get("attendees", ""), "src": "tool daily log"})
        elif x.get("kind") == "inspection":
            inspections.append({"topic": x.get("topic", ""), "date": x["date"], "type": "Internal"})
        elif x.get("kind") == "meeting":
            meetings += 1
    out.update(tbt=tbt, trainings=trainings, inspections=inspections, meetings=meetings)
    if str(last_week_end) < str(last) and not logs:
        out["missing"].append(f"TBT/training/meetings after {last_week_end}: no weekly report yet and nothing in the daily log")

    # --- incidents from the incident register
    inc = []
    try:
        ws = openpyxl.load_workbook(CFG["incidents"]["register_013"], data_only=True).worksheets[0]
        for r in ws.iter_rows(min_row=3, values_only=True):
            d0, kind = cell_date(r[1])
            if d0 and first <= d0 <= last:
                inc.append({"date": str(d0), "type": str(r[2] or ""), "wps": re.sub(r"\D", "", str(r[3] or "")), "description": str(r[8] or "").strip(), "company": CS})
            elif kind == "date" and d0 and d0.day <= 12:
                swapped = dt.date(d0.year, d0.day, d0.month)
                if first <= swapped <= last:
                    out["checks"].append(f"incident register row {r[0]}: date stored as {d0} (could be {swapped}, dd/mm typed into a date cell); not counted")
        out["sources"]["incidents"] = os.path.basename(CFG["incidents"]["register_013"])
    except Exception as ex:
        out["missing"].append(f"incident register not read: {ex}")
    out["incidents"] = inc

    # --- consultant notices from the consultant's SVN / NCR / SOR registers
    letters = []
    for name in ("SVN", "NCR", "SOR"):
        path = CFG.get(name.lower(), {}).get("register")
        if not path:
            out["missing"].append(f"{name} register not set in project.json")
            continue
        try:
            ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        except Exception as ex:
            out["missing"].append(f"{name} register not read: {ex}")
            continue
        hdr_row = next(i for i, r in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1) if r and str(r[0] or "").strip().upper().startswith(name))
        hdr = [str(c or "").strip().lower() for c in next(ws.iter_rows(min_row=hdr_row, max_row=hdr_row, values_only=True))]
        col = lambda *keys: next((i for i, h in enumerate(hdr) if any(k in h for k in keys)), None)
        ci, cs, cr, cst = col("date issued", "issued"), col("subject", "title"), col("reply date"), col("status")
        for r in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            if not r or not r[0]:
                continue
            di = cell_date(r[ci])[0] if ci is not None else None
            dr = cell_date(r[cr])[0] if cr is not None else None
            st = str(r[cst] or "") if cst is not None else ""
            subj = (str(r[cs] or "") if cs is not None else "") or ("Safety Observation Report" if name == "SOR" else "")
            state = "Closed" if st == "Closed" else "In Progress" if st.lower().startswith("reply") else "Open"     # the register's status today
            if di and first <= di <= last:
                letters.append([str(r[0]), subj, CON, CS, di.strftime("%d %b %Y"), state])
            if dr and first <= dr <= last:
                letters.append([f"Reply to {r[0]}", subj, CS, CON, dr.strftime("%d %b %Y"), state])
        out["sources"].setdefault("letters", []).append(os.path.basename(path))
    # Engineer's letters: the tool's `consultant` documents of kind LTR (entered after reading each letter), never by file date;
    # letters.json only for months before the letters were in the tool
    ltrs = [n for n in export_load("consultant") if str(n.get("kind", "")).upper() == "LTR" and not n.get("example")]
    known = {str(n.get("ref", "")) for n in ltrs}
    for n in ltrs:
        di, dr = cell_date(n.get("issued"))[0], cell_date(n.get("replyDate"))[0]
        st = {"closed": "Closed", "replied": "In Progress"}.get(n.get("status"), "Open")
        if di and first <= di <= last:
            letters.append([n.get("ref", ""), n.get("subject", ""), n.get("issuedBy", ENG), CS, di.strftime("%d %b %Y"), st])
        if dr and first <= dr <= last:
            letters.append([n.get("replyRef") or f"Reply to {n.get('ref', '')}", n.get("subject", ""), CS, n.get("issuedBy", ENG), dr.strftime("%d %b %Y"), st])
    if ltrs:
        out["sources"].setdefault("letters", []).append(f"{ENG} letters in the tool (consultant, kind LTR)")
    if ltrs:
        base = LC.get("folder", "")
        new = []
        for dp, _, fs in os.walk(base):
            for f in fs:
                m = LTR_REF.search(f)
                if m and f.lower().endswith(".pdf") and not any(m.group(0) in k for k in known) and \
                        (first - dt.timedelta(days=10)) <= dt.date.fromtimestamp(os.path.getmtime(os.path.join(dp, f))) <= (last + dt.timedelta(days=10)):
                    new.append(os.path.join(dp, f))
        if new:
            out["missing"].append(f"{ENG} letter files not in the tool (read each; the file name may not be the letter's own number): " + "; ".join(new))
    elif letters_json and os.path.exists(letters_json):
        for L in json.load(open(letters_json, encoding="utf-8-sig")):
            d0 = dt.date.fromisoformat(L["date"])
            if first <= d0 <= last:
                letters.append([L["ref"], L.get("subject", ""), L.get("from", ENG), L.get("to", CS), d0.strftime("%d %b %Y"), L.get("status", "In Progress")])
        out["sources"].setdefault("letters", []).append(f"{ENG} letters read one by one (letters.json)")
    else:
        cands = []
        base = LC.get("folder", "")
        for dp, _, fs in os.walk(base):
            for f in fs:
                if re.search(LC.get("file_words", "LTR"), f) and (first - dt.timedelta(days=10)) <= dt.date.fromtimestamp(os.path.getmtime(os.path.join(dp, f))) <= (last + dt.timedelta(days=10)):
                    cands.append(os.path.join(dp, f))
        if cands:
            out["missing"].append(f"{ENG} letters: read these and write letters.json (ref, date, subject, from, to, status) before making the report: " + "; ".join(cands))
    out["letters"] = letters
    return out

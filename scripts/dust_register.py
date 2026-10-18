# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Daily dust suppression log REG-017 from the daily log's water-spraying entries (kind "dust").

    python dust_register.py <export dir> --month YYYY-MM [--out rows.json]

<export dir> holds the `dailyLog` collection exported with ArtifactData. One row per day of the month in the register's
month sheet: Time = the spraying times, Remarks = area, tanker trips, water and who, one line per round. The rows go to
dust_register.ps1, which writes them through Excel (makes the month sheet from the previous one when it is missing,
adds a missing day row, never overwrites a day that already has an entry, and backs the register up first).
"""
import calendar, datetime as dt, json, os, sys
from collections import defaultdict
from weekly_hse import load


def ampm(t):
    try:
        return dt.datetime.strptime(t, "%H:%M").strftime("%I:%M %p")
    except ValueError:
        return t


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:]
    export = a[0]
    opt = dict(zip(a[1::2], a[2::2]))
    y, m = map(int, opt["--month"].split("-"))
    first, last = dt.date(y, m, 1), dt.date(y, m, calendar.monthrange(y, m)[1])
    days = defaultdict(list)
    for x in load(export, "dailyLog"):
        if x.get("kind") != "dust" or x.get("example"):
            continue
        d = str(x.get("date", ""))
        if not (str(first) <= d <= str(last)):
            continue
        days[d].append(x)
    out = {}
    for d, xs in sorted(days.items()):
        xs.sort(key=lambda x: x.get("time") or "")
        times = [ampm(x.get("time", "")) for x in xs]
        lines = []
        for x in xs:
            bits = [x.get("location", "").strip()]
            if x.get("qty"):
                bits.append(f"{x['qty']} tanker trip" + ("s" if int(x["qty"]) != 1 else ""))
            if x.get("water"):
                bits.append(f"{x['water']} m³")
            s = " – ".join(b for b in bits if b)
            if x.get("by"):
                s += f" ({x['by']})"
            if x.get("photo"):
                s += " – photo"
            lines.append(s)
        out[d] = {"time": "\n".join(times), "remarks": "\n".join(lines)}
    res = {"month": f"{y}-{m:02d}", "sheet": calendar.month_name[m], "label": f"{calendar.month_name[m]} {y}", "days_in_month": last.day, "days": out}
    path = opt.get("--out")
    if path:
        json.dump(res, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    missing = [str(first + dt.timedelta(days=i)) for i in range(last.day) if str(first + dt.timedelta(days=i)) <= str(dt.date.today()) and str(first + dt.timedelta(days=i)) not in out]
    print(json.dumps({"days_with_entries": len(out), "rounds": sum(len(v) for v in days.values()),
                      "days_without_entry_so_far": missing, "rows": path or "(pass --out)"}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

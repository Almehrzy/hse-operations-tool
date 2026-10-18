# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Add or update one row of the consultant SVN register (path in project.json).

    python svn_register.py <svn number> field=value [field=value ...]

Fields: issued, received, location, subject, action, reply_date, rev, status, closed, aconex, folder, remarks
Dates as YYYY-MM-DD. status: Open | Reply submitted | Closed | Not on file.
A new number is added as a new row after the last one, with the same formatting.
Existing values are only replaced when given; remarks are appended, not replaced.
"""
import copy, datetime, io, os, sys
import openpyxl
from config import CFG

REG = os.environ.get("SVN_REGISTER", CFG["svn"]["register"])
PREFIX = CFG["svn"]["number_prefix"]
COLS = {"issued": 2, "received": 3, "location": 4, "subject": 5, "action": 6, "reply_date": 7, "rev": 8,
        "status": 9, "closed": 10, "aconex": 11, "folder": 13, "remarks": 14}
DATES = {"issued", "received", "reply_date", "closed"}
STATUSES = {"Open", "Reply submitted", "Closed", "Not on file"}


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    no = int(sys.argv[1])
    ref = f"{PREFIX}{no:03d}"
    updates = dict(a.split("=", 1) for a in sys.argv[2:])
    bad = set(updates) - set(COLS)
    if bad:
        raise SystemExit(f"Unknown field(s): {', '.join(bad)}")
    if "status" in updates and updates["status"] not in STATUSES:
        raise SystemExit(f"status must be one of {sorted(STATUSES)}")
    wb = openpyxl.load_workbook(REG)
    ws = wb["SVN Register"]
    rows = {ws.cell(row=r, column=1).value: r for r in range(6, ws.max_row + 1) if ws.cell(row=r, column=1).value}
    if ref in rows:
        r = rows[ref]
        action = "updated"
    else:
        last_no = max(int(k[-3:]) for k in rows)
        if no != last_no + 1:
            raise SystemExit(f"{ref} is not in the register and the last number is {last_no:03d}. Add the missing numbers first.")
        last = max(rows.values())
        r = last + 1
        for c in range(1, 15):
            src, dst = ws.cell(row=last, column=c), ws.cell(row=r, column=c)
            dst._style = copy.copy(src._style)
        ws.cell(row=r, column=1, value=ref)
        ws.cell(row=r, column=12, value=f'=IF(OR(I{r}="Closed",I{r}="Not on file",B{r}=""),"",TODAY()-B{r})')
        ws.row_dimensions[r].height = ws.row_dimensions[last].height
        action = "added"
    for k, v in updates.items():
        c = ws.cell(row=r, column=COLS[k])
        if k in DATES:
            c.value = datetime.date.fromisoformat(v) if v else None
            c.number_format = "DD-MM-YYYY"
            c.font = copy.copy(ws.cell(row=r, column=5).font)      # a real date, not the grey "file date" style
        elif k == "remarks":
            c.value = (str(c.value) + "; " if c.value else "") + v
        elif k == "folder":
            c.value = os.path.basename(v.rstrip("\\/"))
            c.hyperlink = v
        else:
            c.value = v
    # keep the summary formulas, filter and formatting ranges covering the new row
    end = max(r, ws.max_row)
    for col, letter in ((2, "B"), (5, "E"), (8, "H"), (11, "K")):
        f = ws.cell(row=3, column=col).value
        if isinstance(f, str) and f.startswith("=COUNTIF"):
            ws.cell(row=3, column=col).value = f.split("(")[0] + f"(I6:I{end}," + f.split(",", 1)[1]
    ws.auto_filter.ref = f"A5:N{end}"
    wb.save(REG)
    print(f"{action} {ref} (row {r}): " + ", ".join(f"{k}={v}" for k, v in updates.items()))


if __name__ == "__main__":
    main()

# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Equipment tab: checklist settings, equipment list from the TUV register, and filled checklists on the house forms.

    python equipment.py settings <out.json>          settings/equipment document for the page (checklists per family, colour code)
    python equipment.py seed <out_dir>               one JSON per equipment on site, from the TUV register
    python equipment.py monthly <request.json> <out.xlsx>   request: {"inspection": doc, "equipment": doc, "photos": [local jpg]}
    python equipment.py daily <request.json> <out.xlsx>     request: {"equipment": doc, "month": "YYYY-MM", "inspections": [daily docs]}
    python equipment.py log <request.json> <out.xlsx> [--replace]   request: {"month": "YYYY-MM", "rows": [inspection docs + plateSticker, company, file]}
    python equipment.py regplan <equipment export dir> <changes.json>   then: scripts\tuv_register.ps1 -ChangesJson changes.json

Project setup lives in equipment.json (forms, families, Arabic item names, colour code); see its "_about".
Checklist items keep the forms' own English wording; families without a monthly sheet reuse their daily item list.
"""
import datetime as dt, json, os, re, sys
import openpyxl
from config import CFG, ROOT

EQ = json.load(open(os.path.join(ROOT, "equipment.json"), encoding="utf-8"))
NON_LATIN = re.compile(r"[؀-ۿऀ-ॿ]")


def path(p):
    return p if os.path.isabs(p) else os.path.join(ROOT, p)


def english(s):
    """The English part of a form item ('OIL LEVEL तेल का…' -> 'OIL LEVEL'), whitespace tidied."""
    s = str(s)
    if "\n" in s and NON_LATIN.search(s.split("\n", 1)[1]):     # translation on its own line
        s = s.split("\n", 1)[0]
    s = re.sub(r"\s+", " ", s).strip()
    m = NON_LATIN.search(s)
    return (s[:m.start()] if m else s).strip(" ,.")


def form_items(ws):
    """[(row, text)] of the numbered item rows of a checklist sheet (No. in column A, item in column B)."""
    out, started = [], False
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if not started:
            started = str(a or "").strip().upper().startswith(("NO", "SL NO", "S NO"))
            continue
        if b is None or not str(b).strip():
            if out and a is None:
                break
            continue
        if str(b).strip().upper().startswith(("SIGNATURE", "OPERATOR'S", "INSPECTOR")):
            break
        if re.match(r"^\d+$", str(a or "").strip()):
            out.append((r, english(b)))
    return out


_books = {}


def sheet(name):
    """A form sheet by name: the monthly form first, then the daily one."""
    for key in ("monthly_form", "daily_form"):
        p = path(EQ[key])
        if p not in _books:
            _books[p] = openpyxl.load_workbook(p)
        if name in _books[p].sheetnames:
            return _books[p][name], p
    raise SystemExit(f"Sheet '{name}' is in neither form")


def title_of(ws):
    for r in range(1, 4):
        v = ws.cell(r, 1).value
        if v and str(v).strip():
            return english(str(v).replace("\n", " "))
    return ""


def items_of(fam, which):
    if fam.get("items"):
        return [{"en": x, "ar": EQ["ar"].get(x, "")} for x in fam["items"]], fam["en"].upper() + " CHECKLIST"
    ws, _ = sheet(fam[which])
    return [{"en": x, "ar": EQ["ar"].get(x, EQ["ar"].get(x.upper(), ""))} for _, x in form_items(ws)], title_of(ws)


def cmd_settings(out):
    fams, missing = [], set()
    for f in EQ["families"]:
        m, mt = items_of(f, "monthly")
        d, dtl = items_of(f, "daily")
        for it in m + d:
            if not it["ar"]:
                missing.add(it["en"])
        fams.append({"key": f["key"], "en": f["en"], "ar": f["ar"], "types": f["types"],
                     "monthly": {"title": mt, "items": m}, "daily": {"title": dtl, "items": d}})
    doc = {"colorCode": EQ["color_code"], "families": fams, "codePrefix": EQ.get("code_prefix", "EQ-"),
           "pageUrl": CFG.get("page_url", "")}   # QR labels carry the page link, so a phone camera opens the tool
    json.dump(doc, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"written": out, "families": [(f["key"], len(f["monthly"]["items"]), len(f["daily"]["items"])) for f in fams],
                      "no_arabic": sorted(missing)}, ensure_ascii=False, indent=1))


def family_of(typ):
    t = re.sub(r"\s+", " ", str(typ or "")).strip().upper()
    for f in EQ["families"]:
        if t in [x.upper() for x in f["types"]]:
            return f["key"]
    for f in EQ["families"]:                       # "BAR BENDING MACHINE" etc. fall through to general
        if any(x.upper() in t for x in f["types"]):
            return f["key"]
    return "general"


def as_date(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v or "").strip()
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$", s)  # 19-12-26 / 19/12/2026 (day first, as typed in the register)
    if m:
        y = int(m[3]) + (2000 if len(m[3]) == 2 else 0)
        try:
            return dt.date(y, int(m[2]), int(m[1])).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def clean(v):
    s = re.sub(r"\s+", " ", str(v if v is not None else "")).strip()
    return "" if s.upper() in ("N/A", "-", "NONE") else s


def cmd_seed(out_dir):
    wb = openpyxl.load_workbook(path(EQ["register"]), data_only=True)
    ws = wb[EQ["register_sheet"]]
    os.makedirs(out_dir, exist_ok=True)
    pre, n, skipped, fams = EQ.get("code_prefix", "EQ-"), 0, 0, {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        no, typ = row[0], row[1]
        if not typ or not isinstance(no, (int, float)):
            continue
        if "OFF" in str(row[11] or "").upper():
            skipped += 1
            continue
        code = f"{pre}{int(no):03d}"
        doc = {"code": code, "regNo": int(no), "type": clean(typ).upper(), "family": family_of(typ),
               "serial": clean(row[2]), "sticker": clean(row[3]), "make": clean(row[4]), "plate": clean(row[5]),
               "certCo": clean(row[6]), "certNo": clean(row[7]), "certDate": as_date(row[8]), "certExpiry": as_date(row[9]),
               "company": clean(row[10]), "status": "unverified", "source": "register",
               "createdAt": dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")}
        json.dump(doc, open(os.path.join(out_dir, code.lower() + ".json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        n += 1
        fams[doc["family"]] = fams.get(doc["family"], 0) + 1
    print(json.dumps({"written": n, "off_site_skipped": skipped, "by_family": fams}, indent=1))


# ---------- filled checklists ----------

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]


def dmy(s):
    s = str(s or "")
    return f"{s[8:10]}-{s[5:7]}-{s[:4]}" if re.match(r"^\d{4}-\d{2}-\d{2}", s) else s


def family(key):
    for f in EQ["families"]:
        if f["key"] == key:
            return f
    return next(f for f in EQ["families"] if f["key"] == "general")


def keep_only(wb, name):
    for n in list(wb.sheetnames):
        if n != name:
            del wb[n]
    return wb[name]


def set_label(ws, pattern, value, last_row=None):
    """Write value after a form label such as 'Equipment Type: ______'. Returns True when the label was found."""
    for row in ws.iter_rows(max_row=last_row or ws.max_row):
        for c in row:
            if isinstance(c.value, str) and re.search(pattern, c.value, re.I):
                base = re.split(r"_{2,}", c.value)[0].rstrip()
                base = base if base.endswith((":", "-")) else base + ":"
                c.value = f"{base} {value}" if value else base
                return True
    return False


def out_path(out):
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    return out


def wrap_down(ws, c, rows=3, last_col=6):
    """Let a long text cell wrap over the free cells to its right (to column F, the form's width) and below,
    merging them when none of them is merged or filled already."""
    from openpyxl.styles import Alignment
    rng = f"{c.coordinate}:{openpyxl.utils.get_column_letter(last_col)}{c.row + rows - 1}"
    cells = [x for r in ws[rng] for x in r]
    if not any(x.coordinate in m for x in cells for m in ws.merged_cells.ranges) and all(x.value is None for x in cells[1:]):
        ws.merge_cells(rng)
    c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")


def add_photos(wb, photos, captions):
    """One photo per block on a 'Photos' sheet that prints on A4."""
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PImage
    if not photos:
        return
    ws = wb.create_sheet("Photos")
    ws.column_dimensions["A"].width = 95
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0
    row = 1
    for i, p in enumerate(photos):
        w, h = PImage.open(p).size
        s = min(1.0, 600 / w, 400 / h)
        img = XLImage(p)
        img.width, img.height = int(w * s), int(h * s)
        ws[f"A{row}"] = captions[i] if i < len(captions) else f"Photo {i + 1}"
        ws.add_image(img, f"A{row + 1}")
        ws.row_dimensions[row + 1].height = img.height * 0.75 + 6
        row += 3


def cmd_monthly(req_json, out):
    """Monthly inspection on the family's own sheet of the monthly form (or a same-layout sheet when the form has none)."""
    req = json.load(open(req_json, encoding="utf-8-sig"))
    ins, eq = req["inspection"], req["equipment"]
    fam = family(eq.get("family"))
    items = ins["items"]
    wb = openpyxl.load_workbook(path(EQ["monthly_form"]))
    own = fam.get("monthly") in wb.sheetnames
    if own:
        ws = keep_only(wb, fam["monthly"])
    else:                       # e.g. welding machine: reuse a monthly sheet with enough item rows
        base = next(n for n in ("Air Compressor", "Const equipment") if len(form_items(wb[n])) >= len(items))
        ws = keep_only(wb, base)
        ws.title = fam["en"][:31].replace("/", "-")
    rows = form_items(ws)
    if not own:
        title_cell = next(ws.cell(r, 1) for r in range(1, 4) if ws.cell(r, 1).value)
        title_cell.value = re.sub(r"\s*DAILY\s*", " ", items_of(fam, "monthly")[1]).strip() or fam["en"].upper() + " INSPECTION CHECKLIST"
        for i, (r, _) in enumerate(rows):
            ws.cell(r, 2).value = items[i]["en"] if i < len(items) else None
            if i >= len(items):
                ws.cell(r, 1).value = None
    else:
        names = [english(x) for _, x in rows]
        mine = [it["en"] for it in items]
        if names[:len(mine)] != mine:
            raise SystemExit(f"Checklist items differ from the form sheet '{fam['monthly']}'. Rebuild settings (equipment.py settings) and check.")
    # header columns
    hdr = next(r for r in range(1, 15) if str(ws.cell(r, 1).value or "").strip().upper().startswith("NO"))
    col = {}
    for c in range(3, 8):
        v = str(ws.cell(hdr, c).value or "").strip().upper()
        if v.startswith("YES"): col["ok"] = c
        elif v.startswith("NO"): col["no"] = c
        elif v.startswith("N/A"): col["na"] = c
        elif v.startswith("COMMENT"): col["note"] = c
        elif v.startswith("CONDITION"): col["cond"] = c
    col.setdefault("note", 6)
    for i, (r, _) in enumerate(rows[:len(items)]):
        it = items[i]
        if "cond" in col:
            ws.cell(r, col["cond"]).value = {"ok": "OK", "no": "NOT OK", "na": "N/A"}[it["r"]]
        else:
            ws.cell(r, col[it["r"]]).value = "✓"
        if it.get("note"):
            ws.cell(r, col["note"]).value = it["note"]
    last = rows[-1][0]
    set_label(ws, r"EQUIPMENT MANUFACTURE", eq.get("make", ""))
    set_label(ws, r"^\s*Equipment Type", eq.get("type", ""))
    set_label(ws, r"^\s*Equipment No", f"{eq.get('code', '')}  (Sticker {eq['sticker']})" if eq.get("sticker") else eq.get("code", ""))
    set_label(ws, r"^\s*Date\s+of inspection", dmy(ins.get("date")))
    set_label(ws, r"^\s*Registration No", eq.get("plate", "") or eq.get("serial", ""))
    set_label(ws, r"^\s*Equipment T\w+ Validity", dmy(eq.get("certExpiry")))
    set_label(ws, r"^\s*Equipment T\w+ Cert", eq.get("certNo", ""))
    set_label(ws, r"^\s*Operator'?s Name", ins.get("operator", ""))
    set_label(ws, r"^\s*Operator T\w+ Cert", ins.get("operatorCert", ""))
    set_label(ws, r"^\s*Operator T\w+ Validity", dmy(ins.get("operatorCertExpiry")))
    set_label(ws, r"^\s*Inspector.s Name", ins.get("reporter", ""))
    rec = [ins.get("notes", "")] + [f"{it['en']}: {it['note']}" for it in items if it["r"] == "no" and it.get("note")]
    if ins.get("result") == "stopped":
        rec.append("Equipment stopped until repaired.")
    if ins.get("color"):
        rec.append(f"Inspection colour code: {ins['color']}.")
    for row in ws.iter_rows(min_row=last + 1):
        for c in row:
            if isinstance(c.value, str) and re.search(r"HSSE Recommendations", c.value, re.I):
                c.value = "HSSE Recommendations - " + " ".join(x for x in rec if x)
                wrap_down(ws, c)
            elif isinstance(c.value, str) and re.match(r"^\s*Signature", c.value, re.I):
                c.value = f"Inspected by: {ins.get('reporter', '')}    Date: {dmy(ins.get('date'))}    Signature:"
    add_photos(wb, req.get("photos", []), [f"{eq.get('code', '')} – photo {i + 1}, {dmy(ins.get('date'))}" for i in range(len(req.get('photos', [])))])
    wb.save(out_path(out))
    print(json.dumps({"written": out, "sheet": ws.title, "own_sheet": own, "items": len(items),
                      "defects": sum(1 for it in items if it["r"] == "no")}, ensure_ascii=False))


def cmd_daily(req_json, out):
    """A month of daily checks for one piece of equipment on the family's daily sheet (one column per day)."""
    from openpyxl.styles import Alignment, Font
    req = json.load(open(req_json, encoding="utf-8-sig"))
    eq, month, checks = req["equipment"], req["month"], sorted(req["inspections"], key=lambda x: x.get("createdAt", ""))
    fam = family(eq.get("family"))
    wb = openpyxl.load_workbook(path(EQ["daily_form"]))
    name = fam.get("daily") if fam.get("daily") in wb.sheetnames else "C.EQUIP NEW"
    ws = keep_only(wb, name)
    rows = form_items(ws)
    by_name = {english(x).upper(): r for r, x in rows}
    y, m = int(month[:4]), int(month[5:7])
    set_label(ws, r"EQUIPE?MENT NO", f"{eq.get('code', '')} {eq.get('type', '')} {eq.get('plate', '') or eq.get('sticker', '')}".strip())
    set_label(ws, r"^\s*MONTH", f"{MONTHS[m - 1]} {y}")
    remarks, missing = [], set()
    for x in checks:
        day = int(x["date"][8:10])
        c = 2 + day
        for it in x.get("items", []):
            r = by_name.get(it["en"].upper())
            if r is None:
                missing.add(it["en"])
                continue
            cell = ws.cell(r, c)
            cell.value = {"ok": "✓", "no": "✗", "na": "NA"}[it["r"]]
            cell.alignment = Alignment(horizontal="center", vertical="center")
        bad = [f"{it['en']}: {it.get('note', '')}" for it in x.get("items", []) if it["r"] == "no"]
        who = ", ".join(v for v in [x.get("reporter") and "checked by " + x["reporter"], x.get("operator") and "operator " + x["operator"]] if v)
        remarks.append(f"{dmy(x['date'])}: {who or 'checked'}" + (f"; defects – {'; '.join(bad)}" if bad else "; all OK")
                       + ("; equipment stopped" if x.get("result") == "stopped" else ""))
    r0 = ws.max_row + 2
    ws.cell(r0, 1).value = "REMARKS"
    ws.cell(r0, 1).font = Font(bold=True)
    for i, t in enumerate(remarks):
        ws.cell(r0 + 1 + i, 1).value = t
        ws.cell(r0 + 1 + i, 1).alignment = Alignment(horizontal="left")
    ws.print_area = f"A1:{openpyxl.utils.get_column_letter(ws.max_column)}{r0 + len(remarks)}"
    wb.save(out_path(out))
    print(json.dumps({"written": out, "sheet": name, "days": len(checks), "items_not_on_form": sorted(missing)}, ensure_ascii=False))


def cmd_log(req_json, out, *flags):
    """Monthly inspection log for the month folder, rebuilt from the tool's records (--replace allowed: it is derived data)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    req = json.load(open(req_json, encoding="utf-8-sig"))
    if os.path.exists(out) and "--replace" not in flags:
        raise SystemExit(f"Refusing to overwrite {out} (pass --replace to rebuild it)")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inspections"
    ws["A1"] = f"{CFG.get('project_name', '')} – Equipment inspection log {req['month']}"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"{CFG.get('contractor', '')}  ·  generated from the site tool on {dt.date.today():%d-%m-%Y}"
    head = ["No", "Date", "Kind", "Equipment", "Type", "Plate / sticker", "Company", "Location", "Result", "Defects", "Colour", "Inspected by", "Observation", "File"]
    thin = Side(style="thin", color="999999")
    for i, hd in enumerate(head, 1):
        c = ws.cell(4, i, hd)
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="17201C")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for n, x in enumerate(sorted(req["rows"], key=lambda x: (x.get("date", ""), x.get("code", ""))), 1):
        bad = "; ".join(f"{it['en']}: {it.get('note', '')}" for it in x.get("items", []) if it["r"] == "no")
        vals = [n, dmy(x.get("date")), x.get("kind", ""), x.get("code", ""), x.get("type", ""), x.get("plateSticker", ""), x.get("company", ""),
                x.get("location", ""), {"ok": "OK", "defects": "Defects", "stopped": "Stopped"}.get(x.get("result"), ""), bad, x.get("color", ""),
                x.get("reporter", ""), x.get("obsNo", ""), x.get("file", "")]
        for i, v in enumerate(vals, 1):
            c = ws.cell(4 + n, i, v)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    for i, w in enumerate([5, 11, 9, 9, 22, 18, 14, 18, 9, 40, 9, 16, 11, 40], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    wb.save(out)
    print(json.dumps({"written": out, "rows": len(req["rows"])}))


def cmd_regplan(eq_dir, out):
    """Register changes for equipment documents marked regPending (export of the `equipment` collection), for tuv_register.ps1."""
    wb = openpyxl.load_workbook(path(EQ["register"]), data_only=True)
    ws = wb[EQ["register_sheet"]]
    reg = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        if isinstance(row[0], (int, float)):
            reg[int(row[0])] = row
    used = [n for n, r in reg.items() if r[1]]
    last = max(used) if used else 0                   # the register has pre-numbered empty rows after the last entry
    changes, report = [], []
    docs = [json.load(open(os.path.join(dp, f), encoding="utf-8-sig")) for dp, _, fs in os.walk(eq_dir) for f in fs if f.endswith(".json")]
    for d in sorted(docs, key=lambda d: d.get("code", "")):
        d = d.get("data", d)
        if not d.get("regPending") or d.get("example"):
            continue
        no = d.get("regNo") or int(re.sub(r"\D", "", d.get("code", "")) or 0)
        new = no not in reg or not reg[no][1]          # no row yet, or a pre-numbered empty row to fill
        want = {"D": d.get("sticker", ""), "F": d.get("plate", ""), "G": d.get("certCo", ""), "H": d.get("certNo", ""),
                "I": d.get("certDate", ""), "J": d.get("certExpiry", ""), "L": "OFF SITE" if d.get("status") == "offsite" else ""}
        if new:
            want.update({"B": d.get("type", ""), "C": d.get("serial", ""), "E": d.get("make", ""), "K": d.get("company", "")})
        else:
            old = reg[no]
            colidx = {"B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11}
            def same(k, v):
                o = old[colidx[k]]
                return (as_date(o) if k in "IJ" else clean(o)) == v
            want = {k: v for k, v in want.items() if not same(k, v) and not (v == "" and k != "L")}
            if old[11] and "OFF" not in str(old[11]).upper() and want.get("L") == "":
                want.pop("L")                       # keep other remarks
        if new:
            want = {k: v for k, v in want.items() if v}
        if want:
            changes.append({"no": no, "new": new, "fill": new and no in reg, "code": d.get("code"), "set": want})
        report.append({"code": d.get("code"), "no": no, "new": new, "changes": want})
    nums = sorted(c["no"] for c in changes if c["new"])
    if nums and nums != list(range(last + 1, last + 1 + len(nums))):
        report.append({"warning": f"new equipment numbers {nums} do not follow the register's last No. {last}"})
    json.dump(changes, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"written": out, "register_last_no": last, "plan": report}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    {"settings": cmd_settings, "seed": cmd_seed, "monthly": cmd_monthly, "daily": cmd_daily, "log": cmd_log,
     "regplan": cmd_regplan}[sys.argv[1]](*sys.argv[2:])

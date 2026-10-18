# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""SOR close-out workbook in the house format (template made from the SOR-019 close-out).

    python sor_reply.py template <existing closeout.xlsx> <template.xlsx>   -> make a clean template (once)
    python sor_reply.py make <reply.json> <out.xlsx>                        -> write the close-out

reply.json:
{
 "sor_no": 24, "ref": "PACE-WP03-HSE-SOR-024", "inspection_date": "2026-10-20", "received": "2026-10-20",
 "items": [{"location": "...", "category": "...", "risk": "E", "requirement": "observation and recommendation text",
            "observation_photo": "C:/.../pace_p1_1.jpg", "rectification_photo": "C:/.../fix.jpg",
            "corrective_action": "...", "closure_date": "2026-10-20"}]
}
risk: E, H, M or L. Photos are optional per item but every item needs a corrective action and closure date.
"""
import copy, datetime as dt, hashlib, io, json, os, re, sys
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from config import CFG, path as cfg_path

RISK = {"E": "E - Immediately", "H": "H - 24 Hours", "M": "M - 48 Hours", "L": "L - 72 Hours"}
RISK_HOURS = {"E": 0, "H": 24, "M": 48, "L": 72}
FIRST = 9            # first observation row in the template
PLACEHOLDER = re.compile(r"\binsert\b|\bTBD\b|\bxxx+\b|\[.*?\]|\bif available\b|lorem", re.I)


def template(src, dst):
    if os.path.exists(dst):
        raise SystemExit(f"Refusing to overwrite {dst}")
    wb = openpyxl.load_workbook(src)
    ws = wb.worksheets[0]
    ws._images = []
    for r in range(FIRST, ws.max_row + 1):
        for c in range(1, 11):
            ws.cell(row=r, column=c).value = None
    for r in range(FIRST + 1, ws.max_row + 1):
        ws.row_dimensions[r].height = None
    ws.delete_rows(FIRST + 1, ws.max_row - FIRST)
    ws["F3"] = None
    ws["F4"] = None
    ws.title = "SOR"
    wb.save(dst)
    return {"template": dst}


def check(req):
    problems = []
    seen = {}
    insp = dt.date.fromisoformat(req["inspection_date"])
    for i, it in enumerate(req["items"], 1):
        if it.get("risk", "").upper()[:1] not in RISK:
            problems.append(f"item {i}: risk must be E, H, M or L")
        if not (it.get("corrective_action") or "").strip():
            problems.append(f"item {i}: corrective action is empty")
        if not it.get("closure_date"):
            problems.append(f"item {i}: closure date is empty")
        elif dt.date.fromisoformat(it["closure_date"]) < insp:
            problems.append(f"item {i}: closure date is before the inspection date")
        for k in ("requirement", "corrective_action"):
            if PLACEHOLDER.search(it.get(k) or ""):
                problems.append(f"item {i}: placeholder-like text in {k}")
        for k in ("observation_photo", "rectification_photo"):
            p = it.get(k)
            if p and not os.path.exists(p):
                problems.append(f"item {i}: {k} not found: {p}")
        p = it.get("rectification_photo")
        if p and os.path.exists(p):
            h = hashlib.sha256(open(p, "rb").read()).hexdigest()
            if h in seen:
                problems.append(f"item {i}: same rectification photo as item {seen[h]}")
            seen.setdefault(h, i)
    return problems


def late_items(req):
    """Items closed after their risk-level deadline (a warning, not an error)."""
    base = dt.date.fromisoformat(req.get("received") or req["inspection_date"])
    out = []
    for i, it in enumerate(req["items"], 1):
        r = it.get("risk", "").upper()[:1]
        if r in RISK_HOURS and it.get("closure_date"):
            due = base + dt.timedelta(hours=RISK_HOURS[r])
            if dt.date.fromisoformat(it["closure_date"]) > due:
                out.append(f"item {i} ({r}) closed {it['closure_date']}, due {due}")
    return out


def picture(path, max_w, max_h):
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    im.thumbnail((1400, 1400))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    buf.seek(0)
    img = XLImage(buf)
    s = min(max_w / img.width, max_h / img.height)
    img.width, img.height = int(img.width * s), int(img.height * s)
    return img


def make(req, out):
    problems = check(req)
    if problems and not req.get("allow_warnings"):
        raise SystemExit("Not written. Fix these first:\n- " + "\n- ".join(problems))
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    tpl = cfg_path(CFG["sor"]["template"])
    wb = openpyxl.load_workbook(tpl)
    ws = wb.worksheets[0]
    ws["F3"] = req["sor_no"]
    ws["F4"] = dt.date.fromisoformat(req["inspection_date"])
    ws["F4"].number_format = "DD-MM-YYYY"
    styles = [copy.copy(ws.cell(row=FIRST, column=c)._style) for c in range(1, 11)]
    for i, it in enumerate(req["items"]):
        r = FIRST + i
        for c in range(1, 11):
            ws.cell(row=r, column=c)._style = copy.copy(styles[c - 1])
        risk = it["risk"].upper()[:1]
        vals = {1: i + 1, 2: req["ref"], 3: it.get("location", ""), 4: it.get("category", ""), 6: RISK[risk],
                7: it.get("requirement", ""), 9: it.get("corrective_action", ""), 10: dt.date.fromisoformat(it["closure_date"])}
        for c, v in vals.items():
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=10).number_format = "DD-MM-YYYY"
        heights = [120]
        if it.get("observation_photo"):
            img = picture(it["observation_photo"], 210, 260)
            ws.add_image(img, f"E{r}")
            heights.append(img.height)
        if it.get("rectification_photo"):
            img = picture(it["rectification_photo"], 210, 260)
            ws.add_image(img, f"H{r}")
            heights.append(img.height)
        ws.row_dimensions[r].height = max(heights) * 0.75 + 8
    last = FIRST + len(req["items"]) - 1
    ws.print_area = f"A1:J{last}"
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    wb.save(out)
    return {"written": out, "items": len(req["items"]), "late_items": late_items(req), "warnings": problems}


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if len(sys.argv) < 4 or sys.argv[1] not in ("make", "template"):
        raise SystemExit(__doc__)
    if sys.argv[1] == "template":
        res = template(sys.argv[2], sys.argv[3])
    else:
        res = make(json.load(open(sys.argv[2], encoding="utf-8-sig")), sys.argv[3])
    print(json.dumps(res, ensure_ascii=False, indent=1, default=str))


if __name__ == "__main__":
    main()

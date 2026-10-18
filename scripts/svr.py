# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Safety Violation Report (SVR) helper.

    python svr.py plan  <request.json>             -> next number, month folder, prior violations, penalty matrix row
    python svr.py make  <request.json> <out.docx>  -> fill the blank SVR form and append the photos

The blank form, the register and the penalty matrix are only read here. Writing to the
register is done separately by register_append.ps1.

request.json:
  {"no": 214, "date": "2026-10-14", "location": "...", "company": "...", "reported_by": "...",
   "persons": [{"name": "...", "id": "...", "designation": "..."}],
   "details": "...", "item": 13, "penalty": "100 Riyal", "photos": ["C:/.../a.jpg"]}
"""
import datetime, io, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
from config import CFG

SV = CFG["svr"]["folder"]
REGISTER = CFG["svr"]["register"]
MATRIX = CFG["svr"]["penalty_matrix"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
MB = '<w:rFonts w:asciiTheme="majorBidi" w:hAnsiTheme="majorBidi" w:cstheme="majorBidi"/>'
EMU_CM = 360000


def year_dir(year):
    for d in os.listdir(SV):
        if os.path.isdir(os.path.join(SV, d)) and str(year) in d:
            sub = os.listdir(os.path.join(SV, d))
            if any(re.match(r"SVR\s*#", f, re.I) for f in sub) or any(os.path.isdir(os.path.join(SV, d, s)) for s in sub):
                return os.path.join(SV, d)
    return None


def template_path(year):
    yd = year_dir(year)
    for base in filter(None, [yd, year_dir(year - 1), SV]):
        p = os.path.join(base, "SVR # Blank.docx")
        if os.path.exists(p):
            return p
    raise SystemExit("Blank SVR form not found")


def month_folder(date):
    yd = year_dir(date.year)
    name = MONTHS[date.month - 1]
    if yd:
        for d in os.listdir(yd):
            if os.path.isdir(os.path.join(yd, d)) and d.lower().lstrip("0123456789. ").startswith(name[:3].lower()):
                return os.path.join(yd, d), True
        return os.path.join(yd, f"{name} {date.year}"), False
    return os.path.join(SV, str(date.year), f"{name} {date.year}"), False


def file_numbers():
    nums = []
    for d, _, fs in os.walk(SV):
        for f in fs:
            m = re.match(r"SVR\s*#\s*(\d+)", f, re.I)
            if m and f.lower().endswith((".docx", ".doc", ".pdf")) and not f.startswith("~$"):
                nums.append(int(m.group(1)))
    return nums


def norm(s):
    return re.sub(r"[^A-Z ]", " ", str(s or "").upper()).split()


def same_person(a, b):
    ta, tb = set(norm(a)), set(norm(b))
    return len(ta) >= 2 and len(tb) >= 2 and (ta <= tb or tb <= ta)


def read_register():
    import openpyxl
    ws = openpyxl.load_workbook(REGISTER, read_only=True, data_only=True)[CFG["svr"]["register_sheet"]]
    rows = []
    for r in ws.iter_rows(min_row=5, max_col=10, values_only=True):
        if isinstance(r[0], (int, float)) and r[1]:
            rows.append({"svr": int(r[0]), "names": str(r[1]).split("\n"), "ids": re.findall(r"\d{3,}", str(r[2] or "")),
                         "details": str(r[4] or ""), "company": str(r[5] or ""), "penalty": str(r[6] or ""),
                         "date": r[7].strftime("%Y-%m-%d") if hasattr(r[7], "strftime") else str(r[7] or "")})
        elif isinstance(r[0], (int, float)):
            rows.append({"svr": int(r[0]), "names": [], "ids": [], "details": "", "company": "", "penalty": "", "date": ""})
    return rows


def read_matrix(item):
    import openpyxl
    ws = openpyxl.load_workbook(MATRIX, read_only=True, data_only=True).worksheets[0]
    for r in ws.iter_rows(min_row=3, max_col=5, values_only=True):
        if isinstance(r[0], (int, float)) and int(r[0]) == int(item):
            text = re.sub(r"[ \t]+", " ", str(r[1] or "")).strip()
            title, _, desc = text.partition(":")
            clean = lambda v: [re.sub(r"\s+", " ", x).strip() for x in str(v or "").split("\n") if x.strip()]
            return {"item": int(item), "title": title.strip() + ":", "description": re.sub(r"\s+", " ", desc).strip(),
                    "first": clean(r[2]), "second": clean(r[3]), "third": clean(r[4])}
    raise SystemExit(f"Penalty matrix has no item {item}")


def plan(req):
    date = datetime.date.fromisoformat(req["date"])
    reg = read_register()
    reg_last = max((r["svr"] for r in reg), default=0)
    files_last = max(file_numbers(), default=0)
    folder, exists = month_folder(date)
    persons = []
    for p in req.get("persons", []):
        pid = re.sub(r"\D", "", str(p.get("id") or ""))
        prior = []
        for r in reg:
            how = "id" if pid and pid in r["ids"] else "name" if any(same_person(p.get("name"), n) for n in r["names"]) else None
            if how:
                prior.append({"svr": r["svr"], "date": r["date"], "details": r["details"], "penalty": r["penalty"], "company": r["company"], "matched_by": how})
        persons.append({"name": p.get("name"), "id": p.get("id"), "prior": prior})
    out = {"next_no": max(reg_last, files_last) + 1, "register_last": reg_last, "files_last": files_last,
           "month_folder": folder, "month_folder_exists": exists, "template": template_path(date.year),
           "persons": persons, "matrix": read_matrix(req["item"]) if req.get("item") else None}
    if reg_last != files_last:
        out["warning"] = f"Register ends at {reg_last} but SVR files go up to {files_last}. Resolve before issuing a number."
    return out


# ---------- document ----------
def run(text, rpr):
    parts = str(text).split("\n")
    inner = "<w:br/>".join(f'<w:t xml:space="preserve">{escape(p)}</w:t>' for p in parts)
    return f"<w:r><w:rPr>{rpr}</w:rPr>{inner}</w:r>"


def para(text, rpr, ppr=""):
    return f"<w:p><w:pPr>{ppr}<w:rPr>{rpr}</w:rPr></w:pPr>{run(text, rpr) if text else ''}</w:p>"


def once(xml, needle, start=0):
    i = xml.find(needle, start)
    if i < 0:
        raise SystemExit(f"Blank form changed: cannot find {needle[:60]!r}")
    return i


def last_open(xml, tag, i):
    """Start of the nearest <w:tag> before i that is still open at i (skips text-box content)."""
    pos = i
    while True:
        starts = [m.start() for m in re.finditer(rf"<w:{tag}[ >]", xml[:pos])]
        if not starts:
            raise SystemExit(f"Blank form changed: no open <w:{tag}> found")
        cand = starts[-1]
        seg = xml[cand:i]
        if len(re.findall(rf"<w:{tag}[ >]", seg)) - seg.count(f"</w:{tag}>") == 1:
            return cand
        pos = cand


def has_run(xml):
    return re.search(r"<w:r[ >]", xml) is not None


def set_cell(tc, paragraphs):
    m = re.match(r"(<w:tc>\s*<w:tcPr>.*?</w:tcPr>)", tc, re.S)
    if not m:
        raise SystemExit("Blank form changed: table cell without properties")
    return m.group(1) + "".join(paragraphs) + "</w:tc>"


def person_lines(req):
    ps = req.get("persons", [])
    company = (req.get("company") or "").strip()
    lines = []
    if len(ps) == 1:
        p = ps[0]
        lines.append(str(p.get("name", "")).upper())
        second = " ".join(x for x in [str(p.get("designation") or "").upper(), company.upper()] if x)
        if second:
            lines.append(second)
        if p.get("id"):
            lines.append(str(p["id"]))
    else:
        for p in ps:
            lines.append(" ".join(x for x in [str(p.get("name", "")).upper(), str(p.get("designation") or "").upper(), str(p.get("id") or "")] if x))
        if company:
            lines.append(company.upper())
    return lines


def prep_photo(path, n_photos):
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    im.thumbnail((1600, 1600))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    max_w, max_h = 15.0, (19.0 if n_photos == 1 else 10.5)
    w, h = im.size
    s = min(max_w / w, max_h / h)
    return buf.getvalue(), int(w * s * EMU_CM), int(h * s * EMU_CM)


def drawing(rid, n, cx, cy):
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    return (f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{9000 + n}" name="Evidence {n}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="{a}"><a:graphicData uri="{pic}"><pic:pic xmlns:pic="{pic}">'
            f'<pic:nvPicPr><pic:cNvPr id="{9000 + n}" name="evidence{n}.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>')


def make(req, out_path):
    date = datetime.date.fromisoformat(req["date"])
    mx = read_matrix(req["item"])
    src = template_path(date.year)
    with zipfile.ZipFile(src) as z:
        files = {n: z.read(n) for n in z.namelist()}
        order = z.namelist()
    xml = files["word/document.xml"].decode("utf-8")
    xml = re.sub(r' w:rsid\w*="[^"]*"', "", xml)   # revision ids only; removing them makes tags findable
    b, bb, n12 = "<w:b/>", '<w:b/><w:bCs/>', '<w:sz w:val="24"/>'

    # Ref.
    i = once(xml, "<w:t>Ref.</w:t>")
    j = once(xml, "</w:p>", i)
    xml = xml[:j] + run(str(req["no"]), TNR + b + n12) + xml[j:]

    # Location ... Date (date sits on a right tab so a long location cannot push it to the next line)
    i = once(xml, "<w:t>Location</w:t>")
    p0 = last_open(xml, "p", i)
    r0 = last_open(xml, "r", i)
    j = once(xml, "</w:p>", i)
    rpr = MB + n12 + '<w:szCs w:val="24"/>'
    new_runs = (run("Location: " + (req.get("location") or ""), rpr) + f"<w:r><w:rPr>{rpr}</w:rPr><w:tab/></w:r>" +
                run("Date: " + date.strftime("%b %d, %Y").replace(" 0", " "), rpr))
    head = xml[p0:r0]
    if "<w:tabs>" in head[:head.find("</w:pPr>")]:
        raise SystemExit("Blank form changed: Location line already has tab stops")
    head = head.replace("<w:pPr>", '<w:pPr><w:tabs><w:tab w:val="right" w:pos="9180"/></w:tabs>', 1)
    xml = xml[:p0] + head + new_runs + xml[j:]

    # Issued to
    i = once(xml, "<w:t>This notice is given to</w:t>")
    p_end = once(xml, "</w:p>", i) + 6
    q_end = once(xml, "</w:p>", p_end) + 6
    if has_run(xml[p_end:q_end]):
        raise SystemExit("Blank form changed: the name line is not empty")
    xml = xml[:p_end] + "".join(para(t, TNR + b + n12) for t in person_lines(req)) + xml[q_end:]

    # Violation details (first empty line of the description row)
    i = once(xml, '<w:trHeight w:val="1840"/>')
    j = once(xml, "</w:pPr></w:p>", i) + len("</w:pPr>")
    xml = xml[:j] + run(req.get("details") or "", TNR + bb + n12) + xml[j:]

    # Penalty matrix row
    i = once(xml, '<w:trHeight w:val="1264"/>')
    r_start = last_open(xml, "tr", i)
    r_end = once(xml, "</w:tr>", i)
    row = xml[r_start:r_end]
    pre = row[:row.find("<w:tc>")]
    cells = re.findall(r"<w:tc>.*?</w:tc>", row, re.S)
    if len(cells) != 6:
        raise SystemExit(f"Blank form changed: penalty row has {len(cells)} cells, expected 6")
    centre, tab = '<w:jc w:val="center"/>', '<w:tabs><w:tab w:val="left" w:pos="7755"/></w:tabs>'
    small = '<w:sz w:val="20"/><w:szCs w:val="20"/>' if len(mx["description"]) > 150 else n12
    mb12 = MB + '<w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    cells[0] = set_cell(cells[0], [para(f"{mx['item']:02d}", MB + b, tab + centre)])
    cells[1] = set_cell(cells[1], [para(mx["title"], TNR + bb + n12, tab + '<w:ind w:hanging="51"/>'),
                                   para(mx["description"], TNR + small, tab + '<w:ind w:hanging="51"/>')])
    for k, key in ((2, "first"), (3, "second"), (4, "third")):
        cells[k] = set_cell(cells[k], [para(t, mb12, centre) for t in (mx[key] or ["-"])])
    cells[5] = set_cell(cells[5], [para(req.get("penalty") or "", MB + b + '<w:sz w:val="24"/><w:szCs w:val="24"/>', centre)])
    xml = xml[:r_start] + pre + "".join(cells) + xml[r_end:]

    # Photos on the second page
    photos = [p for p in req.get("photos", []) if p]
    if photos:
        first = once(xml, "<w:sectPr")
        a = once(xml, "</w:p>", first) + 6
        z_ = xml.rfind("<w:sectPr")
        if z_ <= a or has_run(xml[a:z_]):
            raise SystemExit("Blank form changed: photo page is not empty")
        rels = files["word/_rels/document.xml.rels"].decode("utf-8")
        blocks = []
        for n, path in enumerate(photos, 1):
            data, cx, cy = prep_photo(path, len(photos))
            files[f"word/media/evidence{n}.jpeg"] = data
            order.append(f"word/media/evidence{n}.jpeg")
            rid = f"rIdEv{n}"
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/evidence{n}.jpeg"/></Relationships>')
            blocks.append(f'<w:p><w:pPr><w:spacing w:after="160"/><w:jc w:val="center"/></w:pPr><w:r>{drawing(rid, n, cx, cy)}</w:r></w:p>')
        files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        xml = xml[:a] + "".join(blocks) + xml[z_:]
        ct = files["[Content_Types].xml"].decode("utf-8")
        if 'Extension="jpeg"' not in ct:
            ct = ct.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ', 1)
            files["[Content_Types].xml"] = ct.encode("utf-8")

    files["word/document.xml"] = xml.encode("utf-8")
    import xml.etree.ElementTree as ET
    for part in ("word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"):
        try:
            ET.fromstring(files[part])
        except ET.ParseError as e:
            raise SystemExit(f"Generated {part} is not valid XML ({e}). Nothing written.")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    if os.path.exists(out_path):
        raise SystemExit(f"Refusing to overwrite an existing file: {out_path}")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(n, files[n])
    return {"written": out_path, "photos": len(photos), "template": src}


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if len(sys.argv) < 3 or sys.argv[1] not in ("plan", "make"):
        raise SystemExit(__doc__)
    req = json.load(open(sys.argv[2], encoding="utf-8-sig"))
    res = plan(req) if sys.argv[1] == "plan" else make(req, sys.argv[3])
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

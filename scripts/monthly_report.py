# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Monthly HSE report (JLL template) from the tool's records. Setup in monthly.json.

    python monthly_report.py plan <export dir> --month YYYY-MM --weekly weekly.json [--letters letters.json] [--prev report.docx]   -> values.json (path in MONTHLY_VALUES or %TEMP%)
    python monthly_report.py make <values.json> <out.docx> [--photos-dir DIR]

<export dir> holds the collections exported with ArtifactData: observations, violations, permits, consultant,
equipmentInspections, manpower, dailyLog. The month is the calendar month (the report is dated its last day).
`plan` lists the photo asset ids it wants; download them (Artifact read with path = id) into --photos-dir, where each
is saved as <id>.<ext>. `make` copies the previous report and refreshes: cover and issue-sheet dates, the incidents
table, consultant issues, achievements, photos, manpower, letters, training/TBT and meetings. Reference numbers are
kept as they are (the house reports repeat them), and the script says so. It never overwrites a file.
"""
import calendar, datetime as dt, glob, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
from config import CFG, ROOT
import permit as P
import monthly_sources as MS
import weekly_hse as W

MJ = json.load(open(os.path.join(ROOT, "monthly.json"), encoding="utf-8"))
MON = ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec"]   # as the house file names write them
CLASS = {"nearmiss": "Near Miss", "fac": "First Aid Case", "mtc": "Medical Treatment Case", "lti": "Lost Time Injury", "ftl": "Fatality",
         "fire": "Fire Incident", "mvi": "Motor Vehicle Incident", "illness": "Occupational Illness", "env": "Environmental Incident", "damage": "Property Damage"}


def dmy(s):
    return f"{s[8:10]}-{s[5:7]}-{s[:4]}" if s and re.match(r"^\d{4}-\d{2}-\d{2}", str(s)) else str(s or "")


def text_of(xml):
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml)).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")


# ---------- plan ----------

def prev_report(year, month):
    files = []
    for f in glob.glob(os.path.join(MJ["folder"], "**", "*Monthly HSE Report*.docx"), recursive=True):
        m = re.search(r"Report - ([A-Za-z]+) (\d{4})\.docx$", f)
        if not m:
            continue
        mi = next((i + 1 for i, n in enumerate(MON) if m.group(1).lower().startswith(n.lower()[:3])), None)
        if mi and (int(m.group(2)), mi) < (year, month):
            files.append(((int(m.group(2)), mi), f))
    if not files:
        raise SystemExit("No earlier monthly report found.")
    return sorted(files)[-1][1]


def cmd_plan(export, *args):
    opt = dict(zip(args[::2], args[1::2]))
    y, m = map(int, opt["--month"].split("-"))
    first, last = dt.date(y, m, 1), dt.date(y, m, calendar.monthrange(y, m)[1])
    a, b = str(first), str(last)
    inm = lambda s: bool(s) and a <= s <= b
    month_name = f"{calendar.month_name[m]} {y}"
    prev = opt.get("--prev") or prev_report(y, m)
    notes = []

    logs = [x for x in W.load(export, "dailyLog") if inm(x.get("date"))]
    mp = {d.get("date") or d["_id"]: d for d in W.load(export, "manpower")}
    cons = [c for c in W.load(export, "consultant") if not c.get("example")]
    insp = [x for x in W.load(export, "equipmentInspections") if inm(x.get("date"))]

    # every figure from its source (monthly_sources.py); nothing is estimated
    if "--weekly" not in opt:
        raise SystemExit("Pass --weekly <weekly.json> from scripts\\read_weekly.ps1: the weekly reports are the source of manpower, hours, TBT and training.")
    S = MS.compute(f"{y}-{m:02d}", opt["--weekly"], lambda c: W.load(export, c), opt.get("--letters"))
    notes += S["checks"]
    hours, avg, ptd = S["hours"], S["avg"], S["toDate"]
    x = zipfile.ZipFile(prev).read("word/document.xml").decode("utf-8")
    for p in re.findall(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", x, flags=re.S):
        mm = re.search(r"milestone of ([\d,\.]+) safe", text_of(p))
        if mm and ptd is not None:
            was = int(re.sub(r"\D", "", mm.group(1)))
            notes.append(f"last month's report said {mm.group(1)} to date; the weekly reports give {ptd - hours:,} at its month end")
    incidents = S["incidents"]
    lti = any(re.search(r"\bLTI\b|fatal", i["type"], re.I) for i in incidents)
    if lti:
        notes.append("an LTI or fatality this month: the 'without a Lost Time Incident' sentence is left out; write the LTI statement by hand")
    tbts, trainings, insp_rows, meetings_held = S["tbt"], S["trainings"], S["inspections"], S["meetings"]
    drills = [e for e in logs if e.get("kind") == "drill"]
    ph = MJ["phrases"]
    bullets = []
    short = f"{MON[m - 1]} {y}"
    if ptd is not None and not lti:
        if ptd // 1_000_000 > (ptd - hours) // 1_000_000:          # crossed a million this month (as in July 2025)
            bullets.append(ph["million"].format(n=ptd // 1_000_000, month=short))
        bullets.append(ph["ptd"].format(ptd=ptd))
    bullets.append((ph["month_lti"] if lti else ph["month"]).format(month=short, hours=hours))
    if tbts:
        bullets.append(ph["daily_tbt"])
    num = lambda v: int(re.sub(r"\D", "", str(v)) or 0)
    mass = {}
    for e in tbts:                                   # distinct topics, ignoring case and spacing, in the order held
        k = re.sub(r"\W+", " ", str(e.get("topic", ""))).strip().lower()
        if k and num(e.get("attendees")) >= MJ["mass_tbt_min"] and k not in mass:
            mass[k] = e["topic"].strip()
    if mass:
        topics = list(mass.values())
        bullets.append(ph["mass_tbt"].format(topics=", ".join(topics[:4]) + (" and other topics" if len(topics) > 4 else "")))
    used = set()
    for e in insp_rows:
        topic = str(e.get("topic", ""))
        hit = next((s for k, s in ph["inspections"] if re.search(k, topic, re.I)), None)
        s = hit or ph["inspection_other"].format(topic=topic)
        if s not in used:
            used.add(s); bullets.append(s)
    if any(x2.get("kind") == "monthly" for x2 in insp):
        bullets.append(ph["equipment"])
    tr_topics = sorted({e["topic"] for e in trainings if e.get("topic")})
    if tr_topics:
        bullets.append(ph["training"].format(topics=", ".join(tr_topics)))
    for e in sorted([e for e in logs if e.get("kind") == "achievement"], key=lambda e: e.get("date", "")):
        bullets.append(str(e.get("description") or e.get("topic") or "").strip())       # events, awards, new facilities
    if drills:
        bullets.append(ph["drill"].format(topics=", ".join(sorted({d.get("topic", "") for d in drills if d.get("topic")}))))
    for p in re.findall(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", x, flags=re.S):
        if 'w:val="JLLBullet1"' in p and text_of(p).startswith("HSE Staff Ratio"):
            bullets.append(text_of(p).strip())
            break

    # tables
    inc_rows = [[dt.date.fromisoformat(i["date"]).strftime("%d-%m-%y"), i["company"], i["type"], i["wps"], i["description"], "-"] for i in incidents]
    pace_rows = []        # the house reports leave "Issues & Concerns" blank and list consultant notices under letters
    letters = S["letters"]
    train_rows = [[str(n + 1), e["topic"], e.get("scope") or "Internal", str(e.get("attendees") or ""), MJ["hse_team"]] for n, e in enumerate(tbts + trainings)]
    weeks = sum(1 for d in range(1, last.day + 1) if dt.date(y, m, d).weekday() == 2)
    meet = MJ["meeting"]
    meet_rows = [[meet["title"], meet["orgs"], meet["frequency"], str(meet["per_week"] * weeks), str(meetings_held)]]
    worked_days = S["workingDays"]

    # photos: one per kind and topic, newest first, up to 9
    photos, seen = [], set()
    def caption(e):
        t = str(e.get("topic") or "").strip()
        t = " ".join(w if w.isupper() else w[:1].upper() + w[1:] for w in t.split())     # "fire extinguisher" -> "Fire Extinguisher"
        k = e["kind"]
        if k == "tbt":
            return "Weekly Mass TBT" if int(e.get("attendees") or 0) >= MJ["mass_tbt_min"] else "Daily TBT"
        if k == "inspection":
            return t if re.search(r"inspection", t, re.I) else f"{t} Inspection".strip()
        if k == "training":
            return t if re.search(r"training", t, re.I) else f"{t} Training".strip()
        if k == "drill":
            return "Emergency Drill"
        if k == "meeting":
            return "HSE Meeting"
        return t or "Site Tour"
    kinds = ["tbt", "inspection", "training", "achievement", "drill", "meeting"]
    for e in sorted([e for e in logs if e.get("photo") and e.get("kind") in kinds], key=lambda e: (kinds.index(e["kind"]), e.get("date", ""))):
        c = caption(e)
        if c.lower() in seen:
            continue
        seen.add(c.lower())
        photos.append({"id": e["photo"], "caption": c})
        if len(photos) >= 9:
            break
    eqp = next((x2 for x2 in insp if x2.get("kind") == "monthly" and x2.get("photos")), None)
    if eqp and len(photos) < 9 and "equipment inspection" not in seen:
        photos.append({"id": eqp["photos"][0], "caption": "Equipment Inspection"})

    values = {"prev": prev, "month": month_name, "monthShort": f"{MON[m - 1]} {y}", "date_long": f"{last.day:02d} {calendar.month_name[m]} {y}",
              "date_dmy": f"{last.day:02d}-{m:02d}-{y}", "bullets": bullets, "incidents": inc_rows, "pace": pace_rows, "letters": letters,
              "training": train_rows, "meetings": meet_rows, "manpower": [[MJ["company"], str(avg), str(hours), str(worked_days)]], "photos": photos,
              "out_name": MJ["file_name"].format(Mon=MON[m - 1], year=y)}
    dest = os.environ.get("MONTHLY_VALUES") or os.path.join(os.environ.get("TEMP", "."), "monthly_values.json")
    json.dump(values, open(dest, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"month": month_name, "previous": prev, "values": dest, "manpower": values["manpower"][0], "bullets": bullets,
                      "incidents": len(inc_rows), "pace": len(pace_rows), "letters": len(letters), "training_rows": len(train_rows),
                      "meetings": meet_rows[0], "photo_ids": [p["id"] for p in photos], "sources": S["sources"], "missing": S["missing"], "notes": notes}, ensure_ascii=False, indent=1))


# ---------- make ----------

def para(ppr, rpr, text):
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def first(pattern, s, default=""):
    m = re.search(pattern, s, flags=re.S)
    return m.group(0) if m else default


def cell_text(tc, text):
    """The cell with one paragraph holding text, in the cell's own paragraph and run formatting."""
    head = first(r"<w:tc>|<w:tc [^>]*>", tc) + first(r"<w:tcPr>.*?</w:tcPr>", tc)
    p = first(r"<w:p[ >].*?</w:p>|<w:p/>", tc)
    ppr = first(r"<w:pPr>.*?</w:pPr>", p)
    run = first(r"<w:r[ >].*?</w:r>", p)
    rpr = first(r"<w:rPr>.*?</w:rPr>", run) or first(r"<w:rPr>.*?</w:rPr>", ppr)
    return head + para(ppr, rpr, text) + "</w:tc>"


def inline_drawing(cell):
    """Turn a floating picture (<wp:anchor>, as in the house photo grid) into an inline one, so it sits in its cell."""
    m = re.search(r"<wp:anchor\b.*?</wp:anchor>", cell, flags=re.S)
    if not m:
        return cell
    a = m.group(0)
    parts = [first(p, a) for p in (r"<wp:extent [^>]*/>", r"<wp:effectExtent [^>]*/>", r"<wp:docPr [^>]*?(/>|>.*?</wp:docPr>)",
                                   r"<wp:cNvGraphicFramePr>.*?</wp:cNvGraphicFramePr>|<wp:cNvGraphicFramePr/>", r"<a:graphic .*?</a:graphic>")]
    inline = '<wp:inline distT="0" distB="0" distL="0" distR="0">' + "".join(parts) + "</wp:inline>"
    return cell[:m.start()] + inline + cell[m.end():]


def find_table(xml, test):
    for t in P.tables(xml):
        if t and test([P.norm(c.text(xml)) for c in t[0]], t):
            return t
    return None


def table_span(xml, t):
    """Start/end of the whole <w:tbl> holding row list t."""
    s = xml.rfind("<w:tbl>", 0, t[0][0].start)
    e = xml.find("</w:tbl>", t[-1][-1].end) + len("</w:tbl>")
    return s, e


def row_at(xml, pos):
    """Start of the <w:tr> row holding pos (not its <w:trPr>, which also begins with '<w:tr')."""
    return max(xml.rfind("<w:tr>", 0, pos), xml.rfind("<w:tr ", 0, pos))


def rebuild_rows(xml, t, rows, min_rows, blank="-"):
    """Replace the data rows of table t (all rows after the header) with rows, copying the first data row's look."""
    row_start = lambda c: row_at(xml, c.start)
    data = t[1:]
    if not data:
        return xml
    s = row_start(data[0][0])
    e = xml.find("</w:tr>", data[-1][-1].end) + len("</w:tr>")
    tpl = xml[s:xml.find("</w:tr>", data[0][-1].end) + len("</w:tr>")]
    cells = [xml[c.start:c.end] for c in data[0]]
    pre = tpl[:tpl.find("<w:tc")]
    out = []
    for r in rows + [[blank] * len(cells)] * max(0, min_rows - len(rows)):
        out.append(pre + "".join(cell_text(c, str(r[i]) if i < len(r) else "") for i, c in enumerate(cells)) + "</w:tr>")
    return xml[:s] + "".join(out) + xml[e:]


def replace_paras(xml, test, new_text):
    def fix(m):
        p = m.group(0)
        t = text_of(p)
        if not test(t):
            return p
        ppr = first(r"<w:pPr>.*?</w:pPr>", p)
        rpr = first(r"<w:rPr>.*?</w:rPr>", first(r"<w:r[ >].*?</w:r>", p))
        return para(ppr, rpr, new_text(t))
    return re.sub(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", fix, xml, flags=re.S)


def cmd_make(values_json, out, *args):
    opt = dict(zip(args[::2], args[1::2]))
    v = json.load(open(values_json, encoding="utf-8-sig"))
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    z = zipfile.ZipFile(v["prev"])
    files = {n: z.read(n) for n in z.namelist()}
    x = files["word/document.xml"].decode("utf-8")
    notes = []

    # cover and issue sheet
    x = replace_paras(x, lambda t: t.strip().startswith("Date:") and re.search(r"\d{4}", t), lambda t: f"Date: {v['date_long']}")
    x = replace_paras(x, lambda t: t.strip().startswith("Month:"), lambda t: f"Month: {v['month']}")
    t = find_table(x, lambda h, t: any(c.startswith("DOCUMENT NUMBER") or c.startswith("PROJECT NAME") for c in h) and any(P.norm(c.text(x)) == "REV." for r in t for c in r))
    if t:
        hdr = next(i for i, r in enumerate(t) if P.norm(r[0].text(x)) == "REV.")
        if hdr + 1 < len(t):
            c = t[hdr + 1][1]
            x = x[:c.start] + cell_text(x[c.start:c.end], v["date_dmy"]) + x[c.end:]
    notes.append("reference numbers kept as in the previous report (cover 'Ref.' and issue-sheet 'Document Number'); update them if your numbering changes monthly")

    # achievements: the JLLBullet1 paragraphs become this month's bullets
    bl = [m for m in re.finditer(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", x, flags=re.S) if 'w:val="JLLBullet1"' in m.group(0)]
    if bl:
        p0 = bl[0].group(0)
        ppr = first(r"<w:pPr>.*?</w:pPr>", p0)
        rpr = first(r"<w:rPr>.*?</w:rPr>", first(r"<w:r[ >].*?</w:r>", p0))
        new = "".join(para(ppr, rpr, b) for b in v["bullets"])
        x = x[:bl[0].start()] + new + x[bl[-1].end():]
    # tables
    for test, rows, mn, blank in (
            (lambda h, t: h[:3] == ["DATE", "COMPANY", "CLASS"], v["incidents"], 3, "-"),
            (lambda h, t: h[:2] == ["TOPIC", "ISSUE/CONCERN"], v["pace"], 6, ""),
            (lambda h, t: h[:2] == ["COMPANY", "AVE MANPOWER"], v["manpower"], 1, ""),
            (lambda h, t: h[:1] == ["LETTER REFERENCE"], v["letters"], 3, ""),
            (lambda h, t: h[:2] == ["S.NO", "TITLE"], v["training"], 1, ""),
            (lambda h, t: h[:1] == ["TYPE/TITLE"], v["meetings"], 1, "")):
        t = find_table(x, test)
        if t:
            x = rebuild_rows(x, t, rows, mn, blank)
        else:
            notes.append("a table was not found: " + str(rows[:1]))

    # photos
    pdir = opt.get("--photos-dir")
    # the photo grid: a table of two or more columns with pictures in several cells (not the cover, which has one picture)
    t = find_table(x, lambda h, t: len(t[0]) >= 2 and sum("<w:drawing>" in x[c.start:c.end] for r in t for c in r) >= 2)
    if t and pdir:
        rels = files["word/_rels/document.xml.rels"].decode("utf-8")
        ct = files["[Content_Types].xml"].decode("utf-8")
        img_cell = next(x[c.start:c.end] for r in t for c in r if "<w:drawing>" in x[c.start:c.end])
        cap_cell = x[t[0][0].start:t[0][0].end]
        ncols = len(t[0])
        from PIL import Image
        imgs = []
        for i, p in enumerate(v["photos"]):
            src = next(iter(glob.glob(os.path.join(pdir, p["id"] + ".*"))), None)
            if not src:
                notes.append(f"photo {p['id']} not downloaded; left out")
                continue
            ext = os.path.splitext(src)[1].lower().lstrip(".")
            ext = "jpeg" if ext in ("jpg", "jpeg") else ext
            name = f"media/mrpt_{i + 1}.{ext}"
            files["word/" + name] = open(src, "rb").read()
            rid = f"rIdMrpt{i + 1}"
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{name}"/></Relationships>')
            if f'Extension="{ext}"' not in ct:
                ct = ct.replace("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
                                f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="{ext}" ContentType="image/{ext}"/>')
            w, h = Image.open(src).size
            cx0, cy0 = map(int, re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', img_cell).groups())
            s = min(cx0 / w, cy0 / h)
            cx, cy = int(w * s), int(h * s)
            c = re.sub(r'r:embed="[^"]+"', f'r:embed="{rid}"', img_cell)
            c = re.sub(r'<wp:extent cx="\d+" cy="\d+"', f'<wp:extent cx="{cx}" cy="{cy}"', c)
            c = re.sub(r'<a:ext cx="\d+" cy="\d+"', f'<a:ext cx="{cx}" cy="{cy}"', c)
            c = re.sub(r'<wp:docPr id="\d+"', f'<wp:docPr id="{7000 + i}"', c)
            c = inline_drawing(c)
            c = re.sub(r' w14:(paraId|textId)="[^"]*"', "", c)        # copies must not repeat the template's paragraph ids
            # the cell style has an exact line height, which would clip an inline picture to one line: use automatic spacing
            ppr = first(r"<w:pPr>.*?</w:pPr>", c)
            ppr2 = re.sub(r"<w:spacing [^>]*/>|<w:jc [^>]*/>", "", ppr)
            ppr2 = ppr2.replace("</w:pStyle>", "</w:pStyle>").replace("<w:rPr>", '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr>', 1) \
                if "<w:rPr>" in ppr2 else ppr2.replace("</w:pPr>", '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>')
            c = c.replace(ppr, ppr2, 1) if ppr else c
            imgs.append((p["caption"], c))
        empty_img = cell_text(img_cell, "")
        rows_xml = []
        # each kind of row keeps its own row settings: the first caption row (repeating header), later caption rows,
        # and the picture rows (which carry the fixed row height)
        pre = lambda r: x[row_at(x, r[0].start):r[0].start]
        tr_head = pre(t[0])
        tr_img = pre(t[1]) if len(t) > 1 else tr_head
        tr_cap = pre(t[2]) if len(t) > 2 else tr_head.replace("<w:tblHeader/>", "")
        for k in range(0, max(len(imgs), 1), ncols):
            chunk = imgs[k:k + ncols]
            caps = [cell_text(cap_cell, c[0]) for c in chunk] + [cell_text(cap_cell, "")] * (ncols - len(chunk))
            pics = [c[1] for c in chunk] + [empty_img] * (ncols - len(chunk))
            rows_xml.append((tr_head if k == 0 else tr_cap) + "".join(caps) + "</w:tr>" + tr_img + "".join(pics) + "</w:tr>")
        s0 = row_at(x, t[0][0].start)
        e0 = x.find("</w:tr>", t[-1][-1].end) + len("</w:tr>")
        x = x[:s0] + "".join(rows_xml) + x[e0:]
        files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        files["[Content_Types].xml"] = ct.encode("utf-8")
    elif t:
        notes.append("no photos directory given: the previous month's photos were kept; replace them or pass --photos-dir")

    # ask Word to refresh the table of contents on opening
    st = files["word/settings.xml"].decode("utf-8")
    if "<w:updateFields" not in st:
        st = re.sub(r"(<w:settings[^>]*>)", r'\1<w:updateFields w:val="true"/>', st, count=1)
        files["word/settings.xml"] = st.encode("utf-8")
    files["word/document.xml"] = x.encode("utf-8")
    import xml.dom.minidom
    for n in ("word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml", "word/settings.xml"):
        xml.dom.minidom.parseString(files[n])
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, b in files.items():
            zo.writestr(n, b)
    print(json.dumps({"written": out, "notes": notes}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    {"plan": cmd_plan, "make": cmd_make}[sys.argv[1]](*sys.argv[2:])

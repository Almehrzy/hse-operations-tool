# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Incident reports on the project's forms, from the page's Incidents tab (procedure AVK01-JLL-HSE-PRC-0001).

    python incident.py plan <incident.json> [--code PD]                 numbers, folder, deadlines and register rows
    python incident.py initial <request.json> <out.docx>                Initial Incident Report (FRM-00012, house copy of INI-PD-00001)
    python incident.py investigation <request.json> <out.docx>          Incident Investigation Report (FRM-00013, house copy of INV-PD-00001)
    python incident.py witness <request.json> <n> <out.docx>            Witness Statement for witness n (1 = first)
    python incident.py capa <request.json> <out.docx>                   CAPA plan for the incident (house copy of CAPA-002)
    python incident.py alert <request.json> <out.pptx>                  Safety Alert (FRM-00015 slide) for WPS4-5: ref <prefix>SA-<ref_no>;
                                                                        text: alert_title (≤45 chars), what, why [..], actions_taken [..],
                                                                        lessons [..]; the first two photos; warns when text will not fit

incident.json: the `incidents` document exported from the page (with its "id").
request.json: the same document plus what Claude prepares (formal English; the supervisor's words are data, not text
to paste as they are):
 "ref_no": "PD-00002"                         from plan (code and number); refs are <ref_prefix>INI-/INV-/WIT-<ref_no>
 "report_date": "YYYY-MM-DD"                   default today
 "text": {"what", "when", "where", "who", "how", "recommendations": [..], "commentary",           (initial report)
          "brief", "findings": [..], "timeline": [..], "docs", "emergency": [..], "immediate_actions": [..],
          "immediate_causes": [[factor, detail]], "system_causes": [[factor, detail]], "analysis",
          "learning": [..], "preventive": [..], "effectiveness": [..], "closure"}                (investigation / CAPA)
 "comm": {"MS approved": "Y", ...}            the Y/N grid of the investigation form, by its labels (default: kept blank)
 "witnesses": [{"name", "department", "company", "dob", "trade", "statement", "recorded_by", "date"}]  (default: from the page)
 "actions": [{"no", "cause", "action", "owner", "due", "status"}]  corrective actions (default: the linked CAPA documents)
 "photos": [{"file", "caption"}]              downloaded incident photos, in order
 "authority_notified": "YYYY-MM-DD"            optional
 "capa_doc": "AVK01-WP.03-HSE-CAPA-003"       CAPA plan number (plan proposes the next one)
 "reg_no": 6                                  the incident's No. in REG-013 (set when the initial row is written)
Signature images in the house copies are removed: each person signs the printed or PDF copy. Never overwrites a file.

Registers (rows for incident_register.ps1, which writes them through Excel with a backup):
    python incident.py regrows <request.json> initial|investigation|closed <rows.json>    REG-013 row: add, then fill
    python incident.py regrows <request.json> minor <rows.json>                           REG-014 row (minor events)
    python incident.py capa-rows <capa export dir> <incidents export dir> <rows.json>    REG-011: new and changed CAPA
"""
import datetime as dt, glob, io, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
from config import CFG, path as cfg_path
import permit as P
import monthly_report as MR
import svn_reply as SR

IC = CFG["incidents"]
PLACEHOLDER = SR.PLACEHOLDER


def dmy(s, sep="/"):
    return dt.date.fromisoformat(str(s)[:10]).strftime(f"%d{sep}%m{sep}%Y") if s else ""


def long(s):
    return dt.date.fromisoformat(str(s)[:10]).strftime("%d %B %Y") if s else ""


def refs(req):
    n = req["ref_no"]
    pre = IC["ref_prefix"]
    return {"ini": f"{pre}INI-{n}", "inv": f"{pre}INV-{n}", "wit": f"{pre}WIT-{n}"}


# ---------- plan ----------

def existing_numbers():
    """{code: highest number} from report file names in the incident folder and the minor incident register."""
    seen = {}
    pat = re.compile(r"HSE-(?:INI|INV|WIT)-([A-Z]+)-(\d{5})")
    names = [os.path.basename(f) for f in glob.glob(os.path.join(IC["folder"], "**", "*"), recursive=True)]
    try:
        import openpyxl
        ws = openpyxl.load_workbook(IC["register_014"], read_only=True, data_only=True).worksheets[0]
        names += [str(c) for r in ws.iter_rows(values_only=True) for c in r if c]
    except Exception:
        pass
    for s in names:
        for code, num in pat.findall(s):
            seen[code] = max(seen.get(code, 0), int(num))
    return seen


def next_capa_doc():
    pat = re.compile(re.escape(IC["capa_doc_prefix"]) + r"(\d{3,})")
    hi = 0
    for f in glob.glob(os.path.join(IC["folder"], "**", "*.docx"), recursive=True):
        try:
            x = zipfile.ZipFile(f).read("word/document.xml").decode("utf-8", "ignore")
        except Exception:
            continue
        for n in pat.findall(MR.text_of(x)):
            hi = max(hi, int(n))
    return f"{IC['capa_doc_prefix']}{hi + 1:03d}"


def cmd_plan(inc_json, *args):
    opt = dict(zip(args[::2], args[1::2]))
    x = json.load(open(inc_json, encoding="utf-8-sig"))
    x = x.get("data", x)
    code = opt.get("--code") or IC["codes"].get(x.get("type"), "INC")
    nums = existing_numbers()
    ref_no = f"{code}-{nums.get(code, 0) + 1:05d}"
    at = dt.datetime.fromisoformat(f"{x['date']}T{x.get('time') or '00:00'}")
    wps = int(x.get("wps") or 0) or None
    formal = x.get("formal")
    notes = []
    if wps is None:
        notes.append("WPS not set by HSE yet (the page suggests one from the type): ask the safety officer before issuing")
    if formal is False:
        notes.append("classified as minor: REG-014 (minor incident log) only, no formal report")
    w = wps or 2
    due = {"notify": at + dt.timedelta(minutes=15), "initial": at + dt.timedelta(hours=24), "investigation": at + dt.timedelta(days=5 if w >= 4 else 3)}
    if x.get("notifiedAt"):
        na = dt.datetime.fromisoformat(x["notifiedAt"].replace("Z", "+00:00")).astimezone(dt.timezone(dt.timedelta(hours=CFG["trend"].get("utc_offset_hours", 3)))).replace(tzinfo=None)
        if na > due["notify"]:
            notes.append(f"consultant and Engineer were notified at {na:%H:%M}, after the 15 minutes")
    else:
        notes.append("no notification time recorded: ask when PACE and JLL were told")
    if w >= 4:
        notes.append("WPS4-5: the Engineer joins the investigation; a safety alert is due within 48 h of the report's approval")
    folder = os.path.join(IC["folder"], f"{IC['type_text'].get(x.get('type'), 'Incident')} {at:%d-%m-%Y}")
    pre = IC["ref_prefix"]
    print(json.dumps({"ref_no": ref_no, "initial": f"{pre}INI-{ref_no}", "investigation": f"{pre}INV-{ref_no}", "witness": f"{pre}WIT-{ref_no}",
                      "capa_doc": next_capa_doc(), "folder": folder, "folder_exists": os.path.isdir(folder),
                      "files": {"initial": f"{pre}INI-{ref_no} (Initial Report).docx", "investigation": f"{pre}INV-{ref_no} (Final Report) - Incident Investigation Report.docx",
                                "witness": f"{pre}WIT-{ref_no} - Witness Statement.docx", "capa": f"CAPA - <short title>.docx"},
                      "due": {k: v.strftime("%Y-%m-%d %H:%M") for k, v in due.items()}, "now": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
                      "highest_existing": nums, "notes": notes}, ensure_ascii=False, indent=1))


# ---------- Word editing ----------

class Doc:
    def __init__(self, template):
        z = zipfile.ZipFile(cfg_path(template))
        self.files = {n: z.read(n) for n in z.namelist()}
        self.x = self.files["word/document.xml"].decode("utf-8")
        self.notes = []
        self.media = 0
        self.form_text = set(PLACEHOLDER.findall(MR.text_of(self.x)))      # the form's own bracketed hints, e.g. "[where applicable]"

    # paragraphs
    def relabel(self, label_re, value):
        """'Label: old value' paragraph → 'Label: value' (label in the first run's look, value in the last run's)."""
        done = [False]

        def fix(m):
            p = m.group(0)
            t = MR.text_of(p)
            mm = re.match(r"\s*(" + label_re + r")", t)
            if done[0] or not mm:
                return p
            done[0] = True
            runs = re.findall(r"<w:r[ >](?:(?!</w:r>).)*?</w:r>", p, flags=re.S)
            runs = [r for r in runs if "<w:t" in r]
            rpr1 = MR.first(r"<w:rPr>.*?</w:rPr>", runs[0]) if runs else ""
            rpr2 = MR.first(r"<w:rPr>.*?</w:rPr>", runs[-1]) if runs else ""
            ppr = MR.first(r"<w:pPr>.*?</w:pPr>", p)
            lab = mm.group(1).rstrip() + " "
            return (f'<w:p>{ppr}<w:r>{rpr1}<w:t xml:space="preserve">{escape(lab)}</w:t></w:r>'
                    f'<w:r>{rpr2}<w:t xml:space="preserve">{escape(str(value))}</w:t></w:r></w:p>')
        self.x = re.sub(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", fix, self.x, flags=re.S)
        if not done[0]:
            self.notes.append(f"label not found: {label_re}")

    # cells
    def cells(self):
        return [(t, r, c) for t in P.tables(self.x) for r in t for c in r]

    def put(self, cell, lines):
        """Replace a cell's paragraphs with lines (each its own paragraph), in the cell's first paragraph's look."""
        lines = [lines] if isinstance(lines, str) else [str(l) for l in lines if str(l).strip()] or [""]
        tc = self.x[cell.start:cell.end]
        head = MR.first(r"<w:tc>|<w:tc [^>]*>", tc) + MR.first(r"<w:tcPr>.*?</w:tcPr>", tc)
        p = MR.first(r"<w:p[ >].*?</w:p>|<w:p/>", tc)
        ppr = MR.first(r"<w:pPr>.*?</w:pPr>", p)
        run = next((r for r in re.findall(r"<w:r[ >](?:(?!</w:r>).)*?</w:r>", p, flags=re.S) if "<w:t" in r), "")
        rpr = MR.first(r"<w:rPr>.*?</w:rPr>", run) or MR.first(r"<w:rPr>.*?</w:rPr>", ppr)
        ppr = re.sub(r' w14:(paraId|textId)="[^"]*"', "", ppr)
        body = "".join(MR.para(ppr, rpr, l) for l in lines)
        self.x = self.x[:cell.start] + head + body + "</w:tc>" + self.x[cell.end:]

    def find(self, test):
        for t, r, c in self.cells():
            if test(P.norm(c.text(self.x))):
                return t, r, c
        return None, None, None

    def beside(self, label, value):
        """Cell holding label → the next cell in the same row gets value."""
        t, r, c = self.find(lambda s: s.rstrip(":").strip() == label.upper())
        if not c:
            self.notes.append(f"cell not found: {label}")
            return
        i = r.index(c)
        if i + 1 < len(r):
            self.put(r[i + 1], value)

    def below(self, label_start, value, exact=False):
        """Row whose single cell starts with (or, exact, is) label → the next row's first cell gets value."""
        for t in P.tables(self.x):
            for k, r in enumerate(t):
                s = P.norm(r[0].text(self.x))
                if (s == label_start.upper() if exact else s.startswith(label_start.upper())) and k + 1 < len(t):
                    self.put(t[k + 1][0], value)
                    return
        self.notes.append(f"section not found: {label_start}")

    def rows_after(self, header, rows, min_rows=1):
        """Replace the data rows under a header row (cells == header) with rows, copying the first data row's look.
        Data rows are the following rows with the same number of cells that are not another header."""
        H = [h.upper() for h in header]
        for t in P.tables(self.x):
            for k, r in enumerate(t):
                if [P.norm(c.text(self.x)) for c in r][:len(H)] == H and len(r) == len(H):
                    data = []
                    for r2 in t[k + 1:]:
                        if len(r2) != len(H) or [P.norm(c.text(self.x)) for c in r2] == H:
                            break
                        data.append(r2)
                    if not data:
                        self.notes.append(f"no data row under {header}")
                        return
                    s = MR.row_at(self.x, data[0][0].start)
                    e = self.x.find("</w:tr>", data[-1][-1].end) + len("</w:tr>")
                    tr_pre = self.x[s:data[0][0].start]
                    cells = [self.x[c.start:c.end] for c in data[0]]
                    out = []
                    for row in rows + [[""] * len(cells)] * max(0, min_rows - len(rows)):
                        out.append(tr_pre + "".join(MR.cell_text(cl, str(row[i]) if i < len(row) else "") for i, cl in enumerate(cells)) + "</w:tr>")
                    self.x = self.x[:s] + "".join(out) + self.x[e:]
                    return
        self.notes.append(f"table not found: {header}")

    def blank_drawings_in(self, test):
        """Remove pictures (signatures) from cells whose row has a label matching test."""
        for t in P.tables(self.x):
            for r in t:
                labels = " ".join(P.norm(c.text(self.x)) for c in r)
                if test(labels, r):
                    for c in reversed(r):
                        tc = self.x[c.start:c.end]
                        if "<w:drawing" in tc:
                            self.put(c, "")
                    return True
        return False

    def footer_replace(self, old, new):
        for n in list(self.files):
            if re.match(r"word/(header|footer)\d*\.xml", n):
                s = self.files[n].decode("utf-8")
                s2, ok = SR.replace_across_runs(s, old, new)
                if ok:
                    self.files[n] = s2.encode("utf-8")

    # pictures
    def add_image(self, path, max_w_cm, max_h_cm):
        """New inline picture run for an image file; registers the media part."""
        self.media += 1
        data, cx, cy = SR.prep_image(path, max_w_cm, max_h_cm)
        name = f"media/inc_{self.media}.jpeg"
        rid = f"rIdInc{self.media}"
        self.files["word/" + name] = data
        rels = self.files["word/_rels/document.xml.rels"].decode("utf-8")
        rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{name}"/></Relationships>')
        self.files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        ct = self.files["[Content_Types].xml"].decode("utf-8")
        if 'Extension="jpeg"' not in ct:
            ct = re.sub(r"(<Types [^>]*>)", r'\1<Default Extension="jpeg" ContentType="image/jpeg"/>', ct, count=1)
            self.files["[Content_Types].xml"] = ct.encode("utf-8")
        return f'<w:r>{SR.drawing(rid, 900 + self.media, cx, cy)}</w:r>'

    def save(self, out):
        if os.path.exists(out):
            raise SystemExit(f"Refusing to overwrite {out}")
        left = [m for m in PLACEHOLDER.findall(MR.text_of(self.x)) if m.strip("[]").strip() and m not in self.form_text]
        if left:
            raise SystemExit("Not written: placeholder text left in the report: " + ", ".join(sorted(set(left))))
        self.files["word/document.xml"] = self.x.encode("utf-8")
        import xml.dom.minidom
        for n in self.files:
            if n.endswith(".xml") or n.endswith(".rels"):
                xml.dom.minidom.parseString(self.files[n])
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for n, b in self.files.items():
                zo.writestr(n, b)
        print(json.dumps({"written": out, "notes": self.notes}, ensure_ascii=False, indent=1))


# ---------- content helpers ----------

def load(path):
    """A request file, or an exported document ({"id", "data": {...}}): the document's fields plus any top-level keys."""
    r = json.load(open(path, encoding="utf-8-sig"))
    base = r["data"] if isinstance(r.get("data"), dict) else {}
    return {**base, **{k: v for k, v in r.items() if k != "data"}}


def label_row(d, label, values):
    """A row like [Name | value | Name | value]: the cells after each label cell get values, in order."""
    for t in P.tables(d.x):
        for r in t:
            idx = [i for i, c in enumerate(r) if P.norm(c.text(d.x)).rstrip(":") == label.upper() and i + 1 < len(r)]
            if idx:
                for k in reversed(range(min(len(idx), len(values)))):
                    d.put(r[idx[k] + 1], values[k])
                return
    d.notes.append(f"row not found: {label}")


def T(req, key, default=""):
    v = (req.get("text") or {}).get(key)
    return v if v not in (None, "", []) else default


def type_text(req):
    return IC["type_text"].get(req.get("type"), "Incident")


def when_text(req):
    d = dt.datetime.fromisoformat(f"{req['date']}T{req.get('time') or '00:00'}")
    return f"{d:%A}, {d:%d %b %Y}, {d:%I:%M %p}" + (f", {T(req, 'visibility')}" if T(req, "visibility") else "")


def who_text(req):
    ps = req.get("persons") or []
    out = []
    for p in ps:
        out.append(", ".join(x for x in [f"Name: {p['name']}" if p.get("name") else "", f"DOB: {p['dob']}" if p.get("dob") else "",
                                           f"Trade: {p['trade']}" if p.get("trade") else "", f"Company: {p['company']}" if p.get("company") else "",
                                           f"Iqama/ID: {p['idNo']}" if p.get("idNo") else ""] if x))
    if req.get("thirdParty"):
        out.append(req["thirdParty"])
    return out or ["No person injured."]


def witnesses(req):
    if req.get("witnesses") and isinstance(req["witnesses"][0], dict) and "department" in req["witnesses"][0]:
        return req["witnesses"]
    return [{"name": w.get("name", ""), "department": w.get("trade", ""), "company": w.get("company", ""), "statement": w.get("statement", ""),
             "recorded_by": w.get("by", ""), "date": str(w.get("at", ""))[:10]} for w in (req.get("witnesses") or [])]


def actions(req):
    if req.get("actions"):
        return req["actions"]
    out = []
    for i, c in enumerate(req.get("capa") or []):
        st = {"verified": "Closed", "done": "Completed – awaiting verification"}.get(c.get("status"), "Open")
        out.append({"no": f"{i + 1:02d}", "cause": c.get("rootCause") or c.get("finding", ""), "action": c.get("action", ""),
                    "owner": c.get("owner", ""), "due": long(c.get("due")), "status": st})
    return out


# ---------- initial report ----------

def cmd_initial(req_json, out):
    req = load(req_json)
    R = refs(req)
    d = Doc(IC["templates"]["initial"])
    rdate = req.get("report_date") or str(dt.date.today())
    d.relabel(r"Ref\. ?No:", R["ini"])
    d.relabel(r"Date:", dmy(rdate, "-"))
    d.relabel(r"Project:", IC["project_title"])
    d.relabel(r"Location:", req.get("location", ""))
    d.relabel(r"Company:", IC["company_full"])
    d.relabel(r"Subject:", type_text(req))
    d.below("WHAT HAPPENED", T(req, "what", req.get("description", "")))
    d.below("WHEN DID IT HAPPEN", T(req, "when", when_text(req)))
    d.below("WHERE DID IT HAPPEN", T(req, "where", req.get("location", "")))
    d.below("WHO DID IT HAPPEN TO", T(req, "who", who_text(req)))
    d.below("HOW DID IT HAPPEN", T(req, "how", "To be established by the investigation."))
    d.below("RECOMMENDATIONS", T(req, "recommendations", [f"Immediate notification to the Consultant and the Engineer as per the WPS{req.get('wps') or 2} protocol."]))
    d.below("FURTHER COMMENTARY", T(req, "commentary", f"This incident is classified as WPS{req.get('wps') or 2} – {IC['severity_text'].get(str(req.get('wps') or 2), '')} as per the project HSE procedure. The final investigation report will follow."))
    # completed by (HSE engineer) / approved by (HSE manager): names and dates; the signature pictures go
    ppl = IC["people"]
    label_row(d, "Name", [ppl["hse_engineer"][0], ppl["hse_manager"][0]])
    d.blank_drawings_in(lambda labels, r: labels.startswith("SIGNATURE"))
    label_row(d, "Date", [dmy(rdate), dmy(rdate)])
    photo_grid(d, req.get("photos") or [])
    d.save(out)


def photo_grid(d, photos):
    """The 'Photos and Attachments' table: rows of (Description:, caption) × 2, then a picture row; rebuilt for photos."""
    tbls = [t for t in P.tables(d.x) if t and P.norm(t[0][0].text(d.x)).startswith("PHOTOS AND ATTACHMENTS")]
    if not tbls:
        d.notes.append("photo table not found")
        return
    t = tbls[0]
    cap_row = next((r for r in t if len(r) == 4 and P.norm(r[0].text(d.x)).startswith("DESCRIPTION")), None)
    img_row = next((r for r in t if len(r) == 2 and any("<w:drawing" in d.x[c.start:c.end] for c in r)), None)
    if not cap_row or not img_row:
        d.notes.append("photo table layout not recognised")
        return
    tr_cap = d.x[MR.row_at(d.x, cap_row[0].start):cap_row[0].start]
    tr_img = d.x[MR.row_at(d.x, img_row[0].start):img_row[0].start]
    cap_cells = [d.x[c.start:c.end] for c in cap_row]
    heads = [MR.first(r"<w:tc>|<w:tc [^>]*>", d.x[c.start:c.end]) + MR.first(r"<w:tcPr>.*?</w:tcPr>", d.x[c.start:c.end]) for c in img_row]
    rows = []
    for k in range(0, max(len(photos), 2), 2):
        pair = photos[k:k + 2] + [None] * (2 - len(photos[k:k + 2]))
        caps = "".join(MR.cell_text(cap_cells[2 * j], "Description:" if p else "") + MR.cell_text(cap_cells[2 * j + 1], p["caption"] if p else "") for j, p in enumerate(pair))
        pics = "".join(heads[j] + (f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{d.add_image(p["file"], 7.6, 6.0)}</w:p>' if p else "<w:p/>") + "</w:tc>"
                       for j, p in enumerate(pair))
        rows.append(tr_cap + caps + "</w:tr>" + tr_img + pics + "</w:tr>")
    s = MR.row_at(d.x, cap_row[0].start)
    e = d.x.find("</w:tr>", t[-1][-1].end) + len("</w:tr>")
    d.x = d.x[:s] + "".join(rows) + d.x[e:]


# ---------- investigation report ----------

def cmd_investigation(req_json, out):
    req = load(req_json)
    R = refs(req)
    d = Doc(IC["templates"]["investigation"])
    rdate = req.get("report_date") or str(dt.date.today())
    wps = int(req.get("wps") or 2)
    ppl = IC["people"]
    for label, value in (("Project", IC["project_no"]), ("Reference No", R["inv"]), ("Date of Report", dmy(rdate)), ("Incident Type", type_text(req)),
                         ("Date of Incident", dmy(req["date"])), ("Incident Severity", IC["severity_text"].get(str(wps), "")),
                         ("Work Package", CFG["package"]), ("Worst Potential Severity", f"WPS{wps}"), ("Contractor", CFG["contractor_short"]),
                         ("Site Supervision Consultant", CFG["consultant_short"]), ("HSE Manager", ppl["hse_manager"][0]),
                         ("SSC HSE Manager", IC["ssc_hse_manager"]), ("The Engineer", CFG["pmc"]), ("The Engineer HSE Manager", IC["engineer_hse_manager"]),
                         ("Client", CFG.get("client", "")), ("Date Authority Notified: [where applicable]", dmy(req.get("authority_notified")) or "N/A")):
        d.beside(label, value)
    d.below("BRIEF DESCRIPTION OF INCIDENT", T(req, "brief", req.get("description", "")))
    findings = list(T(req, "findings", [req.get("description", "")]))
    if T(req, "timeline"):
        findings += ["Timeline of Event:"] + list(T(req, "timeline"))
    d.below("INCIDENT FINDINGS", findings)
    d.below("RELEVANT DOCUMENTATION", T(req, "docs", (req.get("inv") or {}).get("docs") or "N/A"))
    d.below("EMERGENCY RESPONSE", T(req, "emergency", [req.get("immediateAction", "")]))
    comm = req.get("comm") or {}
    for label in ("Method Statement (MS) approved", "Activity covered in MS", "Risk Assessment (RA) Approved", "Activity covered in RA",
                  "Injured party briefed on MS/RA", "MS/RA revision required", "Hazards/Risks Identified", "Risk Controls Adequate", "Previous non-compliance noted"):
        d.beside(label, comm.get(label, ""))
    if not comm:
        d.notes.append("communication Y/N grid left blank: fill it (comm) or in Word")
    d.below("CONTRACTOR’S IMMEDIATE CORRECTIVE ACTION", T(req, "immediate_actions", [req.get("immediateAction", "")]))
    d.rows_after(["Name", "Department", "Company", "Statement Provided?"],
                 [[w["name"], w.get("department", ""), w.get("company", ""), "Yes" if w.get("statement") or w.get("photo") else "No"] for w in witnesses(req)] or [["N/A", "", "", ""]])
    inv = req.get("inv") or {}
    imm = T(req, "immediate_causes", [[c.get("factor", ""), c.get("detail", "")] for c in inv.get("immediate", [])])
    sysc = T(req, "system_causes", [[c.get("factor", ""), c.get("detail", "")] for c in inv.get("system", [])])
    # the form has two tables with the same header: immediate causes first, then system causes
    _two_cause_tables(d, imm, sysc)
    acts = actions(req)
    d.rows_after(["Causal Factor No. to be addressed", "Required Action", "Responsible Person", "Due Date"],
                 [[a.get("no", f"{i + 1:02d}"), a["action"], a.get("owner", ""), a.get("due", "")] for i, a in enumerate(acts)])
    if not acts:
        d.notes.append("no corrective actions: add CAPA in the page or 'actions' in the request")
    d.below("INCIDENT ANALYSIS", T(req, "analysis", inv.get("analysis") or ""))
    d.below("KEY LEARNING POINTS", T(req, "learning", [inv.get("learning")] if inv.get("learning") else []))
    d.below("PREVENTIVE ACTIONS", T(req, "preventive", []))
    # contractor signatories (names only; they sign the copy): the first Name/Position/Date rows are the contractor's
    for label, vals in (("Name", [ppl["hse_engineer"][0], ppl["hse_manager"][0], ppl["project_manager"][0]]),
                        ("Position", [ppl["hse_engineer"][1], ppl["hse_manager"][1], ppl["project_manager"][1]]),
                        ("Date", [dmy(rdate)] * 3)):
        _fill_first_row(d, label, vals)
    for label, yn in (("HSE Alert required", "Y" if wps >= 4 else "N"), ("Further investigation", "N"), ("Senior Management Briefing required?", "Y" if wps >= 3 else "N")):
        d.beside(label, yn)
    att = [f"{R['wit']}{'-' + str(i + 1).zfill(2) if len(witnesses(req)) > 1 else ''} - Witness Statement" for i in range(len(witnesses(req)))]
    d.below("ATTACHMENTS", att or ["Photographic evidence"])
    _annex(d, req.get("photos") or [])
    d.footer_replace("AVK01-ACTC-WP.03-HSE-INV-PD-00001", R["inv"])
    d.save(out)


def _two_cause_tables(d, imm, sysc):
    H = ["CAUSAL FACTOR", "DETAILED DESCRIPTION"]
    for k, rows in ((0, imm), (1, sysc)):
        found = -1
        for t in P.tables(d.x):
            for i, r in enumerate(t):
                if [P.norm(c.text(d.x)) for c in r] == H:
                    found += 1
                    if found == k:
                        data = []
                        for r2 in t[i + 1:]:
                            if len(r2) != 2 or [P.norm(c.text(d.x)) for c in r2] == H:
                                break
                            data.append(r2)
                        s = MR.row_at(d.x, data[0][0].start)
                        e = d.x.find("</w:tr>", data[-1][-1].end) + len("</w:tr>")
                        tr_pre = d.x[s:data[0][0].start]
                        cells = [d.x[c.start:c.end] for c in data[0]]
                        body = "".join(tr_pre + MR.cell_text(cells[0], a) + MR.cell_text(cells[1], b) + "</w:tr>" for a, b in (rows or [["", ""]]))
                        d.x = d.x[:s] + body + d.x[e:]
                        break
            if found >= k:
                break


def _fill_first_row(d, label, vals):
    """First row whose first cell is label (the contractor block): fill the cells after it."""
    for t in P.tables(d.x):
        for r in t:      # "Name:" with its colon: the witness table's "Name" header must not match
            if P.norm(r[0].text(d.x)) == label.upper() + ":" and len(r) > len(vals):
                for i, v in reversed(list(enumerate(vals))):
                    d.put(r[i + 1], v)
                return
    d.notes.append(f"signatory row not found: {label}")


def _annex(d, photos):
    """Everything after the last table (the house annex pages with text boxes) → a plain photographic annex."""
    body_end = d.x.rfind("<w:sectPr")
    last_tbl = d.x.rfind("</w:tbl>", 0, body_end) + len("</w:tbl>")
    tail = d.x[last_tbl:body_end]
    keep = "".join(re.sub(r"<w:r[ >].*?</w:r>", "", p, flags=re.S) for p in re.findall(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", tail, flags=re.S) if "<w:sectPr" in p)
    out = []
    if photos:
        out.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + SR.p("Appendix A – Photographic Evidence", size=26, bold=True, center=True))
        for i, ph in enumerate(photos):
            out.append(f'<w:p><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr>{d.add_image(ph["file"], 15.0, 10.0)}</w:p>')
            out.append(SR.p(f"Photo {i + 1} – {ph['caption']}", size=20, center=True, after=240))
    d.x = d.x[:last_tbl] + "".join(out) + keep + d.x[body_end:]


# ---------- witness statement ----------

def cmd_witness(req_json, n, out):
    req = load(req_json)
    R = refs(req)
    ws = witnesses(req)
    k = int(n) - 1
    if not 0 <= k < len(ws):
        raise SystemExit(f"witness {n} not found ({len(ws)} recorded)")
    w = ws[k]
    ref = R["wit"] + (f"-{k + 1:02d}" if len(ws) > 1 else "")
    d = Doc(IC["templates"]["witness"])
    d.relabel(r"Ref\. ?No:", ref)
    d.relabel(r"Date:", dmy(w.get("date") or req.get("report_date") or str(dt.date.today()), "-"))
    d.relabel(r"Project No\.:", IC["project_no"])
    d.relabel(r"Incident Reference:", R["ini"])
    d.relabel(r"Location:", req.get("location", ""))
    d.relabel(r"Incident Type:", type_text(req))
    d.below("WITNESS DETAILS", ", ".join(x for x in [w["name"], f"DOB: {w['dob']}" if w.get("dob") else "", w.get("company", ""), w.get("trade") or w.get("department", "")] if x))
    d.below("TRANSLATOR DETAILS", w.get("translator") or "N/A")
    d.below("STATEMENT RECORDED BY", f"Name: {w.get('recorded_by') or IC['people']['hse_engineer'][0]}")
    d.below("STATEMENT", w.get("statement") or "Signed statement attached.", exact=True)
    d.footer_replace("AVK01-ACTC-WP.03-HSE-WIT-PD-00001", ref)
    d.save(out)


# ---------- CAPA plan ----------

def cmd_capa(req_json, out):
    req = load(req_json)
    R = refs(req)
    doc_no = req.get("capa_doc") or next_capa_doc()
    d = Doc(IC["templates"]["capa"])
    rdate = req.get("report_date") or str(dt.date.today())
    for label, value in (("Incident Reference", R["ini"]), ("Document Number", doc_no), ("Project", IC["project_title"].replace("WP03", "WP-03")),
                         ("Rev No", "00"), ("Incident", f"{T(req, 'title', type_text(req))} ({IC['codes'].get(req.get('type'), '')})"),
                         ("Incident date", long(req["date"])), ("Contractor", CFG["contractor_short"]), ("CAPA date", long(rdate))):
        d.beside(label, value)
    acts = actions(req)
    d.rows_after(["Ref. No.", "Root Cause / Causal Factor Identified", "Corrective & Preventive Action Implemented", "Responsible Person", "Due Date", "Status"],
                 [[f"CAPA-{i + 1:02d}", a.get("cause", ""), a["action"], a.get("owner", ""), a.get("due", ""), a.get("status", "Open")] for i, a in enumerate(acts)])
    eff = T(req, "effectiveness", [])
    _replace_list_after(d, "To ensure the effectiveness", eff)
    if T(req, "closure"):
        closure = T(req, "closure")
        d.x = MR.replace_paras(d.x, lambda s: s.strip().startswith("Ahmadiah Contracting"), lambda s: closure)
    else:
        d.notes.append("closure statement kept from the house copy: write 'closure' for this incident")
    d.blank_drawings_in(lambda labels, r: any("<w:drawing" in d.x[c.start:c.end] for c in r))
    d.footer_replace("AVK01-WP.03-HSE-CAPA-002", doc_no)
    d.save(out)


def _replace_list_after(d, intro_start, items):
    """The bullet paragraphs after the intro paragraph (up to the next numbered heading) → items."""
    ps = list(re.finditer(r"<w:p[ >](?:(?!</w:p>|<w:p[ >]).)*?</w:p>", d.x, flags=re.S))
    i = next((k for k, m in enumerate(ps) if MR.text_of(m.group(0)).strip().startswith(intro_start)), None)
    if i is None:
        d.notes.append(f"paragraph not found: {intro_start}")
        return
    j = i + 1
    while j < len(ps) and not re.match(r"\s*\d+\.\s", MR.text_of(ps[j].group(0))) and MR.text_of(ps[j].group(0)).strip():
        j += 1
    if j == i + 1 or not items:
        if not items:
            d.notes.append("effectiveness verification kept from the house copy: write 'effectiveness' for this incident")
        return
    p0 = ps[i + 1].group(0)
    ppr = MR.first(r"<w:pPr>.*?</w:pPr>", p0)
    rpr = MR.first(r"<w:rPr>.*?</w:rPr>", MR.first(r"<w:r[ >].*?</w:r>", p0))
    new = "".join(MR.para(ppr, rpr, it) for it in items)
    d.x = d.x[:ps[i + 1].start()] + new + d.x[ps[j - 1].end():]


# ---------- safety alert (FRM-00015, one portrait slide) ----------

BRAND = "08475E"


def _sp(sid, name, x, y, cx, cy, paras):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')


def _rpr(sz, bold=False, color="262626"):
    f = '<a:latin typeface="Source Sans Pro"/><a:cs typeface="Arial"/>'
    return f'<a:rPr lang="en-GB" sz="{sz}" b="{1 if bold else 0}" dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{f}</a:rPr>'


def _ap(text, sz=1050, bold=False, color="262626", bullet=False, after=300):
    ppr = (f'<a:pPr marL="171450" indent="-171450"><a:spcAft><a:spcPts val="{after}"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>' if bullet
           else f'<a:pPr><a:spcAft><a:spcPts val="{after}"/></a:spcAft><a:buNone/></a:pPr>')
    return f'<a:p>{ppr}<a:r>{_rpr(sz, bold, color)}<a:t>{escape(str(text))}</a:t></a:r></a:p>'


def _pic(sid, rid, x, y, cx, cy):
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="Alert photo {sid}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></p:spPr></p:pic>')


def cmd_alert(req_json, out):
    """Safety alert for WPS4-5 (and when the Engineer asks): what happened, photos, why, actions, key messages."""
    req = load(req_json)
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    tpl = cfg_path(IC["templates"]["alert"])
    z = zipfile.ZipFile(tpl)
    files = {n: z.read(n) for n in z.namelist()}
    s = files["ppt/slides/slide1.xml"].decode("utf-8")
    rdate = req.get("report_date") or str(dt.date.today())
    ref = req.get("alert_ref") or f"{IC['ref_prefix']}SA-{req['ref_no']}"
    title = T(req, "alert_title", type_text(req))
    notes = []
    if len(title) > 45:
        notes.append(f"title is {len(title)} characters: the bar holds about 45")
    ppl = IC["people"]
    # the two bars: Ref # / Title / Date and Originator / Approved By / Date
    i = s.find("Ref #")
    a, b = s.rfind("<p:txBody>", 0, i), s.find("</p:txBody>", i) + len("</p:txBody>")
    s = s[:a] + _bar_body(s[a:b], [f"Ref # {ref}", f"Title: {title}", f"Date: {dmy(rdate)}"], [2950000, 5600000], 900) + s[b:]
    i = s.find("Originator:")
    a, b = s.rfind("<p:txBody>", 0, i), s.find("</p:txBody>", i) + len("</p:txBody>")
    s = s[:a] + _bar_body(s[a:b], [f"Originator: {ppl['hse_engineer'][0]} ({CFG['contractor_short']} HSE)", f"Approved By: {ppl['hse_manager'][0]}", f"Date: {dmy(rdate)}"],
                          [3150000, 5600000], 900) + s[b:]
    # body: what happened, photos, why, actions, key messages
    X, W = 342900, 6172200
    top, bottom = 1330000, 7950000
    what = T(req, "what", req.get("description", ""))
    why = T(req, "why", [f"{a2}: {b2}" if a2 and b2 else a2 or b2 for a2, b2 in
                         T(req, "immediate_causes", [[c.get("factor", ""), c.get("detail", "")] for c in (req.get("inv") or {}).get("immediate", [])])])
    acts = T(req, "actions_taken", [a2["action"] for a2 in actions(req)])
    lessons = T(req, "lessons", T(req, "learning", []))
    photos = (req.get("photos") or [])[:2]
    sid = 100
    shapes = []
    hA = 1150000
    shapes.append(_sp(sid, "What happened", X, top, W, hA, [_ap("WHAT HAPPENED", 1200, True, BRAND), _ap(what)]))
    sid += 1
    y = top + hA + 60000
    rels = files["ppt/slides/_rels/slide1.xml.rels"].decode("utf-8")
    if photos:
        boxh = 2250000
        boxw = W if len(photos) == 1 else (W - 172200) // 2
        for k, ph in enumerate(photos):
            data, cx, cy = SR.prep_image(ph["file"], boxw / 360000, boxh / 360000)
            name = f"media/alert_{k + 1}.jpeg"
            files["ppt/" + name] = data
            rid = f"rIdAlert{k + 1}"
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../{name}"/></Relationships>')
            bx = X + k * (boxw + 172200)
            shapes.append(_pic(sid, rid, bx + (boxw - cx) // 2, y + (boxh - cy) // 2, cx, cy)); sid += 1
            shapes.append(_sp(sid, f"Caption {k + 1}", bx, y + boxh + 20000, boxw, 200000, [_ap(f"Photo {k + 1} – {ph['caption']}", 900, False, "595959", after=0)])); sid += 1
        y += boxh + 260000
        ct = files["[Content_Types].xml"].decode("utf-8")
        if 'Extension="jpeg"' not in ct:
            files["[Content_Types].xml"] = re.sub(r"(<Types [^>]*>)", r'\1<Default Extension="jpeg" ContentType="image/jpeg"/>', ct, count=1).encode("utf-8")
    files["ppt/slides/_rels/slide1.xml.rels"] = rels.encode("utf-8")
    paras = []
    for head, items in (("WHY IT HAPPENED", why), ("ACTIONS TAKEN", acts), ("KEY MESSAGES", lessons)):
        if items:
            paras.append(_ap(head, 1200, True, BRAND, after=200))
            paras += [_ap(t, bullet=True, after=200) for t in items]
    shapes.append(_sp(sid, "Causes and actions", X, y, W, bottom - y, paras))
    # rough fit check: about 92 characters a line at 10.5 pt across the box, 0.19 inch a line
    inch = lambda emu: emu / 914400
    if -(-len(str(what)) // 92) + 1.4 > inch(hA) / 0.19:
        notes.append("'what' is too long for its box: shorten it")
    lines_b = sum(1.4 for it in (why, acts, lessons) if it) + sum(-(-len(str(t)) // 88) for it in (why, acts, lessons) for t in it)
    if lines_b > inch(bottom - y) / 0.19:
        notes.append("causes, actions and messages overflow the page: shorten them or use one photo")
    if not why:
        notes.append("no causes: the investigation is not written yet")
    s = s.replace("</p:spTree>", "".join(shapes) + "</p:spTree>")
    left = [m for m in PLACEHOLDER.findall(MR.text_of(s.replace("<a:t>", "<w:t>").replace("</a:t>", "</w:t>"))) if m.strip("[]").strip()]
    if left:
        raise SystemExit("Not written: placeholder text left: " + ", ".join(sorted(set(left))))
    files["ppt/slides/slide1.xml"] = s.encode("utf-8")
    import xml.dom.minidom
    for n in ("ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels", "[Content_Types].xml"):
        xml.dom.minidom.parseString(files[n])
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, b2 in files.items():
            zo.writestr(n, b2)
    print(json.dumps({"written": out, "ref": ref, "notes": notes}, ensure_ascii=False, indent=1))


def _bar_body(txbody, parts, tabs, sz):
    """Replace the single paragraph of a bar's <p:txBody> with the parts at tab stops (white, bold, as the form)."""
    tl = "".join(f'<a:tab pos="{p}" algn="l"/>' for p in tabs)
    runs = "".join(f'<a:r>{_rpr(sz, True, "FFFFFF")}<a:t>{escape(("" if i == 0 else chr(9)) + t)}</a:t></a:r>' for i, t in enumerate(parts))
    return re.sub(r"<a:p>.*</a:p>", lambda _: f'<a:p><a:pPr algn="l"><a:tabLst>{tl}</a:tabLst></a:pPr>{runs}</a:p>', txbody, count=1, flags=re.S)


# ---------- registers (rows for incident_register.ps1) ----------

def _reg013_last():
    import openpyxl
    ws = openpyxl.load_workbook(IC["register_013"], data_only=True)["Incident Details"]
    nos = [r[0] for r in ws.iter_rows(min_row=3, values_only=True) if isinstance(r[0], (int, float))]
    last_g = None          # the contractor as the register writes it (e.g. "Ahmadiah- Road Improvement")
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[6]:
            last_g = r[6]
    return int(max(nos)) if nos else 0, last_g or CFG["contractor_short"]


def cmd_regrows(req_json, stage, out):
    """Rows for REG-013 (stages: initial, investigation, closed) or REG-014 (stage: minor)."""
    req = load(req_json)
    R = refs(req) if req.get("ref_no") else {}
    wps = int(req.get("wps") or 2)
    rdate = req.get("report_date") or str(dt.date.today())
    acts = actions(req)
    inv = req.get("inv") or {}
    causes = T(req, "immediate_causes", [[c.get("factor", ""), c.get("detail", "")] for c in inv.get("immediate", [])]) + \
        T(req, "system_causes", [[c.get("factor", ""), c.get("detail", "")] for c in inv.get("system", [])])
    cause_text = "\n".join(f"{a}: {b}" if a and b else a or b for a, b in causes)
    act_text = "\n".join(f"{i + 1}. {a['action']}" for i, a in enumerate(acts))
    if stage == "initial":
        last, contractor = _reg013_last()
        no = req.get("reg_no") or last + 1
        rows = {"register": "013", "append": [{"A": str(no), "B": "date:" + req["date"], "C": IC["register_type"].get(req.get("type"), type_text(req)),
                "D": f"WPS{wps}", "E": CFG["pmc"], "F": CFG["consultant_short"], "G": contractor, "H": CFG["package"],
                "I": T(req, "brief", req.get("description", "")), "J": "Yes" if req.get("notifiedAt") else "No", "K": "date:" + rdate, "P": "Open"}]}
    elif stage == "investigation":
        rows = {"register": "013", "update": [{"key": str(req["reg_no"]), "set": {"L": "date:" + rdate, "M": cause_text, "N": act_text}}]}
    elif stage == "closed":
        rows = {"register": "013", "update": [{"key": str(req["reg_no"]), "set": {"O": "date:" + str(req.get("closedAt") or rdate)[:10], "P": "Closed"}}]}
    elif stage == "minor":
        tp = req.get("thirdParty", "")
        at = req.get("notifiedAt")
        rep = ""
        if at:
            na = dt.datetime.fromisoformat(at.replace("Z", "+00:00")).astimezone(dt.timezone(dt.timedelta(hours=CFG["trend"].get("utc_offset_hours", 3))))
            rep = na.strftime("%Y-%m-%d %H:%M")
        first = acts[0] if acts else {}
        row = {"A": req.get("minor_ref") or f"Log - {dmy(req['date'], '-')} {req.get('time', '')}".strip(),      # text, so Excel keeps it "B": "date:" + req["date"], "C": req.get("time", ""),
               "D": req.get("location", ""), "F": T(req, "event_type", type_text(req)), "G": T(req, "brief", req.get("description", "")),
               "H": "Yes" if any(p.get("injury") for p in req.get("persons") or []) else "No", "I": T(req, "damage", "No"),
               "J": "Yes" if tp else "No", "K": T(req, "plate", ""), "L": "None", "M": "SSC + Engineer" if at else "", "N": T(req, "method", "Phone + Email" if at else ""),
               "O": rep, "P": req.get("reporter", ""), "Q": T(req, "immediate_actions", [req.get("immediateAction", "")])[0],
               "R": cause_text, "S": first.get("action", ""), "T": first.get("owner", ""), "U": first.get("due", ""),
               "V": T(req, "action_by", ""), "W": "", "X": "Closed" if req.get("status") == "closed" else "Open",
               "Y": "date:" + str(req["closedAt"])[:10] if req.get("closedAt") else "", "Z": "N/A", "AA": T(req, "photo_ref", ""), "AB": ""}
        rows = {"register": "014", "append": [row]}
    else:
        raise SystemExit("stage: initial | investigation | closed | minor")
    json.dump(rows, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(rows, ensure_ascii=False, indent=1))


CAPA_SOURCE_TEXT = {"audit_internal": "Internal Audit", "audit_pace": "PACE Audit", "audit_jll": "JLL Audit", "mgmt": "Management Review", "incident": "Incident Investigation"}


def cmd_capa_rows(capa_dir, inc_dir, out):
    """REG-011 rows for CAPA documents from the page: new ones appended, changed ones updated (by their Finding ID).
    Column C carries the register's source names, which its dashboard counts; the source reference goes in M."""
    from weekly_hse import load as wload
    caps = wload(os.path.dirname(capa_dir.rstrip("\\/")), os.path.basename(capa_dir.rstrip("\\/")))
    incs = {x["_id"]: x for x in wload(os.path.dirname(inc_dir.rstrip("\\/")), os.path.basename(inc_dir.rstrip("\\/")))}
    caps = [c for c in caps if not c.get("example")]
    caps.sort(key=lambda c: c.get("createdAt", ""))
    seq, append, update, ids, skipped = {}, [], [], {}, []
    year = dt.date.today().year
    status = lambda c: {"verified": "Closed - Verified", "done": "In Progress"}.get(c.get("status"), "Open")
    for c in caps:
        src = c.get("source", "other")
        key = (src, c.get("sourceId") or c.get("sourceRef") or "")
        seq[key] = seq.get(key, 0) + 1
        k = seq[key]
        if c.get("regId"):
            if c.get("regStatus") != c.get("status"):
                update.append({"key": c["regId"], "set": _capa_cells(c, status, only_close=True)})
                ids[c["_id"]] = c["regId"]
            continue
        if src == "incident":
            x = incs.get(c.get("sourceId"))
            ini = ((x or {}).get("ini") or {}).get("ref", "")
            m = re.search(r"INI-([A-Z]+)-(\d{5})", ini)
            if not m:
                skipped.append(f"{c['_id']}: its incident has no initial report number yet")
                continue
            rid, ref = f"INV-{m.group(1)}-{m.group(2)}-{k:02d}", ini
        elif src in ("svn", "ncr", "sor", "ltr"):
            ref = c.get("sourceRef", "").split(" — ")[0]
            rid = f"{ref}-{k:02d}"
        else:
            ref = c.get("sourceRef", "")
            rid = ref if ref and k == 1 else f"CAPA-{year}-{len(append) + 1:03d}"
        row = {"A": rid, "B": "date:" + str(c.get("createdAt", ""))[:10], "C": CAPA_SOURCE_TEXT.get(src, "Other"), "D": c.get("finding", ""),
               "E": c.get("rootCause", ""), "F": c.get("action", ""), "G": c.get("owner", ""), "H": "date:" + c["due"] if c.get("due") else ""}
        row.update(_capa_cells(c, status))
        row["M"] = "; ".join(x for x in [f"Ref: {ref}" if ref else "", row.get("M", "")] if x)
        append.append(row)
        ids[c["_id"]] = rid
    rows = {"register": "011", "append": append, "update": update}
    json.dump(rows, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"append": len(append), "update": len(update), "ids": ids, "skipped": skipped, "rows": out}, ensure_ascii=False, indent=1))


def _capa_cells(c, status, only_close=False):
    s = {"I": status(c)}
    when = c.get("verifiedAt") or c.get("doneAt")
    if when:
        s["J"] = "date:" + str(when)[:10]
    if c.get("status") == "verified":
        s["K"] = "1"
        role = {"admin": "", "hse_manager": "HSE Manager", "hse_officer": "HSE Engineer"}.get(c.get("verifiedRole"), "")
        s["L"] = f"{c.get('verifiedBy', '')}" + (f", {role}" if role else "")
    if c.get("doneNote"):
        s["M"] = c["doneNote"]
    return s


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    {"plan": cmd_plan, "initial": cmd_initial, "investigation": cmd_investigation, "witness": cmd_witness, "capa": cmd_capa,
     "alert": cmd_alert, "regrows": cmd_regrows, "capa-rows": cmd_capa_rows}[sys.argv[1]](*sys.argv[2:])

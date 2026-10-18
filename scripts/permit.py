# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Permits to work on the company's own Word forms (setup in permits.json).

    python permit.py settings <out.json>               settings/permits document for the page (types, hazards, precautions per form)
    python permit.py plan <permits export dir> --approvals <approvals export dir> --roles <roles export dir> [--week YYYY-MM-DD] [--base main=139 exc=49]
                                                        next numbers for approved requests, from the register and issued permits
    python permit.py make <request.json> <out.docx>     fill the type's form; request = the permit document plus "permitNo"

The forms are filled by their labels (DATE, PERMIT NO, FROM, TO, LOCATION, ...) and checklist rows (item, YES, NO, N/A),
so a revised form keeps working as long as its wording stays. Never overwrites a file.
"""
import datetime as dt, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
from config import CFG, ROOT

PJ = json.load(open(os.path.join(ROOT, "permits.json"), encoding="utf-8"))
NON_LATIN = re.compile(r"[؀-ۿ]+")


def norm(s):
    s = NON_LATIN.sub(" ", str(s or "")).replace(" ", " ")
    s = re.sub(r"\(\s*\)", " ", s)
    return re.sub(r"\s+", " ", s).strip().upper()


# ---------- Word table reader ----------

TAG = re.compile(r"<(/?)w:(tbl|tr|tc)(?=[\s>/])[^>]*?(/?)>")
BOX = re.compile(r'<w:sym w:font="Wingdings" w:char="F06F"/>')                 # the excavation form's empty box
CHOICE = re.compile(r"^(STABLE ?ROCK|TYPE [ABC]|STAIRWAY|LADDERS|RAMP|R?P?E? ?[A-Z]{1,3} –)")      # options, not checklist items


class Cell:
    def __init__(self, start):
        self.start, self.end, self.nested, self.span, self.col, self.vmerge = start, None, False, 1, 0, False

    def text(self, xml):
        return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml[self.start:self.end])).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")


def tables(xml):
    """[[ [Cell] per row ] per table] for every table, nested ones included; cells know whether they hold a table."""
    out, stack = [], []          # stack of ("tbl", rows) / ("tr", cells) / ("tc", Cell)
    for m in TAG.finditer(xml):
        close, kind = m.group(1) == "/", m.group(2)
        if not close:
            if kind == "tbl":
                for k, c in stack:
                    if k == "tc":
                        c.nested = True
                stack.append(("tbl", []))
            elif kind == "tr":
                stack.append(("tr", []))
            else:
                stack.append(("tc", Cell(m.start())))
        else:
            k, obj = stack.pop()
            if k == "tc":
                obj.end = m.end()
                body = xml[obj.start:obj.end]
                sp = re.search(r'<w:gridSpan w:val="(\d+)"', body[:body.find("</w:tcPr>") + 1] if "</w:tcPr>" in body else "")
                obj.span = int(sp.group(1)) if sp else 1
                obj.vmerge = "<w:vMerge" in body[:body.find("</w:tcPr>") + 1] if "</w:tcPr>" in body else False
                stack[-1][1].append(obj)
            elif k == "tr":
                col = 0
                for c in obj:
                    c.col, col = col, col + c.span
                stack[-1][1].append(obj)
            else:
                out.append(obj)
    return out


def run(text, size=18, bold=True, symbol=False):
    font = '<w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol"/>' if symbol else '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    return (f'<w:r><w:rPr>{font}{"<w:b/>" if bold else ""}<w:color w:val="1F3864"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
            f'<w:t xml:space="preserve">{escape(str(text))}</w:t></w:r>')


class Filler:
    def __init__(self, xml):
        self.xml, self.edits, self.used = xml, [], set()
        self.tabs = tables(xml)
        self.cells = [c for t in self.tabs for r in t for c in r if not c.nested]

    def put(self, cell, value, **kw):
        if value in (None, ""):
            return
        body = self.xml[cell.start:cell.end]
        sep = "" if kw.pop("tight", False) or not cell.text(self.xml).strip() else " "
        r = run(sep + str(value), **kw)
        i = body.rfind("</w:p>")
        if i >= 0:
            self.edits.append((cell.start + i, r, 0))
        else:                                   # empty paragraph written as <w:p/> or <w:p w:rsid.../>
            m = list(re.finditer(r"<w:p(\s[^>]*)?/>", body))
            if not m:
                return
            m = m[-1]
            self.edits.append((cell.start + m.start(), f"<w:p{m.group(1) or ''}>{r}</w:p>", m.end() - m.start()))
        self.used.add(cell.start)

    def label(self, pattern, value, nth=0, **kw):
        """Append value to the nth cell whose text matches pattern."""
        hits = [c for c in self.cells if re.search(pattern, norm(c.text(self.xml)))]
        if len(hits) > nth:
            self.put(hits[nth], value, **kw)
            return True
        return False

    def label_or_next(self, pattern, value, **kw):
        """Like label(), but into the empty cell right after the label when the form has one (confined-space 'Permit No.:')."""
        for t in self.tabs:
            for r in t:
                for i, c in enumerate(r):
                    if not c.nested and re.search(pattern, norm(c.text(self.xml))):
                        if i + 1 < len(r) and not r[i + 1].nested and not norm(r[i + 1].text(self.xml)):
                            self.put(r[i + 1], value, **kw)
                        else:
                            self.put(c, value, **kw)
                        return True
        return False

    def after(self, pattern, value, **kw):
        """Write value into the cell right after the first cell matching pattern (value cells beside a label)."""
        for t in self.tabs:
            for r in t:
                for i, c in enumerate(r):
                    if not c.nested and re.search(pattern, norm(c.text(self.xml))) and i + 1 < len(r):
                        self.put(r[i + 1], value, **kw)
                        return True
        return False

    def rows_under(self, header_words):
        """Rows below the header row holding all header_words, as lists of non-nested cells (header cells' positions)."""
        for t in self.tabs:
            for ri, r in enumerate(t):
                texts = [norm(c.text(self.xml)) for c in r]
                pos = []
                for w in header_words:
                    j = next((k for k, x in enumerate(texts) if x.startswith(w) and k not in pos), None)
                    if j is None:
                        break
                    pos.append(j)
                else:
                    cols = [r[j].col for j in pos]
                    out = []
                    for r2 in t[ri + 1:]:
                        by_col = {c.col: c for c in r2}
                        if all(cl in by_col for cl in cols):
                            out.append([by_col[cl] for cl in cols])
                    return out
        return []

    def tick_inline(self, cell, a):
        """A 'Yes ☐ No ☐ N/A ☐' cell (Wingdings boxes after each word, as on the confined-space form): tick the chosen box."""
        body = self.xml[cell.start:cell.end]
        boxes = list(re.finditer(r'<w:sym w:font="Wingdings" w:char="F071"/>', body))
        idx = {"yes": 0, "no": 1, "na": 2}.get(a)
        if idx is None or idx >= len(boxes):
            return False
        m = boxes[idx]
        self.edits.append((cell.start + m.start(), '<w:sym w:font="Wingdings" w:char="F0FE"/>', m.end() - m.start()))
        return True

    def tick_box(self, cell, idx=0):
        """Tick the idx-th empty Wingdings box in a cell (excavation form: the box sits in the cell before the item)."""
        body = self.xml[cell.start:cell.end]
        boxes = list(re.finditer(r'<w:sym w:font="Wingdings" w:char="F0(6F|71|A8)"/>', body))
        if idx >= len(boxes):
            return False
        m = boxes[idx]
        self.edits.append((cell.start + m.start(), '<w:sym w:font="Wingdings" w:char="F0FE"/>', m.end() - m.start()))
        return True

    def blank(self, pattern, value):
        """Write value over the ____ line of the cell whose text matches pattern (e.g. 'DEPTH ______')."""
        if value in (None, ""):
            return False
        for c in self.cells:
            if re.search(pattern, norm(c.text(self.xml))):
                body = self.xml[c.start:c.end]
                # the line is often split over several runs: the first stretch takes the value, the rest is removed
                parts = [m for m in re.finditer(r">([^<]*_[^<]*)<", body) if "_" in m.group(1)]
                if parts:
                    main = next((n for n, m in enumerate(parts) if "__" in m.group(1)), 0)
                    for n, m in enumerate(parts):
                        txt = re.sub(r"_+", f" {escape(str(value))} ", m.group(1), count=1) if n == main else m.group(1)
                        txt = txt.replace("_", "")
                        self.edits.append((c.start + m.start(1), txt, m.end(1) - m.start(1)))
                    return True
        return False

    def tick_choice(self, pattern, idx=0):
        """Tick the box in the cell just before the cell whose text matches pattern."""
        for t in self.tabs:
            for r in t:
                for i, c in enumerate(r):
                    if i and not c.nested and re.search(pattern, norm(c.text(self.xml))):
                        return self.tick_box(r[i - 1], idx)
        return False

    def checklist(self, answers):
        """answers: {normalised item: 'yes'|'no'|'na'}; marks the YES/NO/N/A cell after each item cell."""
        done = set()
        for t in self.tabs:
            for r in t:
                for i, c in enumerate(r):
                    k = norm(c.text(self.xml))
                    if k in answers and k not in done and not c.nested and i and BOX.search(self.xml[r[i - 1].start:r[i - 1].end]) \
                            and not (i + 1 < len(r) and norm(r[i + 1].text(self.xml)) in ("YES", "YES NO N/A")):
                        if answers[k] == "yes":                                    # box before the item: ticked = done
                            self.tick_box(r[i - 1])
                        done.add(k)
                        continue
                    if k in answers and k not in done and not c.nested and i + 1 < len(r) and norm(r[i + 1].text(self.xml)) == "YES NO N/A":
                        if self.tick_inline(r[i + 1], answers[k]):
                            done.add(k)
                        continue
                    if k in answers and k not in done and i + 3 < len(r) + 0 and not c.nested:
                        a = answers[k]
                        idx = {"yes": 1, "no": 2, "na": 3}.get(a)
                        if idx and i + idx < len(r):
                            self.put(r[i + idx], "✓", symbol=True, bold=False, size=20, tight=True)
                            done.add(k)
        return done

    def result(self):
        x = self.xml
        for pos, ins, cut in sorted(self.edits, key=lambda e: -e[0]):
            x = x[:pos] + ins + x[pos + cut:]
        return x


def form_items(path):
    """Hazards and precautions of a form, by the grid column of their list header."""
    x = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    F = Filler(x)
    out = {"hazards": [], "precautions": []}
    for t in F.tabs:
        heads = {}
        for r in t:
            for i, c in enumerate(r):
                k = norm(c.text(x))
                if k in ("HAZARDS", "SAFETY PRCAUTIONS", "SAFETY PRECAUTIONS"):
                    heads[c.col] = "hazards" if k == "HAZARDS" else "precautions"
            for i, c in enumerate(r):
                k = norm(c.text(x))
                if not c.nested and k and i + 1 < len(r) and norm(r[i + 1].text(x)) == "YES NO N/A":     # confined-space style
                    if k not in out["precautions"]:
                        out["precautions"].append(k)
                    continue
                if i and not c.nested and k and BOX.search(x[r[i - 1].start:r[i - 1].end]) and not CHOICE.search(k):   # excavation style
                    if k not in out["precautions"]:
                        out["precautions"].append(k)
                    continue
                if len(r) == 1 and k.startswith("OTHER PRECAUTIONS"):                                  # free-text row
                    if k + ":" not in out["precautions"]:
                        out["precautions"].append(k + ":")
                    continue
                if i == 0 or c.nested or i + 3 >= len(r) + 0:
                    continue
                sn, k = norm(r[i - 1].text(x)), norm(c.text(x))
                if re.fullmatch(r"\d{1,2}", sn) and k and all(not norm(r[i + j].text(x)) for j in (1, 2, 3) if i + j < len(r)):
                    lst = heads.get(c.col)
                    if lst and k not in out[lst]:
                        out[lst].append(k)
    return out


def cmd_settings(out):
    types, notes = [], []
    for ty in PJ["types"]:
        d = {k: ty[k] for k in ("key", "en", "ar", "register", "series", "extras")}
        d["noIsFlag"] = ty.get("noIsFlag", True)        # false where the form asks questions that may rightly be answered No
        if os.path.exists(ty["form"]):
            items = form_items(ty["form"])
            ar = {norm(k): v for k, v in PJ["ar"].items()}
            for lst in ("hazards", "precautions"):
                d[lst] = [{"en": k, "ar": ar.get(k, "")} for k in items[lst]]
            d["formReady"] = True
            notes.append({ty["key"]: {lst: len(d[lst]) for lst in ("hazards", "precautions")},
                          "no_arabic": [i["en"] for lst in ("hazards", "precautions") for i in d[lst] if not i["ar"]]})
        else:
            d["hazards"], d["precautions"], d["formReady"] = [], [], False
            notes.append({ty["key"]: "form missing: " + ty["form"]})
        types.append(d)
    doc = {"types": types, "shifts": PJ["shifts"], "weekStarts": PJ.get("week_starts", "Saturday")}
    json.dump(doc, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"written": out, "types": notes}, ensure_ascii=False, indent=1))


# ---------- numbering ----------

def week_start(day):
    return day - dt.timedelta(days=(day.weekday() - 5) % 7)          # Saturday


def parse_no(s):
    m = re.match(r"^\s*0*(\d+)\s*[-\s]?\s*([A-Z])?\s*$", str(s or "").upper())
    return (int(m.group(1)), m.group(2) or "") if m else (None, "")


def register_rows():
    import openpyxl
    wb = openpyxl.load_workbook(PJ["register"], data_only=True)
    rows = []
    for ws in wb.worksheets:
        for r in ws.iter_rows(min_row=4, values_only=True):
            if r[1] and r[2] is not None:
                rows.append({"sheet": ws.title, "type": str(r[1]).strip().upper(), "no": str(r[2]).strip(), "date": r[3]})
    return rows


def series_of(register_label):
    for ty in PJ["types"]:
        if ty["register"].upper() == register_label.upper():
            return ty["series"]
    return "exc" if "EXCAV" in register_label.upper() else "main"


def cmd_plan(export_dir, *args):
    a = dict(zip(args[::2], args[1::2]))
    week = week_start(dt.date.fromisoformat(a["--week"]) if "--week" in a else dt.date.today())
    base_arg = dict(x.split("=") for x in a.get("--base", "").split()) if a.get("--base") else {}
    docs = []
    for dp, _, fs in os.walk(export_dir):
        for f in fs:
            if f.endswith(".json"):
                d = json.load(open(os.path.join(dp, f), encoding="utf-8-sig"))
                d = d.get("data", d)
                d["_id"] = f[:-5]
                docs.append(d)
    reg = register_rows()
    top = {}
    for r in reg:
        n, _ = parse_no(r["no"])
        if n is not None:
            s = series_of(r["type"])
            top[s] = max(top.get(s, 0), n)
    types = {t["key"]: t for t in PJ["types"]}
    issued = [d for d in docs if d.get("permitNo")]
    plan, used = [], {}
    for d in issued:
        if d.get("weekStart") == str(week):
            n, L = parse_no(d["permitNo"])
            s = types[d["type"]]["series"]
            used.setdefault(s, {"n": n, "letters": set()})["letters"].add(L)
    for s in ("main", "exc"):
        if s not in used:
            if s in base_arg:
                used[s] = {"n": int(base_arg[s]), "letters": set()}
            else:
                used[s] = {"n": top.get(s, 0) + 1, "letters": set(), "guess": True}
    pending = [d for d in docs if d.get("status") == "approved" and not d.get("permitNo")]
    # the approval itself must exist in approvals/ (only Editors can write there; a permit's own status field can be
    # written by any Contributor), with the decision "approved"
    # Approvals live at approvals/<approver's id>/items/permit-<id>: the server lets each person write only under their
    # own id, so the folder name is who approved; that person's role comes from roles/ (only the owner writes it).
    unproven = []
    if "--approvals" in a:
        if "--roles" not in a:
            raise SystemExit("Pass --roles <roles export dir> too: the approver's role is checked there.")
        roles = {}
        for f in os.listdir(a["--roles"]):
            if f.endswith(".json"):
                x = json.load(open(os.path.join(a["--roles"], f), encoding="utf-8-sig"))
                roles[f[:-5]] = (x.get("data", x) or {}).get("role")
        appr = {}
        for dp, _, fs in os.walk(a["--approvals"]):
            for f in fs:
                if f.endswith(".json") and os.path.basename(dp) == "items":
                    x = json.load(open(os.path.join(dp, f), encoding="utf-8-sig"))
                    who = os.path.basename(os.path.dirname(dp))
                    appr.setdefault(f[:-5], []).append((who, x.get("data", x)))
        staff = ("admin", "hse_manager", "hse_officer")
        for d in pending:
            ok = [(who, r) for who, r in appr.get("permit-" + d["_id"], []) if r.get("decision") == "approved" and roles.get(who) in staff]
            if not ok:
                unproven.append({"id": d["_id"], "type": d.get("type"), "location": d.get("location"), "approvedBy": d.get("approvedBy"),
                                 "records": [{"by": r.get("by"), "id": who, "role_now": roles.get(who) or "supervisor"} for who, r in appr.get("permit-" + d["_id"], [])]})
            else:
                who, r = ok[-1]
                d["_approval"] = {"by": r.get("by"), "id": who, "role": roles.get(who), "at": r.get("at")}
        pending = [d for d in pending if "_approval" in d]
    elif "--no-approvals-check" not in args:
        raise SystemExit("Pass --approvals <dir with approvals/<id>/items exports> --roles <roles export dir>; --no-approvals-check only for old requests.")
    later = [{"id": d["_id"], "type": d["type"], "weekStart": d.get("weekStart")} for d in pending if d.get("weekStart") and d["weekStart"] != str(week)]
    for d in sorted([d for d in pending if not d.get("weekStart") or d["weekStart"] == str(week)], key=lambda d: d.get("createdAt", "")):
        s = types[d["type"]]["series"]
        u = used[s]
        L = next(c for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if c not in u["letters"])
        u["letters"].add(L)
        plan.append({"id": d["_id"], "type": d["type"], "location": d.get("location"), "permitNo": PJ["number_format"].format(n=u["n"], L=L),
                     "weekStart": str(week), "weekEnd": str(week + dt.timedelta(days=6)), "approval": d.get("_approval")})
    warn = [f"series {s}: base {u['n']} guessed from the register (its highest number + 1); the register may be behind, so confirm or pass --base {s}=N"
            for s, u in used.items() if u.get("guess")]
    if later:
        warn.append("approved requests for another week (run plan with --week for that week): " + json.dumps(later))
    if unproven:
        warn.append("marked approved but with no approval written by someone whose role (in roles/) is HSE staff, so not numbered: " + json.dumps(unproven, ensure_ascii=False))
    print(json.dumps({"week": f"{week} to {week + dt.timedelta(days=6)}", "register_top": top, "plan": plan, "warnings": warn}, ensure_ascii=False, indent=1))


# ---------- fill a form ----------

def dmy(s):
    s = str(s or "")
    return f"{s[8:10]}-{s[5:7]}-{s[:4]}" if re.match(r"^\d{4}-\d{2}-\d{2}", s) else s


def as_map(v, key="a"):
    """The page stores checklist answers as [{en, a}] and other texts as [{en, text}]; accept a plain map too."""
    if isinstance(v, list):
        return {x.get("en", ""): x.get(key, "") for x in v if isinstance(x, dict)}
    return v or {}


def fill(path_in, req):
    req = dict(req, hazards=as_map(req.get("hazards")), precautions=as_map(req.get("precautions")), other=as_map(req.get("other"), "text"))
    z = zipfile.ZipFile(path_in)
    xml = z.read("word/document.xml").decode("utf-8")
    F = Filler(xml)
    # generic header labels; the confined-space form has its own "Date & Time of Entry" and "Location/Name of ..." cells below
    has_expiry = any(norm(c.text(F.xml)).startswith("EXPIRY DATE") for c in F.cells)
    F.label(r"^DATE\b(?!\s*(&|/))", dmy(req.get("weekStart")) if has_expiry else f"{dmy(req.get('weekStart'))} to {dmy(req.get('weekEnd'))}")
    # excavation form
    F.label(r"^EXPIRY DATE$", dmy(req.get("weekEnd")))
    F.label(r"^AREA / LOCATION", req.get("location", ""))
    F.label(r"^CONTRACTOR SUPERVISOR", " ".join(x for x in [req.get("receiver", ""), req.get("receiverContact", "")] if x))
    F.label(r"^JOB DESCRIPTION", req.get("description", ""))
    ex = req.get("exc") or {}
    F.blank(r"^DEPTH\b", ex.get("depth")); F.blank(r"^WIDTH\b", ex.get("width")); F.blank(r"^LENGTH\b", ex.get("length"))
    soil = {"stable": r"^STABLE ?ROCK$", "A": r"^TYPE A$", "B": r"^TYPE B$", "C": r"^TYPE C$"}
    if ex.get("soil") in soil:
        F.tick_choice(soil[ex["soil"]])
    for a in ex.get("access") or []:
        F.tick_choice({"stairway": r"^STAIRWAY$", "ladders": r"^LADDERS$", "ramp": r"^RAMP$"}.get(a, r"^$"))
    osha = {"B": (r"^B –", 0), "C": (r"^C –", 0), "D": (r"^C –", 1), "E": (r"^E –", 0), "F": (r"^E –", 1), "RPE": (r"^RPE –", 0)}
    for o in ex.get("osha") or []:
        if o in osha:
            F.tick_choice(*osha[o])
    F.label_or_next(r"^PERMIT N(O|UMBER)\b", req.get("permitNo", ""))
    F.label(r"^FROM\b", req.get("from", ""))
    F.label(r"^TO\b", req.get("to", ""))
    F.label(r"^LOCATION\b(?!/NAME)", req.get("location", ""))
    F.label(r"^CONTRACTOR", req.get("contractor", ""))
    F.label(r"^WORK SITE SUPERVISOR CONTACT", " ".join(x for x in [req.get("receiver", ""), req.get("receiverContact", "")] if x))
    F.label(r"^DESCRIPTION OF WORK", req.get("description", ""))
    F.after(r"^NUMBER$", req.get("raNo", ""))
    F.after(r"^DESCRIPTION$", req.get("raTitle", ""))
    ans = {}
    for lst in ("hazards", "precautions"):
        for k, v in (req.get(lst) or {}).items():
            ans[norm(k)] = v
    marked = F.checklist(ans)
    for k, other in (req.get("other") or {}).items():           # free text for OTHER: rows
        F.label(r"^" + re.escape(norm(k).rstrip(":")) + r"\s*:?$", other)
    # confined-space form: parties, entry details, workers, air tests, coordinator and receipt
    F.after(r"^CLIENT:$", CFG.get("client", ""))
    F.after(r"^CONSULTANT:$", CFG.get("consultant", ""))
    F.after(r"^CONTRACTOR:$", CFG.get("contractor", ""))
    F.after(r"^ASSOCIATED RISK ASSESSMENT", req.get("raNo", ""))
    F.label(r"^LOCATION/NAME OF CONFINED SPACE", req.get("location", ""))
    F.label(r"^DATE & TIME OF ENTRY", f"{dmy(req.get('weekStart'))} {req.get('from', '')}".strip())
    F.label(r"^PURPOSE OF ENTRY", req.get("description", ""))
    F.label(r"^EXPIRY DATE & TIME", f"{dmy(req.get('weekEnd'))} {req.get('to', '')}".strip())
    F.label(r"^NAME OF PERSON IN CHARGE", req.get("receiver", ""))
    F.label(r"^CONTACT PHONE NO", req.get("receiverContact", ""))
    F.label(r"^NAME OF ATTENDANT", ", ".join(req.get("attendants") or []))
    F.label(r"^HAZARDS ASSOCIATED WITH ENTRY", " — ".join(x for x in [req.get("raNo", ""), req.get("raTitle", "")] if x))
    for i, e in enumerate(F.rows_under(["NAME OF TRAINED WORKER", "FILE NO"])[:len(req.get("workers") or [])]):
        F.put(e[0], req["workers"][i].get("name", "")); F.put(e[1], req["workers"][i].get("fileNo", ""))
    gas_rows = {"OXYGEN": "o2", "COMBUSTIBLE GAS": "lel", "CARBON MONOXIDE": "co", "HYDROGEN SULFIDE": "h2s"}
    for t in F.tabs:
        for r in t:
            k = norm(r[0].text(F.xml)) if r else ""
            g = req.get("gas") or {}
            if k in gas_rows and len(r) > 2 and g.get(gas_rows[k]):
                F.put(r[2], g[gas_rows[k]])
    name_rows = [r for t in F.tabs for r in t if r and norm(r[0].text(F.xml)) == "NAME" and len(r) >= 6]
    for r, who in zip(name_rows, [req.get("hse", ""), req.get("receiver", "")]):
        F.put(r[1], who)
        j = next((j for j, c in enumerate(r) if norm(c.text(F.xml)) == "DATE/TIME"), None)
        if j is not None and j + 1 < len(r):
            F.put(r[j + 1], f"{dmy(req.get('weekStart'))} {req.get('from', '')}".strip())
    for i, e in enumerate(F.rows_under(["SN", "NAME", "SL NUMBER", "QNTY"])[:len(req.get("equipment") or [])]):
        it = req["equipment"][i]
        F.put(e[1], it.get("name", "")); F.put(e[2], it.get("serial", "")); F.put(e[3], it.get("qty", ""))
    for i, e in enumerate(F.rows_under(["GEAR", "NUMBERS", "CONDITION", "COMMENTS"])[:len(req.get("gear") or [])]):
        it = req["gear"][i]
        F.put(e[0], it.get("gear", "")); F.put(e[1], it.get("numbers", "")); F.put(e[2], it.get("condition", "")); F.put(e[3], it.get("comments", ""))
    fw = F.rows_under(["NAME", "SIGN IN TIME", "SIGN OUT TIME"])
    for i, name in enumerate((req.get("fireWatchers") or [])[:len(fw)]):
        F.put(fw[i][0], name)
    gas = req.get("gas") or {}
    for lab, key in (("^LEL", "lel"), ("^O2", "o2"), ("^CO PPM", "co")):
        F.after(lab, gas.get(key, ""))
    # declaration: names and dates under the four roles
    roles = F.rows_under(["WORKSITE SUPERVISOR", "TRANSFER", "", "PERMIT ISSUER"])
    names = [req.get("receiver", ""), "", req.get("hse", ""), req.get("issuer", "")]
    for r in roles[:3]:                         # Name, Sign, Date rows; the closing block below also has "Name"
        k = norm(r[0].text(F.xml))
        for j, c in enumerate(r):
            if k.startswith("NAME") and names[j]:
                F.put(c, names[j])
            elif k.startswith("DATE") and names[j]:
                F.put(c, dmy(req.get("weekStart")))
    new = F.result()
    files = {n: z.read(n) for n in z.namelist()}
    files["word/document.xml"] = new.encode("utf-8")
    return files, marked


def cmd_make(req_json, out):
    req = json.load(open(req_json, encoding="utf-8-sig"))
    ty = next(t for t in PJ["types"] if t["key"] == req["type"])
    if not os.path.exists(ty["form"]):
        raise SystemExit(f"The {ty['en']} form is not available as .docx yet: {ty['form']}")
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    files, marked = fill(ty["form"], req)
    wanted = {norm(k) for lst in ("hazards", "precautions") for k, v in as_map(req.get(lst)).items() if v in ("yes", "no", "na")}
    missing = sorted(wanted - marked)
    import xml.dom.minidom
    xml.dom.minidom.parseString(files["word/document.xml"])            # well-formed or raise
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, b in files.items():
            zo.writestr(n, b)
    res = {"written": out, "checklist_marked": len(marked), "not_found_on_form": missing}
    if ty.get("page3") and req.get("cranes"):
        p3 = os.path.splitext(out)[0] + " - crane details.docx"
        z = zipfile.ZipFile(ty["page3"])
        x = z.read("word/document.xml").decode("utf-8")
        F = Filler(x)
        for i, c in enumerate(req["cranes"][:9]):
            regs = [r for t in F.tabs for r in t if r and norm(r[0].text(x)) == "REGISTRATION NUMBER"]
            tuvs = [r for t in F.tabs for r in t if r and norm(r[0].text(x)) == "TUV AVAILABLE"]
            ops = [r for t in F.tabs for r in t if r and norm(r[0].text(x)) == "OPERATOR LICENSE"]
            oth = [r for t in F.tabs for r in t if r and norm(r[0].text(x)) == "OTHER DETAILS"]
            if i < len(regs):
                F.put(regs[i][1], f"{c.get('code', '')} {c.get('type', '')} {c.get('plate', '')}".strip())
                F.put(tuvs[i][1], c.get("tuv", ""))
                F.put(ops[i][1], c.get("operator", ""))
                F.put(oth[i][1], c.get("other", ""))
        fl = {n: z.read(n) for n in z.namelist()}
        fl["word/document.xml"] = F.result().encode("utf-8")
        if os.path.exists(p3):
            raise SystemExit(f"Refusing to overwrite {p3}")
        with zipfile.ZipFile(p3, "w", zipfile.ZIP_DEFLATED) as zo:
            for n, b in fl.items():
                zo.writestr(n, b)
        res["crane_details"] = p3
    print(json.dumps(res, ensure_ascii=False))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    {"settings": cmd_settings, "plan": cmd_plan, "make": cmd_make}[sys.argv[1]](*sys.argv[2:])

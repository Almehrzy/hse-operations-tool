# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Scan the consultant SVN folders and extract one record per SVN (read-only).

    python svn_scan.py <out.json>

Sources: the SVN folder in project.json, plus its extra_folders for numbers missing there.
For each SVN number it reads the consultant's notice (PDF) for date, location and subject, and
the contractor's close-out reply (Word header table, else the reply file's date).
"""
import datetime, io, json, os, re, sys, zipfile
import xml.etree.ElementTree as ET

from config import CFG

LP = "\\\\?\\"   # long-path prefix: some file paths pass 260 characters
SVN_DIR = CFG["svn"]["folder"]
EXTRA_DIRS = CFG["svn"].get("extra_folders", [])   # used only for numbers missing in SVN_DIR
SKIP = [w for w in CFG["svn"].get("skip_words", ["NCR", "SOR", "LTR"])]
SKIP_RE = re.compile(r"\b(" + "|".join(re.escape(w) for w in SKIP) + r")\b", re.I) if SKIP else None
NUMBER_RE = re.compile(CFG["svn"].get("number_pattern", r"SVN[\s_\-]*0*(\d{1,4})(?!\d)"), re.I)
REPLY_WORDS = re.compile(r"reply|response|acknowledge|close ?out|closure|corrective|\bcap\b|thank you|letter", re.I)
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def svn_no(name):
    """SVN number from a file or folder name (project number pattern, else the number right after 'SVN')."""
    if (SKIP_RE and SKIP_RE.search(name)) or not re.search(r"SVN", name, re.I):
        return None
    m = NUMBER_RE.search(name) or re.search(r"SVN[\s_\-]*0*(\d{1,3})(?!\d)", name, re.I)
    return int(m.group(1)) if m else None


def pdf_text(path, pages=2):
    from PyPDF2 import PdfReader
    try:
        r = PdfReader(path)
        return "\n".join((r.pages[i].extract_text() or "") for i in range(min(pages, len(r.pages))))
    except Exception as e:
        return f"__ERROR__ {e}"


def clean(s):
    return re.sub(r"\s+", " ", s or "").strip(" :.-–")


def parse_notice(text):
    out = {}
    m = (re.search(r"\bDate\b\s*:?\s*([0-3]?\d)\s*/\s*([01]?\d)\s*/\s*(20\d\d)", text)
         or re.search(r"\bDate\b\s*:?\s*([0-3]?\d)\s*/\s*([A-Za-z]{3,9})\s*/\s*(20\d\d)", text))
    if m:
        mon = m.group(2)
        mon = int(mon) if mon.isdigit() else MONTHS.get(mon[:3].lower())
        if mon:
            out["date"] = f"{m.group(3)}-{mon:02d}-{int(m.group(1)):02d}"
    m = re.search(r"\bLocation\b\s*:?\s*([^\n]*)", text)
    loc = clean(m.group(1)) if m else ""
    if not loc or re.match(r"(Subject|Violation|Recommend|Item|Observ|•)", loc, re.I):
        m2 = re.search(r"\n([^\n]{3,40})\n\s*Location\b", text)    # some notices print the value above the label
        loc = clean(m2.group(1)) if m2 and not re.search(r"SVN|Description", m2.group(1)) else ""
    out["location"] = loc[:80]
    m = re.search(r"Subject\s*:?\s*(.+?)(?:\s+A\*|\n)", text)
    if m:
        out["subject"] = clean(m.group(1))[:160]
    return out


def docx_header(path):
    """Label/value pairs from the first table of a Word file (the close-out header)."""
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("word/document.xml"))
    except Exception:
        return {}
    txt = lambda el: re.sub(r"\s+", " ", "".join(t.text or "" for t in el.iter(W + "t"))).strip()
    info = {}
    tbl = root.find(f"{W}body/{W}tbl")
    if tbl is not None:
        for tr in tbl.findall(W + "tr"):
            cells = [txt(tc) for tc in tr.findall(W + "tc")]
            for i in range(0, len(cells) - 1, 2):
                if cells[i]:
                    info[re.sub(r"\s+", " ", cells[i]).rstrip(":").strip()] = cells[i + 1]
    return info


def to_iso(s):
    s = re.sub(r"\s+", " ", re.sub(r"(\d)\s+(?=\d)", r"\1", s or "")).strip()
    for f in ("%d %B %Y", "%d %b %Y", "%B %d, %Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, f).date().isoformat()
        except ValueError:
            pass
    return ""


def subject_from_name(name):
    s = re.sub(r"\.(pdf|docx?)$", "", name, flags=re.I)
    s = re.sub(r"^.*?(?:WP[.\-]?03[.\-]0*\d+|SVN[\s_\-]*\d+(?:[\s_\-]*WP[.\-]?03)?)", "", s, flags=re.I)
    s = re.sub(r"[-–\s]*(reply|close ?out reply|closeout|rev\.? ?\d+)\b.*$", "", s.replace("_", " "), flags=re.I)
    s = clean(re.sub(r"\(\d\)\s*$", "", s))
    return "" if re.match(r"(contractor.?s response|acknowledge|corrective action|cap for)", s, re.I) else s


def collect(root, wanted=None):
    groups = {}
    for d, _, fs in os.walk(LP + root):
        if SKIP_RE and SKIP_RE.search(d[len(LP) + len(root):]):
            continue
        dir_no = svn_no(os.path.basename(d)) if d != LP + root else None
        for f in fs:
            if f.startswith("~") or not f.lower().endswith((".pdf", ".docx", ".doc")):
                continue
            n = svn_no(f) or dir_no
            if n is None or (wanted is not None and n not in wanted):
                continue
            groups.setdefault(n, []).append(os.path.join(d, f))
    return groups


def build(n, files, root):
    rel = lambda p: os.path.relpath(p[len(LP):], root)
    rec = {"no": n, "notes": [], "source_root": root}
    dirs = sorted({rel(os.path.dirname(f)) for f in files})
    rec["folder"] = " | ".join(dirs)
    notices = [f for f in files if f.lower().endswith(".pdf") and not REPLY_WORDS.search(os.path.basename(f))]
    replies = [f for f in files if f not in notices and (REPLY_WORDS.search(os.path.basename(f)) or f.lower().endswith((".docx", ".doc")))]
    notice = {}
    for f in notices:
        t = pdf_text(f)
        if "Violation" in t and "SVN" in t:
            notice = parse_notice(t)
            notice["file"] = os.path.basename(f)
            break
    if not notices:
        rec["notes"].append("consultant notice PDF not on file")
    elif not notice:
        rec["notes"].append("consultant notice PDF could not be read")
    rec.update({k: notice.get(k, "") for k in ("date", "location", "subject")})
    rec.update({"notice_file": notice.get("file", ""), "action_code": "", "reply_date": "", "reply_rev": "",
                "reply_status": "", "reply_file": "", "reply_date_source": ""})
    for f in sorted(replies, key=os.path.getmtime, reverse=True):
        if not f.lower().endswith(".docx"):
            continue
        info = docx_header(f)
        if not (info.get("Action Code") or info.get("Reply Date") or info.get("Status")):
            continue
        hn = svn_no(re.sub(r"\s+", "", "SVN " + info.get("SVN No.", "")))
        if hn is not None and hn != n:
            rec["notes"].append(f"Word reply '{os.path.basename(f)}' holds the close-out for SVN {hn:03d}, not {n:03d}")
            replies = [r for r in replies if r != f]
            continue
        rec["reply_file"] = os.path.basename(f)
        rec["action_code"] = re.sub(r"\s*[–-]\s*.*$", "", info.get("Action Code", ""))
        rec["reply_rev"] = info.get("Revision", "")
        rec["reply_status"] = info.get("Status", "")
        rec["received_doc"] = to_iso(info.get("Date Received", ""))
        rd = to_iso(info.get("Reply Date", "")) or to_iso(info.get("Date Closed", ""))
        if rd:
            rec["reply_date"], rec["reply_date_source"] = rd, "reply"
        for src, dst in (("Location", "location"), ("Subject", "subject"), ("Date Issued", "date")):
            if info.get(src) and (dst != "location" or not rec[dst] or len(rec[dst]) < 3):
                rec[dst] = to_iso(info[src]) if dst == "date" else clean(info[src])
        break
    if replies and not rec["reply_date"]:
        newest = max(replies, key=os.path.getmtime)
        rec["reply_file"] = rec["reply_file"] or os.path.basename(newest)
        rec["reply_date"] = datetime.date.fromtimestamp(os.path.getmtime(newest)).isoformat()
        rec["reply_date_source"] = "file date"
    if not replies:
        rec["notes"].append("reply not on file")
    if not rec["subject"]:
        cands = sorted({os.path.dirname(f) for f in files} - {LP + root}) + sorted(files, key=lambda p: "SVN" not in os.path.basename(p).upper())
        for f in cands:
            s = subject_from_name(os.path.basename(f))
            if len(s) > 6:
                rec["subject"] = s
                rec["notes"].append("subject taken from the file name")
                break
    if not rec["date"]:
        for f in [os.path.dirname(files[0])] + files:
            m = re.match(r"(20\d\d)\.(\d\d)\.(\d\d)", os.path.basename(f))
            if m:
                rec["date"] = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                rec["notes"].append("issue date taken from the folder name")
                break
    if not rec["date"]:
        rec["notes"].append("issue date not readable")
    rec["received"] = rec.get("received_doc", "")
    rec["received_source"] = "reply" if rec["received"] else ""
    if notices and not rec["received"]:
        rec["received_source"] = "file date"
        rec["received"] = datetime.date.fromtimestamp(os.path.getmtime(min(notices, key=os.path.getmtime))).isoformat()
    return rec


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    groups = collect(SVN_DIR)
    records = [build(n, fs, SVN_DIR) for n, fs in sorted(groups.items())]
    top = max(groups)
    missing = [n for n in range(1, top + 1) if n not in groups]
    for extra in EXTRA_DIRS:
        found = collect(extra, set(missing))
        for n, fs in found.items():
            r = build(n, fs, extra)
            r["notes"].insert(0, "files are outside the records folder: " + extra)
            records.append(r)
        missing = [n for n in missing if n not in found]
    records.sort(key=lambda r: r["no"])
    with open(sys.argv[1], "w", encoding="utf-8") as f:
        json.dump({"records": records, "missing": missing}, f, ensure_ascii=False, indent=1)
    for r in records:
        print(f"{r['no']:3d} | {r['date'] or '-':10} | {r['action_code'] or '-':2} | {r['reply_date'] or '-':10} {('(' + r['reply_date_source'] + ')') if r['reply_date_source'] else '':11} | {r['location'][:22]:22} | {r['subject'][:52]:52} | {'; '.join(r['notes'])}")
    print("missing numbers:", missing)


if __name__ == "__main__":
    main()

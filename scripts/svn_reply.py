# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Contractor Closeout Report for a consultant SVN, in the house format (template set in project.json).

    python svn_reply.py photos <notice.pdf> <out_dir>        -> save the photos embedded in the consultant notice
    python svn_reply.py make   <reply.json> <out.docx>       -> write the Closeout Report

reply.json:
{
 "svn": "<SVN number as the consultant wrote it>", "revision": "Rev.00", "action_code": "C – Needs Immediate Attention",
 "contract": "(default from project.json)", "package": "(default from project.json)", "location": "...", "subject": "...", "status": "Submitted for closure",
 "date_issued": "2026-10-08", "date_received": "2026-10-08", "reply_date": "2026-10-13",
 "extra_header": [["Rev.00 Date", "08 October 2026"], ["SSC Review", "..."]],          (optional, revisions)
 "purpose": "...",
 "summary": ["paragraph", ...], "root_cause": "...",                                   (root_cause optional)
 "clarification": ["paragraph", ...],                                                  (optional, "<contractor> Clarification")
 "pace_comments": "quoted consultant comment text",                                     (optional, revisions)
 "corrective": [{"title": "Violation 1 – ...", "bullets": ["..."], "status": "Closed."}],  (title/status optional)
 "preventive_title": "Preventive Measures", "preventive": ["..."],
 "photos": {"before": [{"file": "...", "caption": "..."}], "after": [...], "tbt": [...],
            "before_title": "Before: SVN photo (08 October 2026)", "after_title": "After: corrective actions", "tbt_title": "Toolbox Talk"},
 "closeout": "...",
 "signatures": [["Prepared by (HSE Engineer)", "<name>", "2026-10-13"], ["Reviewed by (HSE Manager)", "<name>", "2026-10-13"]]   (optional; defaults from project.json)
}
Photos are numbered in order: before, after, TBT. Refer to them in the text as (Photo N) in that order.
"""
import datetime, io, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
from config import CFG, path as cfg_path

TEMPLATE = cfg_path(CFG["svn"]["template"])
CAL = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
CAL4 = '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
GREY = '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
EMU_CM = 360000
PLACEHOLDER = re.compile(r"\binsert\b|\bTBD\b|\bxxx+\b|\[.*?\]|\bif available\b|lorem", re.I)


def long_date(s):
    return datetime.date.fromisoformat(s).strftime("%d %B %Y") if s else ""


def t(text):
    return f'<w:t xml:space="preserve">{escape(str(text))}</w:t>'


def p(text, size=22, bold=False, after=120, indent=None, center=False, highlight=False):
    ppr = f'<w:spacing w:after="{after}"/>' + (f'<w:ind w:left="{indent}"/>' if indent else "") + ('<w:jc w:val="center"/>' if center else "")
    rpr = CAL + ("<w:b/>" if bold else '<w:b w:val="0"/>') + f'<w:sz w:val="{size}"/>'
    return f"<w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr>{t(text)}</w:r></w:p>"


def spacer():
    return '<w:p><w:pPr><w:spacing w:after="40"/></w:pPr></w:p>'


def h1(text):
    return p(text, size=24, bold=True, after=80)


def h2(text):
    return p(text, size=22, bold=True, after=120)


def bullet(text):
    return p("•  " + text, after=60, indent=360)


def status(text):
    return (f'<w:p><w:pPr><w:spacing w:after="120"/><w:ind w:left="360"/></w:pPr>'
            f'<w:r><w:rPr>{CAL}<w:b/><w:sz w:val="22"/></w:rPr>{t("Status: ")}</w:r>'
            f'<w:r><w:rPr>{CAL}<w:b w:val="0"/><w:sz w:val="22"/></w:rPr>{t(text)}</w:r></w:p>')


def cell(text, width, label=False, mar=(50, 80)):
    shade = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if label else ""
    tcpr = (f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:tcBorders>{GREY}</w:tcBorders>{shade}'
            f'<w:tcMar><w:top w:w="{mar[0]}" w:type="dxa"/><w:left w:w="{mar[1]}" w:type="dxa"/><w:bottom w:w="{mar[0]}" w:type="dxa"/><w:right w:w="{mar[1]}" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>')
    rpr = CAL4 + ("<w:b/><w:bCs/>" if label else "")
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:spacing w:after="20"/></w:pPr><w:r><w:rPr>{rpr}</w:rPr>{t(text)}</w:r></w:p></w:tc>'


def table(rows, widths, total, mar=(50, 80)):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar>'
           f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>']
    for r in rows:
        out.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + "".join(cell(v, widths[i], label=(i % 2 == 0), mar=mar) for i, v in enumerate(r)) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def quote_box(text):
    tcpr = f'<w:tcPr><w:tcW w:w="9715" w:type="dxa"/><w:tcBorders>{GREY}</w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>'
    body = (f'<w:p><w:pPr><w:spacing w:after="60"/></w:pPr><w:r><w:rPr>{CAL}<w:b/><w:sz w:val="22"/></w:rPr>{t(CFG.get("consultant_short", "Consultant") + " comments on the previous submission")}</w:r></w:p>'
            f'<w:p><w:pPr><w:spacing w:after="20"/></w:pPr><w:r><w:rPr>{CAL}<w:i/><w:sz w:val="22"/></w:rPr>{t(chr(8220) + text + chr(8221))}</w:r></w:p>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="9715" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9715"/></w:tblGrid>'
            f'<w:tr><w:tc>{tcpr}{body}</w:tc></w:tr></w:tbl>')


# ---------- photos ----------
def prep_image(path, max_w_cm, max_h_cm):
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    im.thumbnail((1600, 1600))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    w, h = im.size
    s = min(max_w_cm / w, max_h_cm / h)
    return buf.getvalue(), int(w * s * EMU_CM), int(h * s * EMU_CM)


def drawing(rid, n, cx, cy):
    a, pic = "http://schemas.openxmlformats.org/drawingml/2006/main", "http://schemas.openxmlformats.org/drawingml/2006/picture"
    return (f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{8000 + n}" name="Photo {n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="{a}"><a:graphicData uri="{pic}"><pic:pic xmlns:pic="{pic}"><pic:nvPicPr><pic:cNvPr id="{8000 + n}" name="photo{n}.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>')


class Media:
    def __init__(self):
        self.items, self.n = [], 0

    def add(self, path, max_w, max_h):
        self.n += 1
        data, cx, cy = prep_image(path, max_w, max_h)
        rid = f"rIdPh{self.n}"
        self.items.append((rid, f"media/photo{self.n}.jpeg", data))
        return f'<w:r>{drawing(rid, self.n, cx, cy)}</w:r>', self.n


def photo_block(media, entry, max_w=14.0, max_h=10.0):
    run, n = media.add(entry["file"], max_w, max_h)
    return (f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{run}</w:p>' +
            p(f"Photo {n}: {entry['caption']}", size=20, center=True))


def photo_table(media, entries):
    cw = 4873
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="9715" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="{cw}"/><w:gridCol w:w="{cw}"/></w:tblGrid>']
    for i in range(0, len(entries), 2):
        pair = entries[i:i + 2]
        imgs, caps = [], []
        for e in pair:
            run, n = media.add(e["file"], 8.0, 10.0)
            imgs.append(f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{cw}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr>{run}</w:p></w:tc>')
            caption = p(f"Photo {n}: {e['caption']}", size=20, center=True)
            caps.append(f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{cw}"/></w:tcPr>{caption}</w:tc>')
        while len(imgs) < 2:
            empty = f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{cw}"/></w:tcPr><w:p/></w:tc>'
            imgs.append(empty)
            caps.append(empty)
        out.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + "".join(imgs) + "</w:tr><w:tr>" + "".join(caps) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def grid_table(headers, rows, widths):
    """Plain grid: shaded bold header row, then rows of text."""
    total = sum(widths)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)

    def c(text, w, head):
        shade = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if head else ""
        rpr = CAL4 + ("<w:b/><w:bCs/>" if head else "") + '<w:sz w:val="20"/>'
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:tcBorders>{GREY}</w:tcBorders>{shade}'
                f'<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tcMar></w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr>{rpr}</w:rPr>{t(text)}</w:r></w:p></w:tc>')
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>',
           '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + "".join(c(hd, widths[i], True) for i, hd in enumerate(headers)) + "</w:tr>"]
    for r in rows:
        out.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + "".join(c(str(v), widths[i], False) for i, v in enumerate(r)) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def capa_block(capa, sec):
    """Root cause (5-Why) and corrective / preventive action tables, as in the NCR-006 reply."""
    b = [h1(f"{sec}. Root Cause Analysis and CAPA")]
    if capa.get("why"):
        b.append(h2("5-Why Analysis"))
        b.append(grid_table(["#", "Question", "Finding"], [[i + 1, q, a] for i, (q, a) in enumerate(capa["why"])], [500, 4300, 4915]))
        b.append(spacer())
    if capa.get("root_cause"):
        b.append(p("Root cause: " + capa["root_cause"], bold=False))
    for key, title, prefix in (("corrective", "Corrective Actions", "C"), ("preventive", "Preventive Actions", "P")):
        rows = capa.get(key) or []
        if rows:
            b.append(h2(title))
            b.append(grid_table(["Ref.", "Action", "Responsible", "Target / Completed", "Status"],
                                [[f"{prefix}{i + 1}"] + list(r) for i, r in enumerate(rows)], [700, 4615, 1800, 1500, 1100]))
            b.append(spacer())
    if capa.get("verification"):
        b.append(h2("Verification of Effectiveness"))
        b.append(p(capa["verification"]))
    return b


def signature_table(sigs):
    rows = []
    for role, name, date in sigs:
        rows.append([role, name, "Signature", "Date: " + long_date(date)])
    widths = [2400, 2400, 2400, 2160]
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' +
           "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"]
    for r in rows:
        out.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + "".join(cell(v, widths[i], label=(i in (0, 2)), mar=(140, 90)) for i, v in enumerate(r)) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


# ---------- checks ----------
def check(req):
    problems = []
    svn = req["svn"]
    m = re.search(r"(\d{3})$", svn)
    num = m.group(1) if m else None
    texts = [req.get("purpose", ""), req.get("closeout", ""), req.get("root_cause", ""), req.get("pace_comments", "")]
    texts += req.get("summary", []) + req.get("clarification", []) + req.get("preventive", [])
    for c in req.get("corrective", []):
        texts += [c.get("title", "")] + c.get("bullets", [])
    for k in ("before", "after", "tbt"):
        texts += [e.get("caption", "") for e in req.get("photos", {}).get(k, [])]
    dtype = req.get("doc_type", "SVN").upper()
    # In an SVN reply any other SVN number is a mistake; an NCR reply may cite the SVNs it escalates, so only other NCR numbers are checked.
    pattern = CFG["svn"].get("number_pattern", r"SVN[\s_\-]*0*(\d{1,4})(?!\d)") if dtype == "SVN" else r"NCR[-\s\w.]*?HSE[-\s]*0*(\d{1,3})(?!\d)"
    for tx in texts:
        for ref in re.findall(pattern, tx, re.I):
            if dtype == "SVN" and re.search(r"\bNCR\b", tx) and not re.search(r"SVN", tx):
                continue
            if num and int(ref) != int(num):
                problems.append(f"text mentions {dtype} {ref} but this reply is for {svn}: {tx[:80]}")
        if PLACEHOLDER.search(tx) and "pace_comments" not in tx:
            problems.append(f"placeholder-like text: {tx[:80]}")
    n_photos = sum(len(req.get("photos", {}).get(k, [])) for k in ("before", "after", "tbt"))
    for tx in texts:
        for ref in re.findall(r"Photos?\s+(\d+)", tx):
            if int(ref) > n_photos:
                problems.append(f"text refers to Photo {ref} but there are only {n_photos} photos")
    for k in ("before", "after", "tbt"):
        for e in req.get("photos", {}).get(k, []):
            if not os.path.exists(e["file"]):
                problems.append(f"photo file not found: {e['file']}")
    return problems


def replace_across_runs(xml, old, new):
    """Replace text that Word may have split over several <w:t> runs."""
    ts = list(re.finditer(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)", xml))
    joined = "".join(m.group(2) for m in ts)
    i = joined.find(old)
    if i < 0:
        return xml, False
    pos, starts = 0, []
    for m in ts:
        starts.append(pos)
        pos += len(m.group(2))
    out, last = [], 0
    end = i + len(old)
    placed = False
    for m, s in zip(ts, starts):
        e = s + len(m.group(2))
        txt = m.group(2)
        if e <= i or s >= end:
            continue
        a, b = max(i - s, 0), min(end - s, len(txt))
        txt = txt[:a] + ("" if placed else escape(new)) + txt[b:]
        placed = True
        out.append(xml[last:m.start(2)] + txt)
        last = m.end(2)
    out.append(xml[last:])
    return "".join(out), True


# ---------- build ----------
def make(req, out_path):
    problems = check(req)
    if problems and not req.get("allow_warnings"):
        raise SystemExit("Not written. Fix these first:\n- " + "\n- ".join(problems))
    rev = req.get("revision", "Rev.00")
    dtype = req.get("doc_type", "SVN").upper()
    name = {"SVN": "Violation Notice", "NCR": "Non-Compliance Report"}.get(dtype, "Violation Notice")
    media = Media()
    b = [p(f"Safety, Health and Environment {name} ({dtype}) Closeout Report ({rev})", size=28, bold=True, after=120, center=True)]
    hdr = [[f"{dtype} No.", req["svn"], "Action Code" if dtype == "SVN" else "Remedial Due", req.get("action_code", "C – Needs Immediate Attention") if dtype == "SVN" else long_date(req.get("remedial_due"))],
           ["Project", CFG["project_name"], "Contract No.", req.get("contract", CFG["contract"])],
           ["Package", req.get("package", CFG["package"]), "Location", req.get("location", "")],
           ["Issued by", CFG["consultant"], "PMC / Engineer", CFG["pmc"]],
           ["Date Issued", long_date(req.get("date_issued")), "Date Received", long_date(req.get("date_received"))],
           ["Subject", req["subject"], "Status", req.get("status", "Submitted for closure")],
           ["Revision", rev, "Reply Date", long_date(req.get("reply_date"))]]
    extra = req.get("extra_header", [])
    for i in range(0, len(extra), 2):
        pair = extra[i] + (extra[i + 1] if i + 1 < len(extra) else ["", ""])
        hdr.append(pair)
    b.append(table(hdr, [1900, 2780, 1900, 3135], 9715))
    b.append(spacer())
    sec = 1
    b.append(h1(f"{sec}. Purpose")); sec += 1
    b.append(p(req["purpose"]))
    if req.get("pace_comments"):
        b.append(quote_box(req["pace_comments"]))
        b.append(spacer())
    many = len(req.get("corrective", [])) > 1
    b.append(h1(f"{sec}. " + ("Summary of Non-Compliance" if dtype == "NCR" else "Summary of Violation" + ("s" if many else "")))); sec += 1
    for para in req.get("summary", []):
        b.append(p(para))
    if req.get("root_cause"):
        b.append(h2("Root Cause"))
        b.append(p(req["root_cause"]))
    if req.get("clarification"):
        b.append(h2(CFG.get("contractor_short", "Contractor") + " Clarification"))
        for para in req["clarification"]:
            b.append(p(para))
    b.append(h1(f"{sec}. Corrective Actions Taken")); sec += 1
    for block in req.get("corrective", []):
        if block.get("title"):
            b.append(h2(block["title"]))
        for x in block.get("bullets", []):
            b.append(bullet(x))
        if block.get("status"):
            b.append(status(block["status"]))
    b.append(h1(f"{sec}. {req.get('preventive_title', 'Preventive Measures')}")); sec += 1
    for x in req.get("preventive", []):
        b.append(bullet(x))
    if req.get("capa"):
        b.extend(capa_block(req["capa"], sec)); sec += 1
    ph = req.get("photos", {})
    if any(ph.get(k) for k in ("before", "after", "tbt")):
        b.append(h1(f"{sec}. Photographic Evidence"))
        sub = 1
        if ph.get("before"):
            b.append(h2(f"{sec}.{sub} {ph.get('before_title', 'Before: SVN photo')}")); sub += 1
            for e in ph["before"]:
                b.append(photo_block(media, e, max_w=9.0, max_h=10.0))
        if ph.get("after"):
            b.append(h2(f"{sec}.{sub} {ph.get('after_title', 'After: corrective actions')}")); sub += 1
            b.append(photo_table(media, ph["after"]))
            b.append(spacer())
        if ph.get("tbt"):
            b.append(h2(f"{sec}.{sub} {ph.get('tbt_title', 'Toolbox Talk')}")); sub += 1
            for e in ph["tbt"]:
                b.append(photo_block(media, e))
        sec += 1
    b.append(h1(f"{sec}. Close-Out Statement"))
    b.append(p(req["closeout"]))
    b.append(spacer())
    b.append(signature_table(req.get("signatures") or [[role, name, req.get("reply_date")] for role, name in CFG["signatures"]]))

    with zipfile.ZipFile(TEMPLATE) as z:
        files = {n: z.read(n) for n in z.namelist()}
        order = z.namelist()
    doc = files["word/document.xml"].decode("utf-8")
    body_start = doc.index("<w:body>") + len("<w:body>")
    sect = doc.rindex("<w:sectPr")
    doc = doc[:body_start] + "".join(b) + doc[sect:]
    # drop the template's own body photos, keep header logos
    rels = files["word/_rels/document.xml.rels"].decode("utf-8")
    old_body_media = re.findall(r'<Relationship Id="(rId2[3-6])"[^>]*Target="(media/[^"]+)"/>', rels)
    for rid, target in old_body_media:
        rels = re.sub(rf'<Relationship Id="{rid}"[^>]*/>', "", rels)
        files.pop("word/" + target, None)
        order = [n for n in order if n != "word/" + target]
    for rid, target, data in media.items:
        rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/></Relationships>')
        files["word/" + target] = data
        order.append("word/" + target)
    files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    files["word/document.xml"] = doc.encode("utf-8")
    ftr = files["word/footer1.xml"].decode("utf-8")
    ftr, ok = replace_across_runs(ftr, CFG["svn"]["template_number"], req["svn"])
    if not ok:
        raise SystemExit("Template footer changed: SVN number not found")
    files["word/footer1.xml"] = ftr.encode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    if 'Extension="jpeg"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ', 1)
    files["[Content_Types].xml"] = ct.encode("utf-8")
    import xml.etree.ElementTree as ET
    for part in ("word/document.xml", "word/footer1.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"):
        ET.fromstring(files[part])
    if os.path.exists(out_path):
        raise SystemExit(f"Refusing to overwrite {out_path}")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(n, files[n])
    return {"written": out_path, "photos": media.n, "warnings": problems}


def photos(pdf, out_dir, min_px=400):
    from PyPDF2 import PdfReader
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    from PIL import Image

    def xobjects(res):
        xo = res.get("/XObject") if res else None
        if not xo:
            return
        xo = xo.get_object()
        for name in xo:
            obj = xo[name].get_object()
            if obj.get("/Subtype") == "/Image":
                yield name, obj
            elif obj.get("/Subtype") == "/Form":
                yield from xobjects(obj.get("/Resources"))

    def positions(reader, page):
        """Top-left position of each image drawn on the page (PDF y grows upwards), from the content stream."""
        from PyPDF2.generic import ContentStream
        pos, ctm, stack = {}, [1, 0, 0, 1, 0, 0], []
        try:
            ops = ContentStream(page.get_contents(), reader).operations
        except Exception:
            return pos
        for operands, op in ops:
            if op == b"q":
                stack.append(list(ctm))
            elif op == b"Q" and stack:
                ctm = stack.pop()
            elif op == b"cm":
                a, b, c, d, e, f = [float(x) for x in operands]
                A, B, C, D, E, F = ctm
                ctm = [a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F]
            elif op == b"Do":
                pos.setdefault(str(operands[0]), (ctm[4], ctm[5] + ctm[3]))
        return pos

    reader = PdfReader(pdf)
    for pi, page in enumerate(reader.pages, 1):
        where = positions(reader, page)
        found = list(xobjects(page.get("/Resources")))
        found.sort(key=lambda t: (-where.get(str(t[0]), (0, 0))[1], where.get(str(t[0]), (0, 0))[0]))   # top to bottom, then left to right
        for ii, (name, obj) in enumerate(found, 1):
            try:
                w, h = int(obj["/Width"]), int(obj["/Height"])
                if min(w, h) < min_px:
                    continue          # logos and small marks
                filt = obj.get("/Filter")
                filt = filt[0] if isinstance(filt, list) else filt
                if filt in ("/DCTDecode", "/JPXDecode"):
                    im = Image.open(io.BytesIO(obj._data))
                else:
                    data = obj.get_data()
                    mode = "RGB" if len(data) >= w * h * 3 else "L"
                    im = Image.frombytes(mode, (w, h), data[: w * h * (3 if mode == "RGB" else 1)])
                path = os.path.join(out_dir, f"pace_p{pi}_{ii:02d}.jpg")
                im.convert("RGB").save(path, "JPEG", quality=90)
                x, y = where.get(str(name), (None, None))
                saved.append({"file": path, "page": pi, "size": [w, h], "x": x and round(x), "y": y and round(y)})
            except Exception as e:
                saved.append({"page": pi, "image": str(name), "error": str(e)[:120]})
    return saved


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if len(sys.argv) < 4 or sys.argv[1] not in ("make", "photos"):
        raise SystemExit(__doc__)
    if sys.argv[1] == "photos":
        min_px = int(sys.argv[4]) if len(sys.argv) > 4 else 400     # SOR forms embed small photos: use 100
        print(json.dumps(photos(sys.argv[2], sys.argv[3], min_px), indent=1))
    else:
        req = json.load(open(sys.argv[2], encoding="utf-8-sig"))
        print(json.dumps(make(req, sys.argv[3]), ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

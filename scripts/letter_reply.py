# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Reply to a letter from the Engineer (PMC), in the house format: the letter body in Word, sent through the document
control system (which adds the outgoing number and header). Template: project.json letters.template (Arial 10.5,
footer page numbers).

    python letter_reply.py scaffold <letter.json> <reply.json>    start a reply from the letter in the tool
                                                                   (an exported `consultant` document of kind LTR)
    python letter_reply.py make <reply.json> <out.docx>           write the reply (never overwrites)

reply.json:
{
 "ref": "XYZ01-P1-PMC-LTR-AB-SF-00296", "date": "2026-10-12", "received": "2026-10-12",
 "subject": "Implementation of Disaster Management Plan",
 "references": ["XYZ01-P1-PMC-LTR-AB-MC-00200 – Implementation of Disaster Management Plan"],   (optional, more refs)
 "salutation": "Dear Sir,",                                                                          (optional)
 "opening": "...",                    optional; default: we acknowledge receipt of the Engineer's letter ref. … dated …
 "body": ["paragraph", {"bullets": ["...", "..."]}, {"numbered": ["...", "..."]}, ...],
 "table": {"intro": "...", "rows": [["what the letter asked", "status", "evidence / reference"]]},   (optional)
 "closing": "We trust this addresses the Engineer's concerns ...",
 "signatory": ["<name>", "Project Manager / Contractor's Representative"],                         (optional)
 "enclosures": ["...", "..."],
 "photos": [{"file": "...", "caption": "Photo 1: ..."}]           Appendix A after the letter; added to the enclosures
}
It refuses to write when the text keeps placeholder text, names an Engineer's letter (letters.ref_pattern) that is not this one or one of its
references, or points to a photo file that does not exist. No signature image is ever put in: the PM signs in Aconex.
"""
import datetime as dt, io, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
from config import CFG, path as cfg_path
import svn_reply as SR

LC = CFG.get("letters", {})
TEMPLATE = cfg_path(LC.get("template", r"templates\Letter reply template.docx"))
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>{b}<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
LTR_REF = re.compile(LC.get("ref_pattern", r"LTR-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{5}"))


def long(s):
    return dt.date.fromisoformat(str(s)[:10]).strftime("%d %B %Y") if s else ""


def run(text, bold=False):
    return f'<w:r>{RPR.format(b="<w:b/>" if bold else "")}<w:t xml:space="preserve">{escape(str(text))}</w:t></w:r>'


def para(text="", bold=False, justify=True, after=120, keep=False, runs=None):
    ppr = f'<w:spacing w:after="{after}" w:line="276" w:lineRule="auto"/>' + ("<w:keepNext/>" if keep else "") + ('<w:jc w:val="both"/>' if justify else "")
    return f'<w:p><w:pPr>{ppr}</w:pPr>{runs if runs is not None else (run(text, bold) if text else "")}</w:p>'


def bullet(text, num_id):
    return (f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
            f'<w:spacing w:after="80" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>{run(text)}</w:p>')


def numbered(i, text):
    return (f'<w:p><w:pPr><w:spacing w:after="80" w:line="276" w:lineRule="auto"/><w:ind w:left="426" w:hanging="426"/><w:jc w:val="both"/></w:pPr>'
            f'{run(f"{i}.")}<w:r>{RPR.format(b="")}<w:tab/></w:r>{run(text)}</w:p>')


def label_line(label, value):
    return (f'<w:p><w:pPr><w:spacing w:after="60" w:line="276" w:lineRule="auto"/><w:tabs><w:tab w:val="left" w:pos="1418"/></w:tabs>'
            f'<w:ind w:left="1560" w:hanging="1560"/></w:pPr>{run(label, True)}<w:r>{RPR.format(b="")}<w:tab/></w:r>{run((": " if label else "  ") + value)}</w:p>')


def table(rows, header=("Requirement in the Engineer's letter", "Status", "Detail / Evidence")):
    widths = [3600, 1500, 3926]
    border = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))

    def cell(t, w, head=False):
        shade = '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' if head else ""
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shade}</w:tcPr><w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>'
                f'<w:r>{RPR.format(b="<w:b/>" if head else "").replace("21", "19")}<w:t xml:space="preserve">{escape(str(t))}</w:t></w:r></w:p></w:tc>')
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders><w:tblLayout w:type="fixed"/>'
           f'<w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid>{"".join(f"<w:gridCol w:w={chr(34)}{w}{chr(34)}/>" for w in widths)}</w:tblGrid>']
    out.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>' + "".join(cell(h, w, True) for h, w in zip(header, widths)) + "</w:tr>")
    for r in rows:
        out.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + "".join(cell(r[i] if i < len(r) else "", w) for i, w in enumerate(widths)) + "</w:tr>")
    return "".join(out) + "</w:tbl>" + para(after=60)


def texts(req):
    yield req.get("subject", "")
    for k in ("opening", "closing"):
        yield req.get(k, "")
    for b in req.get("body", []):
        if isinstance(b, str):
            yield b
        else:
            for v in b.values():
                yield from (v if isinstance(v, list) else [v])
    for r in (req.get("table") or {}).get("rows", []):
        yield from r
    yield (req.get("table") or {}).get("intro", "")
    yield from req.get("enclosures", [])
    for p in req.get("photos", []):
        yield p.get("caption", "")


def check(req):
    problems = []
    allowed = {LTR_REF.search(req["ref"]).group(0)} if LTR_REF.search(req.get("ref", "")) else set()
    for r in req.get("references", []):
        allowed |= set(LTR_REF.findall(r))
    for t in texts(req):
        for m in LTR_REF.findall(str(t)):
            if m not in allowed:
                problems.append(f"names letter {m}, which is not the letter replied to or one of its references")
        for m in SR.PLACEHOLDER.findall(str(t)):
            if m.strip("[]").strip():
                problems.append(f"placeholder text left: {m!r} in {str(t)[:60]!r}")
    for p in req.get("photos", []):
        if not os.path.exists(p.get("file", "")):
            problems.append(f"photo not found: {p.get('file')}")
    if not req.get("body"):
        problems.append("no body text")
    return problems


def make(req_json, out):
    req = json.load(open(req_json, encoding="utf-8-sig"))
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    problems = check(req)
    if problems:
        raise SystemExit("Not written. Fix these first:\n- " + "\n- ".join(problems))
    z = zipfile.ZipFile(TEMPLATE)
    files = {n: z.read(n) for n in z.namelist()}
    x = files["word/document.xml"].decode("utf-8")
    num_id = (re.search(r'<w:numId w:val="(\d+)"/>', x) or [None, "1"])[1]
    b0, b1 = x.index("<w:body>") + len("<w:body>"), x.rfind("<w:sectPr")
    P = []
    P.append(label_line("Subject", req["subject"]))
    refs = [f"{req['ref']} dated {long(req.get('date'))}"] + list(req.get("references", []))
    for i, r in enumerate(refs):
        P.append(label_line("Reference" if i == 0 else "", f"({i + 1}) {r}") if len(refs) > 1 else label_line("Reference", r))
    P.append(para(after=60))
    P.append(para(req.get("salutation", "Dear Sir,"), justify=False))
    opening = req.get("opening") or (f"We acknowledge receipt of the Engineer's letter ref. {req['ref']} dated {long(req.get('date'))}"
                                     + (f", received on {long(req['received'])}" if req.get("received") and req.get("received") != req.get("date") else "")
                                     + f", regarding {req['subject']}.")
    P.append(para(opening))
    for b in req.get("body", []):
        if isinstance(b, str):
            P.append(para(b))
        elif "bullets" in b:
            P += [bullet(t, num_id) for t in b["bullets"]]
        elif "numbered" in b:
            P += [numbered(i + 1, t) for i, t in enumerate(b["numbered"])]
    if req.get("table"):
        if req["table"].get("intro"):
            P.append(para(req["table"]["intro"], keep=True))
        P.append(table(req["table"]["rows"]))
    P.append(para(req.get("closing", "We trust this addresses the Engineer's concerns.")))
    sig = req.get("signatory") or [CFG["incidents"]["people"]["project_manager"][0], LC.get("signatory_title", "Project Manager / Contractor's Representative")]
    P.append(para("Yours faithfully,", justify=False, after=600, keep=True))
    P.append(para(sig[0], bold=True, justify=False, after=0, keep=True))
    P.append(para(sig[1] if len(sig) > 1 else "", justify=False, after=0, keep=True))
    P.append(para(CFG.get("contractor_full", CFG["contractor"]), justify=False, after=200))
    enc = list(req.get("enclosures", []))
    if req.get("photos") and not any("Appendix A" in e for e in enc):
        enc.append("Appendix A – Site photographs")
    if enc:
        P.append(para("Enclosures:", bold=True, justify=False, after=60, keep=True))
        P += [numbered(i + 1, e) for i, e in enumerate(enc)]
    if req.get("photos"):
        P.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        P.append(para("Appendix A – Site Photographs", bold=True, justify=False))
        rels = files["word/_rels/document.xml.rels"].decode("utf-8")
        for i, ph in enumerate(req["photos"]):
            data, cx, cy = SR.prep_image(ph["file"], 15.5, 10.5)
            name, rid = f"media/reply_{i + 1}.jpeg", f"rIdReply{i + 1}"
            files["word/" + name] = data
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{name}"/></Relationships>')
            P.append(f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="60"/><w:jc w:val="center"/></w:pPr><w:r>{SR.drawing(rid, 950 + i, cx, cy)}</w:r></w:p>')
            P.append(para(ph["caption"], justify=False, after=200).replace("<w:pPr>", '<w:pPr><w:jc w:val="center"/>', 1))
        files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        ct = files["[Content_Types].xml"].decode("utf-8")
        if 'Extension="jpeg"' not in ct:
            files["[Content_Types].xml"] = re.sub(r"(<Types [^>]*>)", r'\1<Default Extension="jpeg" ContentType="image/jpeg"/>', ct, count=1).encode("utf-8")
    x = x[:b0] + "".join(P) + x[b1:]
    files["word/document.xml"] = x.encode("utf-8")
    import xml.dom.minidom
    for n in ("word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"):
        xml.dom.minidom.parseString(files[n])
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, data in files.items():
            zo.writestr(n, data)
    print(json.dumps({"written": out, "paragraphs": len(P), "photos": len(req.get("photos", []))}, ensure_ascii=False, indent=1))


def scaffold(letter_json, out):
    """A reply.json to finish by hand: header fields from the letter, one table row per requirement (status and
    evidence left for Claude to write from the evidence site uploaded)."""
    d = json.load(open(letter_json, encoding="utf-8-sig"))
    d = d.get("data", d)
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    req = {"ref": d.get("ref", ""), "date": d.get("issued", ""), "received": d.get("received", ""), "subject": d.get("subject", ""),
           "references": [], "body": [], "table": {"intro": "The status of each requirement of the Engineer's letter is as follows:",
                                                    "rows": [[r, "", ""] for r in d.get("requirements", [])]},
           "closing": "We trust this addresses the Engineer's requirements.", "enclosures": [], "photos": []}
    json.dump(req, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"written": out, "requirements": len(d.get("requirements", [])),
                      "todo": "write the body, the status and evidence of each row, enclosures and photos; leave out rows with no evidence"}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    {"make": make, "scaffold": scaffold}[sys.argv[1]](sys.argv[2], sys.argv[3])

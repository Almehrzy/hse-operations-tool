# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Founding and authorship record of the HSE Operations Tool (Word), dated, with the fingerprint of a full copy.

    python founding_doc.py <full-copy .zip> <out.docx>

Built on the reply template's page setup (Arial, page numbers); English record with an Arabic summary. Never overwrites.
"""
import datetime as dt, hashlib, json, os, re, sys, zipfile
from xml.sax.saxutils import escape
import letter_reply as LR
from config import CFG

AUTHOR = CFG.get("author", {})      # project.json author: {role, role_ar} at the time of the record

AR_RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:rtl/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'


def ar(text, bold=False):
    rpr = AR_RPR.replace("<w:rtl/>", "<w:b/><w:bCs/><w:rtl/>") if bold else AR_RPR
    return f'<w:p><w:pPr><w:bidi/><w:spacing w:after="100" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr><w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    zpath, out = sys.argv[1], sys.argv[2]
    if os.path.exists(out):
        raise SystemExit(f"Refusing to overwrite {out}")
    h = hashlib.sha256(open(zpath, "rb").read()).hexdigest()
    man = zipfile.ZipFile(zpath).read("HSE Operations Tool/MANIFEST.txt").decode("utf-8")
    nfiles = re.search(r"(\d+) files, ([\d,]+) bytes", man)
    made = re.search(r"Copy made: (.+?) \(", man).group(1)
    today = dt.date.today().strftime("%d %B %Y")
    P, H = LR.para, lambda t: LR.para(t, bold=True, justify=False, after=80, keep=True)
    tpl_doc = zipfile.ZipFile(LR.TEMPLATE).read("word/document.xml").decode("utf-8")
    num_id = (re.search(r'<w:numId w:val="(\d+)"/>', tpl_doc) or [None, "1"])[1]     # the template's bullet list
    B = lambda t: LR.bullet(t, num_id)
    body = [
        LR.para("HSE Operations Tool", bold=True, justify=False, after=40).replace('w:val="21"', 'w:val="36"'),
        LR.para("Founding and Authorship Record", bold=True, justify=False, after=240).replace('w:val="21"', 'w:val="28"'),
        LR.label_line("Author", "Khaled Almehrzy – founder, designer and developer of the tool"),
        LR.label_line("Contact", "khaled.almehrzy@gmail.com"),
        LR.label_line("Role", AUTHOR.get("role", "")),
        LR.label_line("Record date", today),
        P(""),
        H("1. What the tool is"),
        P("The HSE Operations Tool is a web application with supporting scripts that turns site safety records into the "
          "project's official documents, numbered, filed and entered in the registers. Site staff record on their phones; the "
          "tool keeps the records, deadlines, approvals and evidence, and produces the company's own forms and reports."),
        H("2. What it covers"),
        B("Safety observations, the weekly SOR trend analysis and worker violation notices (SVR) with the penalty matrix."),
        B("Consultant notices (SVN, NCR, SOR) and Engineer's letters: evidence from site, deadlines, close-out reports and replies."),
        B("Incidents under the project procedure (15-minute notification, initial report, investigation, witness statements, safety alert) and corrective and preventive actions (CAPA)."),
        B("Work permits (hot, cold, lifting, excavation, confined space) from request to approval, issue and closing, with a training check of every person on the permit."),
        B("Equipment: QR labels, daily and monthly inspections, third-party certificates and the TUV register."),
        B("Daily log (manpower, toolbox talks, training, inspections, dust suppression, waste) and the weekly and monthly HSE reports."),
        B("Worker competency lookup, team roles with server-enforced approvals, backups and integrity checks."),
        H("3. Development record"),
        P("The earliest dated records of the tool are of 14 October 2026. Main stages, as recorded in the tool's instructions file (CLAUDE.md):"),
        B("14 October 2026 – observation log, worker violations (SVR), consultant SVN register and close-out replies."),
        B("15 October 2026 – NCR and SOR, weekly SOR trend report, equipment inspections, work permits, team roles, daily log, weekly HSE report and monthly HSE report."),
        B("16 October 2026 – incidents and CAPA, safety alert, Engineer's letters and replies, environment records, worker competency lookup, permit training check, server-side permissions, backups and this record."),
        P("The author conceived the tool, specified each part, directed its development and approved the results. The code was "
          "written with the help of Claude, an AI assistant made by Anthropic, working under the author's instructions."),
        H("4. Fingerprint of the tool on this date"),
        P(f"A full copy of the tool was made on {made}: {nfiles.group(1) if nfiles else '?'} files ({nfiles.group(2) if nfiles else '?'} bytes), "
          f"each listed with its SHA-256 fingerprint in the copy's MANIFEST.txt. The fingerprint of the copy itself is:"),
        P(h, justify=False).replace("Arial", "Consolas"),
        P(f"File: {os.path.basename(zpath)}. The same file always gives the same fingerprint, so the copy shows what the tool contained on this date."),
        H("5. Ownership and licence"),
        P("The tool is proprietary. All rights, including copyright, belong to the author. Its use is governed by LICENSE.txt, "
          "kept with the tool's files: use is allowed only to the persons and projects the author names, and copying, modifying, "
          "republishing or selling it needs his written permission. Data entered in the tool belongs to the project that entered it."),
        H("6. Declaration"),
        P("I declare that I am the founder and author of the HSE Operations Tool described above, and that the record above is true."),
        P(""),
        P("Name: Khaled Almehrzy", justify=False, after=60),
        P("Signature: ______________________________", justify=False, after=60),
        P("Date: ______________________________", justify=False, after=240),
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
        ar("ملخص بالعربية", bold=True),
        ar(f"أداة HSE Operations Tool من فكرة وتصميم وتطوير خالد المحرزي (Khaled Almehrzy)"
           f"{'، ' + AUTHOR['role_ar'] if AUTHOR.get('role_ar') else ''}. أقدم السجلات المؤرخة للأداة بتاريخ 14 أكتوبر 2026. هذه الوثيقة مؤرخة في {today}، وتذكر بصمة نسخة كاملة من "
           f"الأداة (SHA-256) تثبت محتواها في هذا التاريخ. الأداة برنامج مملوك للمؤلف، وجميع حقوقه محفوظة له، واستخدامها وفق ملف "
           f"LICENSE.txt. كُتب الكود بمساعدة Claude، المساعد الذكي من شركة Anthropic، بتوجيه المؤلف وتحت إشرافه."),
        ar("للتسجيل الرسمي: تتيح الهيئة السعودية للملكية الفكرية التسجيل الاختياري لبرمجيات الحاسب عبر بوابتها saip.gov.sa (الخدمات ← حق المؤلف)، "
           "ويُرفق معها المصنف وهذه الوثيقة. الاستفسار على 920021421."),
    ]
    z = zipfile.ZipFile(LR.TEMPLATE)
    files = {n: z.read(n) for n in z.namelist()}
    x = files["word/document.xml"].decode("utf-8")
    b0, b1 = x.index("<w:body>") + len("<w:body>"), x.rfind("<w:sectPr")
    x = x[:b0] + "".join(body) + x[b1:]
    files["word/document.xml"] = x.encode("utf-8")
    import xml.dom.minidom
    xml.dom.minidom.parseString(files["word/document.xml"])
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, data in files.items():
            zo.writestr(n, data)
    print(json.dumps({"written": out, "copy": os.path.basename(zpath), "sha256": h}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

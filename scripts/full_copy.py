# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""A complete copy of the HSE Operations Tool in one zip: the page, scripts, templates, setup files, instructions,
the latest database backup (backups/db/<date>, with the photos under assets/) and the register backups, plus
MANIFEST.txt listing every file with its size and SHA-256 fingerprint and the date the copy was made.

    python full_copy.py [--db-backup backups\\db\\YYYY-MM-DD]
    python full_copy.py verify <a full-copy .zip>      what changed in the tool's files since that copy

Writes backups\\full-copies\\HSE-Operations-Tool-full-copy-<date>.zip (never overwrites). Test output, Python caches
and earlier full copies are left out. `verify` compares the page, scripts, templates and setup files now with the
fingerprints in the copy's MANIFEST.txt (the database backups are not compared: they are meant to change).
"""
import datetime as dt, hashlib, json, os, sys, zipfile
from config import ROOT, CFG

SKIP_DIRS = {"test-output", "__pycache__", "full-copies", ".claude"}


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def verify(zpath):
    man = zipfile.ZipFile(zpath).read("HSE Operations Tool/MANIFEST.txt").decode("utf-8")
    then = {}
    for line in man.splitlines():
        parts = line.split(None, 2)
        if len(parts) == 3 and len(parts[0]) == 64 and parts[1].isdigit():
            then[parts[2]] = parts[0]
    tool = lambda f: not f.startswith("backups")
    now = {}
    for dp, dns, fns in os.walk(ROOT):
        dns[:] = [d for d in dns if d not in SKIP_DIRS and d != "backups"]
        for f in fns:
            if not f.startswith("~$"):
                rel = os.path.relpath(os.path.join(dp, f), ROOT)
                now[rel] = sha(os.path.join(ROOT, rel))
    old = {k: v for k, v in then.items() if tool(k)}
    changed = sorted(k for k in old if k in now and now[k] != old[k])
    removed = sorted(k for k in old if k not in now)
    added = sorted(k for k in now if k not in old)
    print(json.dumps({"compared_with": os.path.basename(zpath), "unchanged": len(old) - len(changed) - len(removed),
                      "changed": changed, "added": added, "removed": removed}, ensure_ascii=False, indent=1))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) > 2 and sys.argv[1] == "verify":
        return verify(sys.argv[2])
    opt = dict(zip(sys.argv[1::2], sys.argv[2::2]))
    now = dt.datetime.now()
    dbdirs = sorted(d for d in os.listdir(os.path.join(ROOT, "backups", "db"))) if os.path.isdir(os.path.join(ROOT, "backups", "db")) else []
    db = opt.get("--db-backup") or (os.path.join("backups", "db", dbdirs[-1]) if dbdirs else "")
    files = []
    for dp, dns, fns in os.walk(ROOT):
        rel_dir = os.path.relpath(dp, ROOT)
        dns[:] = [d for d in dns if d not in SKIP_DIRS and not (rel_dir == os.path.join("backups", "db") and os.path.join(rel_dir, d) != db)]
        for f in fns:
            if f.startswith("~$") or f.startswith("HSE Operations Tool - Founding"):     # the record names the copy, so it stays outside it
                continue
            files.append(os.path.relpath(os.path.join(dp, f), ROOT))
    files.sort()
    out_dir = os.path.join(ROOT, "backups", "full-copies")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"HSE-Operations-Tool-full-copy-{now:%Y-%m-%d}.zip")
    if os.path.exists(out):
        out = out[:-4] + f"-{now:%H%M}.zip"
    lines = ["HSE Operations Tool – full copy",
             "Designed and developed by Khaled Almehrzy © 2026 (khaled.almehrzy@gmail.com)",
             f"Copy made: {now:%Y-%m-%d %H:%M} (local time of the PC that made it)",
             f"Published page: {CFG.get('page_url', '')}",
             f"Database backup included: {db or 'none'}",
             "", "Each file below with its size in bytes and SHA-256 fingerprint. The same file always gives the same",
             "fingerprint, so this list shows what the tool contained on the date above.", ""]
    total = 0
    for f in files:
        p = os.path.join(ROOT, f)
        size = os.path.getsize(p)
        total += size
        lines.append(f"{sha(p)}  {size:>10}  {f}")
    lines += ["", f"{len(files)} files, {total:,} bytes"]
    readme = ("HOW TO USE THIS COPY / طريقة استخدام هذه النسخة\n\n"
              "1. Unzip into a folder (for example D:\\Claude HSE\\Site Observations Tool).\n"
              "2. The page is site-observations.html: publish it as a new artifact, or update the existing one, with Claude.\n"
              "3. The database is in backups\\db\\<date>\\<collection>\\<id>.json and the photos in backups\\db\\<date>\\assets\\.\n"
              "   Restore: ask Claude to write each file back with ArtifactData (batch set), as in CLAUDE.md 'Database backup'.\n"
              "4. CLAUDE.md holds every procedure and decision; project.json, permits.json, equipment.json and monthly.json the setup.\n\n"
              "١. فك الضغط في مجلد.\n٢. الصفحة هي site-observations.html: تُنشر كأداة جديدة أو تُحدِّث الحالية عبر Claude.\n"
              "٣. قاعدة البيانات في backups\\db\\<التاريخ> والصور في assets، ويستعيدها Claude بأمر الكتابة من النسخة.\n"
              "٤. ملف CLAUDE.md فيه كل الإجراءات والقرارات.\n")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(os.path.join(ROOT, f), os.path.join("HSE Operations Tool", f))
        z.writestr("HSE Operations Tool/MANIFEST.txt", "\n".join(lines) + "\n")
        z.writestr("HSE Operations Tool/README - restore.txt", readme)
    print(json.dumps({"written": out, "files": len(files), "bytes": total, "zip_bytes": os.path.getsize(out), "zip_sha256": sha(out), "db_backup": db},
                     ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

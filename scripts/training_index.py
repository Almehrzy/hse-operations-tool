# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Worker competency lookup for the page, from the HSE training database (sheet "HSE Training Data Entry").

    python training_index.py <training database .xlsx> <out dir> [--chunk 500]

Writes <out dir>/workersIndex/chunk-NN.json ({"list": [...]}) and <out dir>/settings/training.json (the training topics,
the source file and its date) for ArtifactData. Each worker: f = file no., n = name, c = company, d = designation,
x = Iqama validity date, t = {topic index: date of the training}. Iqama and insurance numbers, nationality and photos
are left out: the page only needs to say who is inducted and trained for what, and until when the Iqama is valid.
"""
import datetime as dt, json, os, re, sys, warnings
import openpyxl

warnings.filterwarnings("ignore")


def iso(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    # dates typed as text, as found in the database: "May-20-26", "June-28-26", "20.Jul.26", "08.10.26", "29/6/26", "20-07-2026"
    s = str(v or "").strip().replace("Sept", "Sep")
    for f in ("%b-%d-%y", "%B-%d-%y", "%b-%d-%Y", "%B-%d-%Y", "%d.%b.%y", "%d.%b.%Y", "%d-%b-%y", "%d-%b-%Y",
              "%d.%m.%y", "%d.%m.%Y", "%d-%m-%y", "%d-%m-%Y", "%d/%m/%y", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(s, f).date().isoformat()
        except ValueError:
            pass
    return ""


def clean(v):
    s = re.sub(r"\s+", " ", str(v or "")).strip()
    return "" if s in ("-", "None", "N/A", "NA") else s


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    src, out = sys.argv[1], sys.argv[2]
    opt = dict(zip(sys.argv[3::2], sys.argv[4::2]))
    size = int(opt.get("--chunk", 500))
    ws = openpyxl.load_workbook(src, read_only=True, data_only=True)["HSE Training Data Entry"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = next(i for i, r in enumerate(rows[:10]) if r and any(str(c or "").strip().upper().startswith("EMPLOYEE'S NAME") for c in r))
    names = rows[hdr]
    col = lambda word: next(i for i, c in enumerate(names) if str(c or "").strip().upper().startswith(word))
    cf, cn, cc, cd, cx = col("FILE NO"), col("EMPLOYEE'S NAME"), col("COMPANY"), col("DESIGNATION"), col("IQAMA VALIDITY")
    top = rows[hdr - 1]
    first_topic = next(i for i, c in enumerate(top) if str(c or "").strip().upper().startswith("HSE INDUCTION"))
    topics = [re.sub(r"\s+", " ", str(c or "")).strip() for c in top[first_topic:] if str(c or "").strip()]
    workers, dropped = [], 0
    for r in rows[hdr + 1:]:
        if not r or len(r) <= cn or not clean(r[cn]):
            continue
        t = {}
        for k in range(len(topics)):
            v = r[first_topic + k] if first_topic + k < len(r) else None
            d = iso(v)
            if d:
                t[str(k)] = d
            elif v not in (None, "", "-", 0):
                dropped += 1
        workers.append({"f": clean(r[cf]), "n": clean(r[cn]).upper(), "c": clean(r[cc]).upper(), "d": clean(r[cd]).title(), "x": iso(r[cx]), "t": t})
    os.makedirs(os.path.join(out, "workersIndex"), exist_ok=True)
    os.makedirs(os.path.join(out, "settings"), exist_ok=True)
    chunks = [workers[i:i + size] for i in range(0, len(workers), size)]
    for i, ch in enumerate(chunks):
        json.dump({"list": ch}, open(os.path.join(out, "workersIndex", f"chunk-{i + 1:02d}.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    as_of = dt.date.fromtimestamp(os.path.getmtime(src)).isoformat()
    json.dump({"topics": topics, "source": os.path.basename(src), "asOf": as_of, "count": len(workers), "chunks": len(chunks)},
              open(os.path.join(out, "settings", "training.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    sizes = [os.path.getsize(os.path.join(out, "workersIndex", f"chunk-{i + 1:02d}.json")) for i in range(len(chunks))]
    no_ind = sum(1 for w in workers if "0" not in w["t"])
    today = dt.date.today().isoformat()
    print(json.dumps({"workers": len(workers), "topics": len(topics), "chunks": len(chunks), "chunk_bytes": sizes, "as_of": as_of,
                      "no_induction_date": no_ind, "iqama_expired": sum(1 for w in workers if w["x"] and w["x"] < today),
                      "cells_not_dates": dropped}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

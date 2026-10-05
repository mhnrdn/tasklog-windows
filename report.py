#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
report.py - Ubah log aktivitas jadi laporan harian + bulanan.

Pakai:
  python report.py                    -> laporan hari ini
  python report.py --day 2026-10-05
  python report.py --month 2026-10      -> rekap 1 bulan
  python report.py --stdout             -> cetak saja
  python report.py --audit             -> TELUSURI sumber tiap baris (buat periksa)

Output:
  Laporan/Laporan_Harian_YYYY-MM-DD.md
  Laporan/Ringkas_YYYY-MM-DD.txt     (siap tempel Google Calendar)
  Laporan/Audit_YYYY-MM-DD.txt
  Laporan/Laporan_Bulanan_YYYY-MM.md
"""
import os
import re
import sys
import json
import time
import sqlite3
import datetime as dt
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
try:
    import tlpaths
except Exception:
    tlpaths = None

if tlpaths:
    BASE = tlpaths.data_dir()
    DBDIR = tlpaths.state_dir()
    OUT = tlpaths.report_dir()
    MANUAL = tlpaths.manual_path()
else:
    DBDIR = os.path.join(os.environ.get("LOCALAPPDATA", BASE), "hermes", "tasklog")
    OUT = os.path.join(BASE, "Laporan")
    MANUAL = os.path.join(BASE, "manual_entries.json")
DB = os.path.join(DBDIR, "activity.db")

NOISE = {"Aktivitas Aplikasi"}
EXCEL_PROCS = {"excel.exe"}          # kerja Excel = sinyal utama
DOC_PROCS = {"excel.exe", "winword.exe", "powerpnt.exe", "acrobat.exe",
             "acroread.exe", "sumatrapdf.exe", "pdfxedit.exe"}

HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
         "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

MIN_SEC = 45          # rekaman < 45 detik = noise (pindah-pindah window)
NONPROJ_MIN = 900     # blok "Lain-lain" muncul kalau >= 15 menit
NONPROJ = "Lain-lain / Non Proyek"

ROOT_PREFIXES = None      # dihitung dari config (lihat _root_prefixes())


def _root_prefixes():
    """Awalan path yang dipotong saat menampilkan folder. Diambil dari tlpaths."""
    global ROOT_PREFIXES
    if ROOT_PREFIXES is None:
        try:
            ROOT_PREFIXES = tlpaths.root_prefixes()
        except Exception:
            ROOT_PREFIXES = [os.path.expanduser("~").lower() + os.sep]
    return ROOT_PREFIXES


# ------------------------------------------------------------------ util
def human(sec):
    m = int(sec // 60)
    if m < 60:
        return f"{m} menit"
    return f"{m // 60} jam {m % 60} menit"


def is_filepath(s):
    """True kalau context berupa path file (lokal atau URL OneDrive)."""
    if not s:
        return False
    if s.lower().startswith(("http://", "https://")):
        return True
    if "\\" not in s and "/" not in s:
        return False
    return "." in re.split(r"[\\/]+", s.strip())[-1]


CARET = {"^j": ", ", "^m": " ", "^i": " ", "^k": " ", "^l": " "}
_IDXC = {"t": 0, "m": None}


def file_index():
    """Indeks nama-file -> path lokal (dibuat tracker.py)."""
    if _IDXC["m"] is not None and time.time() - _IDXC["t"] < 600:
        return _IDXC["m"]
    m = {}
    try:
        with open(os.path.join(DBDIR, "folder_index.json"), "r", encoding="utf-8") as f:
            m = json.load(f)
    except Exception:
        m = {}
    _IDXC["m"] = m
    _IDXC["t"] = time.time()
    return m


def decode_caret(s):
    out = s
    for k, v in CARET.items():
        out = out.replace(k, v).replace(k.upper(), v)
    return out.strip()


def url_to_parts(s):
    """URL OneDrive -> daftar segmen path."""
    from urllib.parse import unquote
    body = unquote(s.split("://", 1)[1])
    parts = [x for x in body.split("/") if x]
    if len(parts) > 2 and "docs.live.net" in parts[0].lower():
        parts = parts[2:]
    elif parts:
        parts = parts[1:]
    return [decode_caret(x) for x in parts]


def short_folder(folder):
    """Folder panjang -> ringkas tapi tetap kelihatan folder proyeknya."""
    parts = [p for p in folder.split("\\") if p]
    if len(parts) <= 3:
        return " \\ ".join(parts)
    return f"{parts[0]} \\ ... \\ {parts[-1]}"


def pretty(s):
    """Path panjang -> nama ringkas."""
    if not s:
        return ""
    s = s.strip().strip('"')
    if not is_filepath(s):
        parts = [p for p in re.split(r"[\\/]+", s) if p]
        return " / ".join(parts[-2:]) if len(parts) >= 2 else s
    return split_path(s)[1]


def split_path(s):
    """Path (lokal atau URL OneDrive) -> (folder tanpa root, nama file)."""
    s = s.strip()
    if s.lower().startswith(("http://", "https://")):
        parts = url_to_parts(s)
        fn = parts[-1] if parts else ""
        hit = file_index().get(fn.lower())
        if hit:
            return split_path(hit)
        return " \\ ".join(parts[:-1][-3:]), fn
    parts = [p for p in re.split(r"[\\/]+", s) if p]
    fn = parts[-1] if parts else ""
    folder = "\\".join(parts[:-1])
    low = folder.lower()
    for pre in _root_prefixes():
        if low.startswith(pre):
            folder = folder[len(pre):]
            break
    return folder, fn


def load_manual():
    try:
        with open(MANUAL, "r", encoding="utf-8") as f:
            j = json.load(f)
        return {k: v for k, v in j.items() if not k.startswith("_")}
    except Exception:
        return {}


def rows(day):
    if not os.path.exists(DB):
        return []
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    r = c.execute("SELECT * FROM activity WHERE day=? ORDER BY start_ts", (day,)).fetchall()
    c.close()
    return [dict(x) for x in r]


# ------------------------------------------------------------------ agregasi
def build_day(day):
    """
    Return dict:
      lines   : 1 baris per proyek (urut durasi terbesar)
      total   : detik aktif
      excel   : {'sec':..., 'focus':proyek atau None, 'focus_sec':..., 'focus_folder':...}
      ev      : bukti/sumber tiap baris (untuk --audit)
    """
    raw = rows(day)
    data = [d for d in raw if d["seconds"] >= MIN_SEC]

    proj_acts = defaultdict(lambda: defaultdict(lambda: {"sec": 0, "samples": [], "first": None, "last": None}))
    proj_files = defaultdict(lambda: defaultdict(int))     # proyek -> {namafile: detik}
    proj_folders = defaultdict(lambda: defaultdict(int))   # proyek -> {folder: detik}
    excel_sec = 0
    ev = []

    for d in data:
        proj = d["project"] or NONPROJ
        act = d["act"] or d["kind"]
        proc = (d["proc"] or "").lower()
        ctx = d["context"] or ""

        slot = proj_acts[proj][act]
        slot["sec"] += d["seconds"]
        if len(slot["samples"]) < 3:
            s = pretty(ctx or d["title"])
            if s and s not in slot["samples"]:
                slot["samples"].append(s[:90])
        t0, t1 = d["start_ts"][11:16], d["end_ts"][11:16]
        if slot["first"] is None or t0 < slot["first"]:
            slot["first"] = t0
        if slot["last"] is None or t1 > slot["last"]:
            slot["last"] = t1

        # sinyal Excel + lokasi folder file
        if proc in EXCEL_PROCS:
            excel_sec += d["seconds"]
            if is_filepath(ctx):
                folder, fn = split_path(ctx)
                proj_files[proj][fn] += d["seconds"]
                if folder:
                    proj_folders[proj][folder] += d["seconds"]

        ev.append({"src": f"DB#{d['id']}", "jam": f"{t0}-{t1}", "detik": d["seconds"],
                   "proc": d["proc"] or "-", "act": act, "project": proj,
                   "path": ctx or d["title"]})

    # entri manual (pekerjaan di luar laptop)
    for i, m in enumerate(load_manual().get(day, [])):
        proj = m.get("project", NONPROJ)
        act = m.get("act", "Catatan Manual")
        slot = proj_acts[proj][act]
        secs = int(m.get("minutes", 0)) * 60
        slot["sec"] += secs
        slot["manual"] = True
        if m.get("note"):
            slot["samples"].append(m["note"][:90])
        if m.get("start"):
            slot["first"] = m["start"]
        if m.get("end"):
            slot["last"] = m["end"]
        ev.append({"src": f"manual_entries.json[{i}]", "jam": f"{m.get('start','-')}-{m.get('end','-')}",
                   "detik": secs, "proc": "(bukan laptop)", "act": act, "project": proj,
                   "path": m.get("note", "")})

    lines = []
    for proj in sorted(proj_acts, key=lambda p: -sum(v["sec"] for v in proj_acts[p].values())):
        acts = sorted(proj_acts[proj].items(), key=lambda kv: -kv[1]["sec"])
        aname, ainfo = acts[0]
        folder = ""
        if proj_folders[proj]:
            folder = sorted(proj_folders[proj].items(), key=lambda kv: -kv[1])[0][0]
        files = sorted(proj_files[proj], key=lambda f: -proj_files[proj][f])

        if files:
            shown = ", ".join(files[:3]) + ("..." if len(files) > 3 else "")
            detail = f" — {len(files)} file Excel ({shown})"
        elif ainfo["samples"]:
            detail = f' — {ainfo["samples"][0]}'
        else:
            detail = ""

        span = f'  ({ainfo["first"]}–{ainfo["last"]})' if ainfo.get("first") and ainfo.get("last") else ""
        extra = []
        if folder:
            extra.append(f"folder: {short_folder(folder)}")
        if ainfo.get("manual"):
            extra.append("sumber: catatan manual (bukan dari laptop)")
        for k, v in acts[1:4]:
            tag = " [manual]" if v.get("manual") else ""
            extra.append(f'{k} ({human(v["sec"])}){tag}')

        lines.append({
            "project": proj, "act": aname, "detail": detail, "span": span,
            "sec": ainfo["sec"], "total_sec": sum(v["sec"] for v in proj_acts[proj].values()),
            "folder": folder, "n_files": len(files), "files": files,
            "excel_sec": sum(proj_files[proj].values()),
            "manual": bool(ainfo.get("manual")) and all(v.get("manual") for v in proj_acts[proj].values()),
            "extra": extra,
        })

    lines = [L for L in lines if not (L["project"] == NONPROJ and L["total_sec"] < NONPROJ_MIN)]

    excel_rank = sorted((L for L in lines if L["excel_sec"] > 0), key=lambda L: -L["excel_sec"])
    focus = excel_rank[0] if excel_rank else None

    return {"day": day, "lines": lines, "total": sum(d["seconds"] for d in data) + sum(
        int(m.get("minutes", 0)) * 60 for m in load_manual().get(day, [])),
        "excel": {"sec": excel_sec, "focus": focus["project"] if focus else None,
                  "focus_sec": focus["excel_sec"] if focus else 0,
                  "focus_folder": short_folder(focus["folder"]) if focus else ""},
        "ev": ev}


# ------------------------------------------------------------------ render
def day_md(day, for_calendar=False):
    R = build_day(day)
    lines, total, X = R["lines"], R["total"], R["excel"]
    d = dt.date.fromisoformat(day)
    head = f"{HARI[d.weekday()]}, {d.day} {BULAN[d.month]} {d.year}"

    if for_calendar:
        out = [f"Laporan Pekerjaan {head}"]
        for L in lines:
            tail = ""
            if L["n_files"]:
                tail = f" - {L['n_files']} file Excel"
                if L["folder"]:
                    tail += f" (folder {short_folder(L['folder'])})"
            out.append(f"{L['project']} : {L['act']}{tail}")
        return "\n".join(out) if len(out) > 1 else out[0] + "\n(tidak ada aktivitas tercatat)"

    out = [f"## {head}", f"_Total waktu aktif tercatat: {human(total)}_", ""]

    if X["focus"]:
        fl = ""
        if X["focus_folder"]:
            fl = f", folder `{X['focus_folder']}`"
        out.append(f"**Dominan: {X['focus']} — {human(X['focus_sec'])} di Excel{fl}**")
    if X["sec"]:
        out.append(f"_Total kerja Excel hari ini: {human(X['sec'])}_")
    if X["focus"] or X["sec"]:
        out.append("")

    if not lines:
        out.append("- (belum ada aktivitas tercatat)")
    for L in lines:
        out.append(f"- **{L['project']}** : {L['act']}{L['detail']}  ({human(L['sec'])}){L['span']}")
        for e in L["extra"]:
            out.append(f"    - {e}")
    return "\n".join(out)


def audit_txt(day):
    """Telusuri sumber tiap baris laporan."""
    R = build_day(day)
    out = [f"AUDIT LAPORAN {day}",
           f"Sumber data: {DB}",
           f"Sumber manual: {MANUAL}",
           f"Filter: rekaman < {MIN_SEC} detik dibuang",
           ""]
    out.append(f"{len(R['ev'])} rekaman dipakai:")
    out.append("-" * 100)
    for e in sorted(R["ev"], key=lambda x: x["jam"]):
        out.append(f"{e['src']:28s} {e['jam']:12s} {e['detik']:>6}s  {e['proc']:14s} "
                   f"{e['act'][:34]:34s} -> {e['project']}")
        if e["path"]:
            out.append(f"{'':28s} path: {e['path']}")
    out.append("-" * 100)
    dropped = [d for d in rows(day) if d["seconds"] < MIN_SEC]
    out.append(f"Dibuang karena < {MIN_SEC}s: {len(dropped)} rekaman "
               f"({sum(d['seconds'] for d in dropped)} detik)")
    out.append("")
    out.append("HASIL:")
    out.append(day_md(day))
    return "\n".join(out)


def month_md(month):
    y, m = [int(x) for x in month.split("-")]
    first = dt.date(y, m, 1)
    last = dt.date(y + (m == 12), (m % 12) + 1, 1) - dt.timedelta(days=1)
    out = [f"# Laporan Bulanan — {BULAN[m]} {y}", ""]
    d = first
    while d <= last:
        R = build_day(d.isoformat())
        if R["lines"]:
            out.append(day_md(d.isoformat()))
            out.append("")
        d += dt.timedelta(days=1)
    return "\n".join(out)


# ------------------------------------------------------------------ CLI
def main():
    a = sys.argv[1:]
    day = dt.date.today().isoformat()
    month = day[:7]
    if "--day" in a:
        day = a[a.index("--day") + 1]
    if "--month" in a:
        month = a[a.index("--month") + 1]

    if "--audit" in a:
        txt = audit_txt(day)
        if "--stdout" in a:
            print(txt)
            return
        os.makedirs(OUT, exist_ok=True)
        p = os.path.join(OUT, f"Audit_{day}.txt")
        with open(p, "w", encoding="utf-8") as f:
            f.write(txt + "\n")
        print(p)
        return

    if "--month" in a:
        txt = month_md(month)
        if "--stdout" in a:
            print(txt)
            return
        os.makedirs(OUT, exist_ok=True)
        p = os.path.join(OUT, f"Laporan_Bulanan_{month}.md")
        with open(p, "w", encoding="utf-8") as f:
            f.write(txt)
        print(p)
        return

    if "--stdout" in a:
        print(day_md(day))
        print("\n----- SIAP TEMPEL KE GOOGLE CALENDAR -----\n")
        print(day_md(day, for_calendar=True))
        return

    os.makedirs(OUT, exist_ok=True)
    p1 = os.path.join(OUT, f"Laporan_Harian_{day}.md")
    with open(p1, "w", encoding="utf-8") as f:
        f.write(day_md(day) + "\n")
    p2 = os.path.join(OUT, f"Ringkas_{day}.txt")
    with open(p2, "w", encoding="utf-8") as f:
        f.write(day_md(day, for_calendar=True) + "\n")
    p3 = os.path.join(OUT, f"Audit_{day}.txt")
    with open(p3, "w", encoding="utf-8") as f:
        f.write(audit_txt(day) + "\n")
    print(p1)
    print(p2)
    print(p3)


if __name__ == "__main__":
    main()

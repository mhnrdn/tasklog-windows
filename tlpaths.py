#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
tlpaths.py - Penentu lokasi file. Dipakai tracker.py, report.py, app.py.

Dua lokasi:
  DATA  : folder yang dilihat user -> projects.json, exclude.json,
          manual_entries.json, Laporan\\, LOG\\  (sebelah file .exe)
  STATE : data mentah + indeks -> %LOCALAPPDATA%\\TaskLog\\
"""
import os
import shutil
import sys


def is_frozen():
    return getattr(sys, "frozen", False)


def app_dir():
    """Folder tempat .exe (atau script) berada."""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def data_dir():
    # TaskLog_data\ di sebelah .exe. Bisa diarahkan lain lewat env (untuk uji coba).
    override = os.environ.get("TASKLOG_DATA_DIR")
    d = override or os.path.join(app_dir(), "TaskLog_data")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        d = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "TaskLog_data")
        os.makedirs(d, exist_ok=True)
    return d


def state_dir():
    override = os.environ.get("TASKLOG_STATE_DIR")
    if override:
        os.makedirs(override, exist_ok=True)
        return override
    d = os.path.join(os.environ.get("LOCALAPPDATA") or data_dir(), "TaskLog")
    os.makedirs(d, exist_ok=True)
    migrate_old(d)
    return d


OLD_STATE_DIRS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "hermes", "tasklog"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "TaskLog"),
]
_MIG = ("activity.db", "folder_index.json", "tracker_status.json", "tracker.pid")


def migrate_old(new):
    """Pindahkan data dari lokasi lama (versi pertama) kalau ada."""
    for old in OLD_STATE_DIRS:
        if not old or os.path.abspath(old) == os.path.abspath(new) or not os.path.isdir(old):
            continue
        for fn in _MIG:
            src, dst = os.path.join(old, fn), os.path.join(new, fn)
            if os.path.exists(src) and not os.path.exists(dst):
                try:
                    shutil.copy2(src, dst)
                except Exception:
                    pass


# ---------------------------------------------------------------- config files
def cfg_path():
    return os.path.join(data_dir(), "projects.json")


def exc_path():
    return os.path.join(data_dir(), "exclude.json")


def manual_path():
    return os.path.join(data_dir(), "manual_entries.json")


def report_dir():
    d = os.path.join(data_dir(), "Laporan")
    os.makedirs(d, exist_ok=True)
    return d


def log_dir():
    d = os.path.join(data_dir(), "LOG")
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------- roots default
def _onedrive_dirs():
    """Cari folder OneDrive, termasuk kalau dipindah ke drive/lokasi lain."""
    cands = []
    for env in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        v = os.environ.get(env)
        if v:
            cands.append(v)
    for drive in ("C:", "D:", "E:", "F:", "G:"):
        root = drive + "\\"
        if not os.path.isdir(root):
            continue
        try:
            for n in os.listdir(root):
                low = n.lower()
                if low == "onedrive" or low.startswith("onedrive - "):
                    cands.append(os.path.join(root, n))
        except Exception:
            pass
    return cands


def default_roots():
    """Folder yang dipindai untuk cari nama proyek. Bisa diubah di projects.json."""
    home = os.path.expanduser("~")
    cands = _onedrive_dirs() + [
        os.path.join(home, "OneDrive"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive", "Documents"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Documents"),
        os.path.join(home, "Downloads"),
        os.path.join(home, "Projects"),
        os.path.join(home, "Proyek"),
    ]
    for od in _onedrive_dirs():                 # subfolder standar di dalam OneDrive
        cands += [os.path.join(od, "Desktop"), os.path.join(od, "Documents"),
                  os.path.join(od, "Dokumen"), os.path.join(od, "Dokumen", "Projects")]
    seen, out = set(), []
    for c in cands:
        if not c:
            continue
        c2 = os.path.normpath(c)
        low = c2.lower()
        if low in seen or not os.path.isdir(c2):
            continue
        seen.add(low)
        out.append(c2.replace("\\", "/"))
    return out


def root_prefixes():
    """Awalan path yang dibuang saat menampilkan folder (biar ringkas)."""
    pre = []
    for r in default_roots() + _extra_roots():
        pre.append(os.path.normpath(r).lower() + os.sep)
    pre.append(os.path.expanduser("~").lower() + os.sep)
    return sorted(set(pre), key=len, reverse=True)


def _extra_roots():
    try:
        cfg = read_json(cfg_path(), {})
        return [x for x in cfg.get("_roots_scan", []) if x]
    except Exception:
        return []


# ---------------------------------------------------------------- default configs
DEFAULT_PROJECTS = {"_comment": "Peta kata kunci -> nama proyek. Urutan = prioritas.",
                    "_roots_scan": [], "projects": []}

DEFAULT_EXCLUDE = {
    "_comment": "Yang TIDAK dicatat. proses = nama exe, kata = potongan judul window.",
    "proses": ["chrome.exe"],
    "kata": ["m-banking", "mbanking", "internet banking", "klikbca", "mybca",
             "bca mobile", "brimo", "livin by mandiri", "mandiri online",
             "bni mobile", "wondr by bni", "jenius", "seabank", "bank jago",
             "bitwarden", "1password", "lastpass", "keepass", "password manager",
             "gopay", "ovo cash", "dana wallet", "shopeepay", "pin atm", "cvv", "otp"],
}

DEFAULT_MANUAL = {
    "_comment": "Pekerjaan di luar laptop. Format tanggal YYYY-MM-DD. minutes = durasi.",
    "_contoh": {"2000-01-01": [{"project": "Contoh Proyek", "act": "Kunjungan Lapangan",
                                "minutes": 60, "start": "09:00", "end": "10:00",
                                "note": "contoh"}]},
}


# ---------------------------------------------------------------- auto-deteksi proyek
import json
import re

STOP_TOKENS = {
    "data", "file", "files", "gambar", "baru", "update", "master", "laporan", "backup",
    "arsip", "archive", "temp", "revisi", "final", "draft", "pdf", "excel", "word",
    "documents", "desktop", "downloads", "onedrive", "program", "programs", "instalasi",
    "instal", "the", "and", "for", "with", "new", "old", "copy", "test", "coba", "lain",
    "umum", "general", "misc", "other", "others", "project", "proyek", "kerja", "work",
    "my", "document", "doc", "folder", "share", "shared", "public", "personal", "ojk",
    "bank", "kantor", "pusat", "daerah", "cabang", "tahun", "bulan", "hari", "materi",
    "video", "videos", "music", "pictures", "pictures", "music", "compressed", "apps",
    "app", "software", "setup", "installer", "driver", "drivers", "backup1", "lama",
    "kadaluarsa", "recycle", "recovery", "system", "windows", "user", "users",
    "one", "drive", "gdrive", "google", "teams", "chat", "screenshots", "capture",
}

# folder yang jelas bukan proyek
SKIP_FOLDERS = {
    "my music", "my pictures", "my videos", "music", "pictures", "videos",
    "documents", "desktop", "downloads", "programs", "compressed",
    "appdata", "application data", "contacts", "favorites", "links", "saved games",
    "searches", "3d objects", "onedrive", "google drive", "camera roll",
    "microsoft copilot chat files", "personal vault", "attachments", "templates",
    "public", "shared", "temp", "tmp", "new folder", "lain lain", "lainlain",
    # aplikasi / utilitas, bukan proyek
    "zoomit", "program", "cubicost", "qubicost", "memu download", "lain-lain",
    "sketchup", "autocad", "planswift", "gstarcad",
    "office", "windows", "installer", "setup", "drivers", "kadaluarsa", "lama",
    "screenshots", "capture", "tutorial", "pelatihan", "kursus", "materi",
}

# kalau folder memuat kata ini di namanya -> bukan proyek
GENERIC_WORDS = {"pelatihan", "tutorial", "kursus", "materi", "backup", "instalasi",
                 "installer", "setup", "drivers", "template", "sample", "contoh",
                 "coba", "test", "latihan", "sertifikat", "foto", "photo", "gambar"}


def clean_name(folder):
    """'23-09 OJK KR-5 SUMBAGUT_MEDAN' -> 'OJK KR-5 SUMBAGUT_MEDAN'."""
    n = re.sub(r"^\s*\d{2,4}\s*[-._]\s*\d{1,2}\s+", "", folder)
    n = re.sub(r"^\s*\d+\s*[-.)]\s*", "", n)
    n = n.replace("_", " ").strip()
    return re.sub(r"\s{2,}", " ", n) or folder


def tokens(folder):
    n = clean_name(folder).lower()
    out = []
    for t in re.split(r"[^0-9a-z]+", n):
        if len(t) >= 3 and not t.isdigit():
            out.append(t)
    return out


CODE_RE = re.compile(r"^\s*\d{2,4}\s*[-._]\s*\d{1,2}\b")          # 23-09, 2026-02
NAME_RE = re.compile(r"[A-Za-z]{3,}[\s\-_.][A-Za-z]{3,}")        # 2 kata sungguhan


def project_score(folder):
    """Makin tinggi = makin mirip folder proyek."""
    s = 0
    if CODE_RE.match(folder):          s += 4       # ada kode proyek 23-09
    if re.search(r"\b(1[89]|20)\d{2}\b", folder):  s += 1
    if NAME_RE.search(folder):         s += 1       # minimal 2 kata bermakna
    if re.search(r"\b(proyek|project|ra|rab|bq|boq|tender|paket)\b", folder, re.I):
        s += 2
    return s


def detect_projects(roots, max_projects=40, max_common=3, min_score=1):
    """Pindai folder tingkat-1 di tiap root -> daftar proyek + kata kunci."""
    candidates = []
    for r in roots or []:
        if not r or not os.path.isdir(r):
            continue
        try:
            names = sorted(os.listdir(r))
        except Exception:
            continue
        for n in names:
            if n.startswith((".", "_", "~", "$", "#")):
                continue
            key = re.sub(r"[^a-z0-9]+", " ", n.lower()).strip()
            key2 = re.sub(r"[^a-z0-9]+", " ", clean_name(n).lower()).strip()
            if key in SKIP_FOLDERS or key2 in SKIP_FOLDERS:
                continue
            if key == os.path.basename(os.path.expanduser("~")).lower():
                continue                     # folder nama user sendiri, bukan proyek
            if any(w in key.split() or w in key2.split() for w in GENERIC_WORDS):
                continue
            if re.search(r"\d+\.\d+\.\d+", n):        # versi aplikasi, mis. 10.2.1.385
                continue
            if not os.path.isdir(os.path.join(r, n)):
                continue
            candidates.append(n)
            if len(candidates) >= 400:
                break

    # buang kata yang muncul di banyak kandidat (terlalu umum)
    freq = {}
    for c in candidates:
        for t in set(tokens(c)):
            freq[t] = freq.get(t, 0) + 1
    common = {t for t, n in freq.items() if n > max_common}

    scored = []
    for c in candidates:
        keys = [t for t in tokens(c) if t not in common and t not in STOP_TOKENS]
        full = clean_name(c).lower().strip()
        if full and len(full) >= 6:
            keys.insert(0, full)
        keys = list(dict.fromkeys(keys))[:6]
        if not keys:
            continue
        sc = project_score(c)
        if sc < min_score:
            continue
        scored.append((sc, len(c), {"name": clean_name(c), "keys": keys}))

    scored.sort(key=lambda x: (-x[0], x[1]))
    return [p for _, _, p in scored[:max_projects]]


def read_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return json.loads(json.dumps(default))


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


def _legacy_cfg_dirs():
    """Tempat config versi lama / lokasi lain, untuk dipindahkan otomatis."""
    la = os.environ.get("LOCALAPPDATA", "")
    home = os.path.expanduser("~")
    cands = [app_dir()]
    if la:
        cands += [os.path.join(la, "hermes", "tasklog"), os.path.join(la, "TaskLog")]
    cands += [home, os.path.join(home, "HERMES_WORK")]
    out, seen = [], set()
    for c in cands:
        if not c:
            continue
        c = os.path.normpath(c)
        if c.lower() in seen:
            continue
        seen.add(c.lower())
        out.append(c)
    return out


def migrate_configs():
    """Ambil projects.json / exclude.json / manual_entries.json dari lokasi lama."""
    # Mode uji (TASKLOG_DATA_DIR diarahkan): JANGAN bawa apa pun dari lokasi lain,
    # termasuk folder Laporan — kalau tidak, data asli ikut tercampur ke sandbox.
    if os.environ.get("TASKLOG_DATA_DIR"):
        return []
    dd = os.path.normpath(data_dir())
    moved = []
    for name in ("projects.json", "exclude.json", "manual_entries.json"):
        dst = os.path.join(dd, name)
        if os.path.exists(dst):
            continue
        for src_dir in _legacy_cfg_dirs():
            if os.path.normpath(src_dir).lower() == dd.lower():
                continue
            src = os.path.join(src_dir, name)
            if os.path.exists(src):
                try:
                    shutil.copy2(src, dst)
                    moved.append(name)
                    break
                except Exception:
                    pass
    dstL = os.path.join(dd, "Laporan")
    if not os.path.isdir(dstL) or not os.listdir(dstL):
        for src_dir in _legacy_cfg_dirs():
            if os.path.normpath(src_dir).lower() == dd.lower():
                continue
            srcL = os.path.join(src_dir, "Laporan")
            if os.path.isdir(srcL) and os.listdir(srcL):
                try:
                    os.makedirs(dstL, exist_ok=True)
                    for fn in os.listdir(srcL):
                        s, d = os.path.join(srcL, fn), os.path.join(dstL, fn)
                        if os.path.isfile(s) and not os.path.exists(d):
                            shutil.copy2(s, d)
                    moved.append("Laporan")
                    break
                except Exception:
                    pass
    return moved


def ensure_configs(autodetect=True, force_detect=False):
    """Bikin file config kalau belum ada. Return dict info."""
    info = {"created": [], "detected": 0, "moved": migrate_configs()}
    if not os.path.exists(cfg_path()):
        paths = default_roots()
        projs = detect_projects(paths) if autodetect else []
        write_json(cfg_path(), {"_comment": DEFAULT_PROJECTS["_comment"],
                                "_roots_scan": paths, "projects": projs})
        info["created"].append("projects.json")
        info["detected"] = len(projs)
    else:
        cfg = read_json(cfg_path(), DEFAULT_PROJECTS)
        changed = False
        if not cfg.get("_roots_scan"):
            cfg["_roots_scan"] = default_roots()
            changed = True
        else:
            for r in default_roots():           # tambah root baru yang belum terdaftar
                if r not in cfg["_roots_scan"]:
                    cfg["_roots_scan"].append(r)
                    changed = True
        if force_detect or not cfg.get("projects"):
            found = detect_projects(cfg.get("_roots_scan", []))
            have = {p.get("name") for p in cfg.get("projects", [])}
            used_keys = {k.lower() for p in cfg.get("projects", []) for k in p.get("keys", [])}
            for p in found:
                if p["name"] in have:
                    continue
                # jangan tambah kalau kata kuncinya sudah tercakup proyek yang ada
                distinct = [k for k in p["keys"] if k != p["name"].lower()]
                if distinct and all(k in used_keys for k in distinct):
                    continue
                cfg.setdefault("projects", []).append(p)
                used_keys.update(k.lower() for k in p["keys"])
                info["detected"] += 1
                changed = True
        if changed:
            write_json(cfg_path(), cfg)

    if not os.path.exists(exc_path()):
        write_json(exc_path(), DEFAULT_EXCLUDE)
        info["created"].append("exclude.json")
    if not os.path.exists(manual_path()):
        write_json(manual_path(), DEFAULT_MANUAL)
        info["created"].append("manual_entries.json")
    return info

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
tracker.py - Pencatat aktivitas harian (Windows foreground activity logger).

Jalan di background. Tiap POLL detik mencatat window aktif:
  - aplikasi + judul window
  - folder Explorer (path asli via Shell.Application)
  - file Excel/Word (path asli via COM kalau tersedia)
  - topik meeting Zoom/Teams/Meet
Semua masuk SQLite. Tidak ada data yang dikirim keluar.

Pakai:  python tracker.py            (foreground / debug)
        pythonw tracker.py           (background, silent)
        python tracker.py --once     (1x polling, untuk test)
"""
import os
import re
import sys
import json
import time
import ctypes
import sqlite3
import datetime as dt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import tlpaths
except Exception:                                     # fallback kalau file hilang
    tlpaths = None

if tlpaths:
    BASE = tlpaths.data_dir()
    DBDIR = tlpaths.state_dir()
    CONF = tlpaths.cfg_path()
    LOGDIR = tlpaths.log_dir()
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    DBDIR = os.path.join(os.environ.get("LOCALAPPDATA", BASE), "TaskLog")
    CONF = os.path.join(BASE, "projects.json")
    LOGDIR = os.path.join(BASE, "LOG")
DB = os.path.join(DBDIR, "activity.db")
LOG = os.path.join(LOGDIR, "tracker.log")

POLL = 10          # detik antar polling
IDLE_LIMIT = 120   # detik tanpa input -> dianggap idle, tidak dicatat
LOCK_PORT = 51999

# ------------------------------------------------------------------ helpers
def now():
    return dt.datetime.now()

def iso(t):
    return t.strftime("%Y-%m-%d %H:%M:%S")

def logline(msg):
    os.makedirs(LOGDIR, exist_ok=True)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"{iso(now())}  {msg}\n")
    except Exception:
        pass

def idle_seconds():
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if not ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
        return 0
    millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
    return max(0, millis // 1000)

# ------------------------------------------------------------------ config
def load_conf():
    try:
        with open(CONF, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logline(f"CONF ERROR {e}")
        return {"projects": []}

CONF_CACHE = {"t": 0, "v": None}

def conf():
    if time.time() - CONF_CACHE["t"] > 60:
        CONF_CACHE["v"] = load_conf()
        CONF_CACHE["t"] = time.time()
    return CONF_CACHE["v"]


EXC_FILE = os.path.join(BASE, "exclude.json")
_EXC = {"t": 0, "v": None}

def exclusions():
    if _EXC["v"] is None or time.time() - _EXC["t"] > 60:
        try:
            with open(EXC_FILE, "r", encoding="utf-8") as f:
                j = json.load(f)
            _EXC["v"] = ([str(x).lower() for x in j.get("proses", [])],
                         [str(x).lower() for x in j.get("kata", [])])
        except Exception:
            _EXC["v"] = ([], [])
        _EXC["t"] = time.time()
    return _EXC["v"]


def is_excluded(proc, *texts):
    """Privasi: m-banking, password manager, browser tertentu -> tidak dicatat."""
    procs, words = exclusions()
    if (proc or "").lower() in procs:
        return True
    blob = " ".join(t.lower() for t in texts if t)
    if not blob:
        return False
    return any(w in blob for w in words)

def resolve_project(*texts):
    """Cari nama proyek dari kumpulan teks (judul, path, dll)."""
    blob = " | ".join(t.lower() for t in texts if t)
    if not blob.strip():
        return None
    for p in conf().get("projects", []):
        for k in p.get("keys", []):
            if k.lower() in blob:
                return p["name"]
    return None

# ------------------------------------------------------------------ DB
DDL = """
CREATE TABLE IF NOT EXISTS activity (
  id        INTEGER PRIMARY KEY AUTOINCREMENT,
  day       TEXT NOT NULL,
  start_ts  TEXT NOT NULL,
  end_ts    TEXT NOT NULL,
  seconds   INTEGER NOT NULL,
  proc      TEXT,
  title     TEXT,
  context   TEXT,
  kind      TEXT,
  act       TEXT,
  project   TEXT
);
CREATE INDEX IF NOT EXISTS ix_day ON activity(day);
CREATE INDEX IF NOT EXISTS ix_proj ON activity(project);
"""

def db():
    os.makedirs(DBDIR, exist_ok=True)
    c = sqlite3.connect(DB, timeout=15)
    c.executescript(DDL)
    return c

# ------------------------------------------------------------------ window info
def fg_window():
    u = ctypes.windll.user32
    hwnd = u.GetForegroundWindow()
    if not hwnd:
        return None
    n = u.GetWindowTextLengthW(hwnd) + 1
    buf = ctypes.create_unicode_buffer(n)
    u.GetWindowTextW(hwnd, buf, n)
    pid = ctypes.c_ulong()
    u.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return {"hwnd": hwnd, "title": buf.value.strip(), "pid": pid.value}

def proc_info(pid):
    try:
        import psutil
        p = psutil.Process(pid)
        return p.name(), (p.exe() or "")
    except Exception:
        return "", ""

_EXPLORER = {}

def explorer_paths():
    """Peta judul-window -> path folder Explorer (via Shell.Application)."""
    from urllib.parse import unquote
    out = {}
    try:
        import win32com.client
        sh = win32com.client.Dispatch("Shell.Application")
        for w in sh.Windows():
            try:
                u = str(w.LocationURL)
                if u.lower().startswith("file:///"):
                    u = unquote(u[8:]).replace("/", "\\")
                out[str(w.LocationName)] = u
            except Exception:
                pass
    except Exception:
        pass
    return out


# ---- indeks nama folder -> path asli (untuk Explorer & judul window) ----
IDX_FILE = os.path.join(DBDIR, "folder_index.json")


def scan_roots():
    """Folder yang dipindai: dari projects.json (_roots_scan), fallback default."""
    try:
        with open(CONF, "r", encoding="utf-8") as f:
            r = [x for x in json.load(f).get("_roots_scan", []) if x]
        if r:
            return r
    except Exception:
        pass
    if tlpaths:
        return tlpaths.default_roots()
    return [BASE]


_IDX = {"t": 0, "m": {}, "building": False}
MAX_DEPTH = 4


DOC_EXT = (".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pdf", ".dwg", ".pptx", ".csv", ".dxf")

# Excel/Word kadang melaporkan path cloud OneDrive, bukan path lokal.
CARET = {"^j": ", ", "^m": " ", "^i": " ", "^k": " ", "^l": " "}


def decode_caret(s):
    """^J = newline, ^M = CR (notasi caret Excel). Namanya aslinya koma/spasi."""
    out = s
    for k, v in CARET.items():
        out = out.replace(k, v).replace(k.upper(), v)
    return out.strip()


def normalize_path(p):
    """URL OneDrive -> path lokal kalau ada di indeks; kalau tidak, path yang sudah didekode."""
    if not p:
        return p
    low = p.lower()
    if not (low.startswith("http://") or low.startswith("https://")):
        return p
    from urllib.parse import unquote
    body = unquote(p.split("://", 1)[1])
    parts = [x for x in body.split("/") if x]
    if len(parts) > 2 and "docs.live.net" in parts[0].lower():
        parts = parts[2:]
    elif parts:
        parts = parts[1:]
    fn = parts[-1] if parts else ""
    if fn:
        hit = load_index().get(fn.lower())
        if hit:
            return hit                      # path lokal asli dari indeks
    return "\\".join(decode_caret(x) for x in parts)


STALE_MARK = ("#lama", "#kadaluarsa", "#corrupt", "archive", "old", "backup", "temp")


def _score(path):
    low = path.lower()
    s = 0
    for mk in STALE_MARK:
        if mk in low:
            s -= 10
    s -= path.count("\\") * 0.01
    return s


def _put(m, key, path):
    """Simpan path terbaik untuk sebuah nama (hindari folder arsip)."""
    sc = _score(path)
    old = m.get(key)
    if old is None or sc > old[1]:
        m[key] = (path, sc)


def build_index():
    """Indeks nama-folder & nama-file -> path penuh (max depth 4)."""
    m = {}
    for root in scan_roots():
        if not os.path.isdir(root):
            continue
        base_depth = root.rstrip("/").count("/")
        for dirpath, dirnames, filenames in os.walk(root, topdown=True):
            if dirpath.count("/") - base_depth >= MAX_DEPTH:
                dirnames[:] = []
            dirnames[:] = [d for d in dirnames if not d.startswith((".", "~$"))]
            name = os.path.basename(dirpath.rstrip("/"))
            p = dirpath.replace("/", "\\")
            if name:
                _put(m, name.lower(), p)
            for fn in filenames:
                if fn.startswith("~$") or not fn.lower().endswith(DOC_EXT):
                    continue
                _put(m, fn.lower(), os.path.join(dirpath, fn).replace("/", "\\"))
            if len(m) > 150000:
                break
    return {k: v[0] for k, v in m.items()}


def load_index():
    if _IDX["m"] and time.time() - _IDX["t"] < 14400:
        return _IDX["m"]
    m = {}
    fresh = False
    try:
        if time.time() - os.path.getmtime(IDX_FILE) <= 86400:
            with open(IDX_FILE, "r", encoding="utf-8") as f:
                m = json.load(f)
            fresh = True
    except Exception:
        m = {}
    if m:
        _IDX["m"] = m
        _IDX["t"] = time.time()
    if not fresh and not _IDX["building"]:
        start_index_thread()
    return _IDX["m"]


def start_index_thread():
    """Bangun indeks di latar belakang - jangan blokir polling."""
    if _IDX["building"]:
        return
    _IDX["building"] = True

    def work():
        try:
            t0 = time.time()
            m = build_index()
            os.makedirs(DBDIR, exist_ok=True)
            tmp = IDX_FILE + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(m, f)
            os.replace(tmp, IDX_FILE)
            _IDX["m"] = m
            _IDX["t"] = time.time()
            logline(f"index siap: {len(m)} entri dalam {time.time()-t0:.0f}s")
        except Exception as e:
            logline(f"index error {type(e).__name__}: {e}")
        finally:
            _IDX["building"] = False

    import threading
    threading.Thread(target=work, daemon=True).start()


def path_from_title(title):
    """Cari path folder dari judul window Explorer (nama folder di judul)."""
    name = re.sub(r"\s*[-–]\s*(File Explorer|Windows Explorer|Explorer).*$", "", title, flags=re.I).strip()
    if not name:
        return None
    idx = load_index()
    return idx.get(name.lower())

_EXCEL = {"t": 0, "map": {}}

def excel_books():
    """Peta nama-workbook -> full path, dari instance Excel yang hidup."""
    if time.time() - _EXCEL["t"] < 30:
        return _EXCEL["map"]
    m = {}
    try:
        import win32com.client
        app = win32com.client.GetActiveObject("Excel.Application")
        for wb in app.Workbooks:
            try:
                m[str(wb.Name)] = str(wb.FullName)
            except Exception:
                pass
    except Exception:
        pass
    _EXCEL["t"] = time.time()
    _EXCEL["map"] = m
    return m

_WORD = {"t": 0, "map": {}}

def word_docs():
    if time.time() - _WORD["t"] < 30:
        return _WORD["map"]
    m = {}
    try:
        import win32com.client
        app = win32com.client.GetActiveObject("Word.Application")
        for d in app.Documents:
            try:
                m[str(d.Name)] = str(d.FullName)
            except Exception:
                pass
    except Exception:
        pass
    _WORD["t"] = time.time()
    _WORD["map"] = m
    return m

# ------------------------------------------------------------------ klasifikasi
MEET_RE = re.compile(r"(zoom|teams|meet\.google|webex|google meet)", re.I)

RULES = [
    ("Meeting",       r"zoom meeting|zoom webinar|microsoft teams|google meet|meet - |webex|gmeet|\bmeeting\b|\brapat\b|conference"),
    ("Penyusunan / Pengecekan BQ-RAB", r"\bbq\b|\bboq\b|\bbow\b|\brab\b|\bhps\b|\bmc-?0\b|\bmc-?100\b|\bdkh\b|\bahsp\b|\bhsd\b|analisa harga|estimate|takeoff"),
    ("Cek Variation Order", r"\bvo\b|variation|change order|addendum|\bco\b|\bklaim\b|\bclaim\b"),
    ("Cek Progress / Kurva S", r"kurva\s?s|\bprogress\b|\bschedule\b|\btimeline\b|termin 20|deviasi"),
    ("Review Gambar / Desain", r"\bgambar\b|\bdwg\b|\bcad\b|\blayout\b|\bdenah\b|\bdetail\b|sketchup|autocad|gstarcad|planswift|\bifc\b"),
    ("Draft Surat / Korespondensi", r"\bsurat\b|\bletter\b|\bdraft\b|\bmemo\b|nota dinas|\bemail\b|\btender\b|\bpenawaran\b|\bproposal\b|korespondensi"),
    ("Cek Termin / BAP", r"\btermin\b|\bbap\b|\binvoice\b|opname|berita acara|pembayaran"),
    ("Review / Cek Dokumen", r"\breview\b|\bcek\b|\bcheck\b|revisi|koreksi|validasi|\baudit\b|forensik|\bbanding\b"),
    ("Susun Laporan", r"\blaporan\b|\breport\b|weekly|bulanan|dashboard"),
]

# aplikasi -> label kalau tidak ada aturan yang cocok
APP_ACT = {
    "explorer.exe": "Buka / Telusuri Dokumen Proyek",
    "excel.exe": "Olah Data Excel",
    "winword.exe": "Olah Dokumen Word",
    "powerpnt.exe": "Olah Presentasi",
    "msedge.exe": "Browsing / Riset",
    "chrome.exe": "Browsing / Riset",
    "firefox.exe": "Browsing / Riset",
    "brave.exe": "Browsing / Riset",
    "outlook.exe": "Email / Korespondensi",
    "thunderbird.exe": "Email / Korespondensi",
    "acrobat.exe": "Baca Dokumen PDF",
    "acroread.exe": "Baca Dokumen PDF",
    "sumatrapdf.exe": "Baca Dokumen PDF",
    "pdfxedit.exe": "Baca / Edit PDF",
    "notepad.exe": "Catatan / Teks",
    "notepad++.exe": "Catatan / Teks",
    "code.exe": "Kerja Kode / Script",
    "cmd.exe": "Terminal / Script",
    "powershell.exe": "Terminal / Script",
    "windowsterminal.exe": "Terminal / Script",
    "gstarcad.exe": "Review Gambar / Desain",
    "acad.exe": "Review Gambar / Desain",
    "calculator.exe": "Kalkulasi",
    "hermes.exe": "Kerja dengan AI Agent",
    "winscp.exe": "Transfer File",
    "7zfm.exe": "Arsip File",
    "winrar.exe": "Arsip File",
}


def classify(title, context, proc):
    blob = f"{title} {context} {proc}"
    if MEET_RE.search(proc) or MEET_RE.search(blob):
        return "Meeting"
    for label, rx in RULES:
        if re.search(rx, blob, re.I):
            return label
    return APP_ACT.get((proc or "").lower(), "Aktivitas Aplikasi")

def meeting_topic(title, proc):
    """Ambil topik meeting dari judul window."""
    t = title
    t = re.sub(r"\s*[-–|]\s*Zoom (Meeting|Webinar).*$", "", t, flags=re.I)
    t = re.sub(r"\s*[-–|]\s*(Microsoft Teams|Google Meet|Webex).*$", "", t, flags=re.I)
    for junk in ["Zoom Meeting", "Zoom Webinar", "Meeting", "Microsoft Teams", "Google Meet"]:
        if t.strip().lower() == junk.lower():
            return t.strip()
    return t.strip() or "Meeting"

def file_name(title):
    """Judul window Excel/Word -> nama file (tanpa ' - Excel')."""
    t = re.sub(r"\s*[-–]\s*(Excel|Word|PowerPoint|Protected View|Read-Only|Compatibility Mode).*$", "", title, flags=re.I)
    return t.strip()

def detailed_context(wi, proc, title):
    """Tebak konteks sebenarnya: folder / file / url."""
    pn = (proc or "").lower()
    name = file_name(title)
    if pn == "explorer.exe":
        paths = explorer_paths()
        for loc_name, loc_path in paths.items():
            if loc_name and loc_name in title:
                return loc_path
        p = path_from_title(title)
        return p or title
    if pn == "excel.exe":
        books = excel_books()
        for bname, bpath in books.items():
            if bname and bname in name:
                return normalize_path(bpath)
        return path_from_title(name) or name
    if pn == "winword.exe":
        docs = word_docs()
        for dname, dpath in docs.items():
            if dname and dname in name:
                return normalize_path(dpath)
        return path_from_title(name) or name
    if name and name.lower().endswith(DOC_EXT):
        return path_from_title(name) or name
    return name

# ------------------------------------------------------------------ main loop
STATUS = os.path.join(DBDIR, "tracker_status.json")

def write_status(cur):
    try:
        os.makedirs(DBDIR, exist_ok=True)
        with open(STATUS, "w", encoding="utf-8") as f:
            json.dump({"pid": os.getpid(), "last": iso(now()),
                       "project": cur.get("project"), "act": cur.get("act"),
                       "title": cur.get("title", "")[:120]}, f)
    except Exception:
        pass

def run_once(cur, conn, sample=None):
    """1x polling. Return state 'cur' terbaru."""
    t = now()
    if idle_seconds() >= IDLE_LIMIT:
        if cur:
            flush(cur, conn, t)
        write_status({})
        return None

    wi = sample or fg_window()
    if not wi or not wi["title"]:
        return cur

    proc, exe = proc_info(wi["pid"])
    if proc.lower() in ("lockapp.exe", "logonui.exe", "searchhost.exe", "shellexperiencehost.exe", "textinputhost.exe"):
        return cur

    title = wi["title"]
    ctx = detailed_context(wi, proc, title)
    if is_excluded(proc, title, ctx):
        if cur:
            flush(cur, conn, t)
        write_status({})
        return None
    proj = resolve_project(title, ctx, exe, proc)
    kind = classify(title, ctx, proc)
    if kind == "Meeting":
        act = meeting_topic(title, proc)
        ctx = act
    else:
        act = kind
    if not proj and cur and cur.get("project") and kind == "Meeting":
        # stickiness: meeting tanpa nama proyek -> ikut proyek terakhir (maks 15 menit)
        if (t - cur["end"]).total_seconds() <= 900:
            proj = cur["project"]

    key = (proc, title)
    if cur and (cur["proc"], cur["title"]) == key:
        cur["end"] = t
        cur["seconds"] = int((t - cur["start"]).total_seconds())
        if not cur["project"] and proj:
            cur["project"] = proj
        if not cur["context"] or len(ctx) > len(cur["context"]):
            cur["context"] = ctx
        write_status(cur)
        return cur

    if cur:
        flush(cur, conn, t)

    cur = {"proc": proc, "title": title, "context": ctx, "kind": kind, "act": act,
           "project": proj, "start": t, "end": t, "seconds": 0}
    write_status(cur)
    return cur

def flush(cur, conn, t):
    if not cur:
        return
    secs = int((t - cur["start"]).total_seconds())
    if secs < 5:
        return
    conn.execute(
        "INSERT INTO activity (day,start_ts,end_ts,seconds,proc,title,context,kind,act,project) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (cur["start"].strftime("%Y-%m-%d"), iso(cur["start"]), iso(t), secs,
         cur["proc"], cur["title"], cur["context"], cur["kind"], cur["act"], cur["project"]))
    conn.commit()

_LOCK_SOCK = None

def already_running():
    """Cegah dobel instance - kunci via socket loopback (anti race)."""
    global _LOCK_SOCK
    import socket
    import threading
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", LOCK_PORT))
        s.listen(64)
        _LOCK_SOCK = s

        def accepter():
            """Terima & tutup koneksi probe, supaya backlog tidak penuh
            (kalau tidak, probe berikutnya timeout dan tracker dianggap mati)."""
            while True:
                try:
                    c, _ = s.accept()
                    try:
                        c.close()
                    except Exception:
                        pass
                except Exception:
                    return

        threading.Thread(target=accepter, daemon=True).start()
        return False
    except OSError:
        try:
            s.close()
        except Exception:
            pass
        return True


def is_alive(timeout=1.0):
    """Cek tracker jalan atau tidak (dari proses lain)."""
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(timeout)
    try:
        s.connect(("127.0.0.1", LOCK_PORT))
        return True
    except Exception:
        return False
    finally:
        try:
            s.close()
        except Exception:
            pass


def main():
    os.makedirs(DBDIR, exist_ok=True)
    if tlpaths:
        try:
            tlpaths.ensure_configs()
        except Exception:
            pass
    if "--once" not in sys.argv and already_running():
        logline("SKIP - sudah jalan")
        return
    conn = db()
    logline(f"START poll={POLL}s db={DB}")
    cur = None
    only_once = "--once" in sys.argv
    try:
        while True:
            try:
                cur = run_once(cur, conn)
            except Exception as e:
                logline(f"ERR {type(e).__name__}: {e}")
            if only_once:
                break
            time.sleep(POLL)
    finally:
        if cur:
            flush(cur, conn, now())
        logline("STOP")
        conn.close()

if __name__ == "__main__":
    main()

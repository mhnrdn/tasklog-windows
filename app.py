#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
app.py - TaskLog: pencatat tugas harian. Satu file, satu klik.

Mode:
  TaskLog.exe                      -> buka aplikasi (GUI)
  TaskLog.exe --tracker            -> jalankan pencatat di background (silent)
  TaskLog.exe --report [--day D]   -> tulis laporan lalu keluar
  TaskLog.exe --report --audit     -> tulis laporan + audit
  TaskLog.exe --setup              -> siapkan config + autostart, tanpa GUI
  TaskLog.exe --autostart on|off   -> pasang/copot autostart
"""
import os
import sys
import json
import time
import socket
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import tlpaths                                                  # noqa: E402
import report as R                                              # noqa: E402

LOCK_PORT = 51999
APP_NAME = "TaskLog"
NO_WINDOW = 0x08000000          # CREATE_NO_WINDOW
DETACHED = 0x00000008           # DETACHED_PROCESS


# ------------------------------------------------------------------ proses
def self_cmd(*args):
    """Perintah untuk menjalankan diri sendiri."""
    if getattr(sys, "frozen", False):
        return [sys.executable] + list(args)
    py = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = py if os.path.exists(py) else sys.executable
    return [exe, os.path.abspath(__file__)] + list(args)


def tracker_alive(timeout=1.0):
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


def start_tracker():
    if tracker_alive():
        return "sudah jalan"
    try:
        # stdio ke DEVNULL: kalau tidak, proses anak menahan pipe pemanggil
        # dan shell yang menjalankan TaskLog.exe tidak pernah selesai.
        subprocess.Popen(self_cmd("--tracker"),
                         stdin=subprocess.DEVNULL,
                         stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL,
                         creationflags=NO_WINDOW | DETACHED,
                         close_fds=True)
    except Exception as e:
        return f"gagal: {e}"
    for _ in range(25):
        time.sleep(0.4)
        if tracker_alive():
            return "jalan"
    return "dijalankan (menunggu)"


def stop_tracker():
    killed = 0
    try:
        import psutil
        me = os.getpid()
        for p in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                if p.info["pid"] == me:
                    continue
                cl = p.info["cmdline"] or []
                if "--tracker" in cl:
                    p.kill()
                    killed += 1
            except Exception:
                pass
    except Exception:
        pass
    return f"{killed} proses dihentikan"


# ------------------------------------------------------------------ autostart
def startup_dir():
    return os.path.join(os.environ.get("APPDATA", ""),
                        "Microsoft", "Windows", "Start Menu", "Programs", "Startup")


def autostart_path():
    return os.path.join(startup_dir(), f"{APP_NAME}.lnk")


def autostart_on():
    import win32com.client
    tgt = self_cmd("--tracker")
    sh = win32com.client.Dispatch("WScript.Shell")
    lnk = sh.CreateShortcut(autostart_path())
    lnk.TargetPath = tgt[0]
    lnk.Arguments = " ".join(f'"{a}"' if " " in a else a for a in tgt[1:])
    lnk.WorkingDirectory = tlpaths.app_dir()
    lnk.Description = "TaskLog - pencatat tugas harian (otomatis saat login)"
    lnk.Save()
    return autostart_path()


def autostart_off():
    p = autostart_path()
    if os.path.exists(p):
        os.remove(p)
        return True
    return False


def autostart_active():
    return os.path.exists(autostart_path())


# ------------------------------------------------------------------ laporan
def do_report(day=None, month=None, audit=False):
    day = day or __import__("datetime").date.today().isoformat()
    os.makedirs(R.OUT, exist_ok=True)
    written = []
    if month:
        txt = R.month_md(month)
        p = os.path.join(R.OUT, f"Laporan_Bulanan_{month}.md")
        open(p, "w", encoding="utf-8").write(txt)
        written.append(p)
    else:
        p1 = os.path.join(R.OUT, f"Laporan_Harian_{day}.md")
        open(p1, "w", encoding="utf-8").write(R.day_md(day) + "\n")
        p2 = os.path.join(R.OUT, f"Ringkas_{day}.txt")
        open(p2, "w", encoding="utf-8").write(R.day_md(day, for_calendar=True) + "\n")
        written += [p1, p2]
        p3 = os.path.join(R.OUT, f"Audit_{day}.txt")
        open(p3, "w", encoding="utf-8").write(R.audit_txt(day) + "\n")
        written.append(p3)
    return written


def open_path(p):
    try:
        if os.path.isfile(p):
            os.startfile(p)                       # noqa: S606
        else:
            subprocess.Popen(["explorer", os.path.normpath(p)])
    except Exception:
        pass


# ------------------------------------------------------------------ GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    tlpaths.ensure_configs()
    root = tk.Tk()
    root.title(f"{APP_NAME}  —  Pencatat Tugas Harian")
    root.geometry("880x660")
    root.minsize(760, 560)

    style = ttk.Style()
    try:
        style.theme_use("vista")
    except Exception:
        pass

    # ---------- header
    head = tk.Frame(root, bg="#12324f", height=64)
    head.pack(fill="x")
    tk.Label(head, text=f"{APP_NAME}", bg="#12324f", fg="white",
             font=("Segoe UI", 16, "bold")).pack(side="left", padx=14, pady=12)
    status = tk.Label(head, text="", bg="#12324f", fg="#9fe8b0",
                      font=("Segoe UI", 10, "bold"))
    status.pack(side="right", padx=14)

    # ---------- tombol
    bar = tk.Frame(root, padx=10, pady=8)
    bar.pack(fill="x")

    def mk(parent, text, cmd, col, row, width=22):
        b = tk.Button(parent, text=text, command=cmd, width=width, height=1,
                      font=("Segoe UI", 9), cursor="hand2")
        b.grid(row=row, column=col, padx=4, pady=3, sticky="ew")
        return b

    for c in range(3):
        bar.columnconfigure(c, weight=1)

    # ---------- notebook
    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=10, pady=(0, 6))

    def text_tab(title):
        f = tk.Frame(nb)
        nb.add(f, text=title)
        t = tk.Text(f, wrap="none", font=("Consolas", 9), undo=False)
        ys = ttk.Scrollbar(f, orient="vertical", command=t.yview)
        xs = ttk.Scrollbar(f, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        t.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)
        return t

    txt_rep = text_tab("Laporan Hari Ini")
    txt_aud = text_tab("Audit / Periksa")
    txt_inf = text_tab("Info & Bantuan")

    def set_text(w, s):
        w.configure(state="normal")
        w.delete("1.0", "end")
        w.insert("1.0", s)
        w.configure(state="disabled")

    # ---------- status bar
    sbar = tk.Label(root, text="", anchor="w", font=("Segoe UI", 8), fg="#555")
    sbar.pack(fill="x", padx=12, pady=(0, 6))

    def refresh():
        alive = tracker_alive()
        status.config(text=("● PENCATAT AKTIF" if alive else "● PENCATAT MATI"),
                      fg=("#8ef0a4" if alive else "#ff9b9b"))
        as_ = "aktif" if autostart_active() else "tidak aktif"
        sbar.config(text=f"Data: {tlpaths.data_dir()}   |   Autostart: {as_}")
        try:
            set_text(txt_rep, R.day_md(__import__("datetime").date.today().isoformat()))
        except Exception as e:
            set_text(txt_rep, f"gagal baca laporan: {e}")
        root.after(20000, refresh)

    def refresh_audit():
        day = __import__("datetime").date.today().isoformat()
        try:
            set_text(txt_aud, R.audit_txt(day))
        except Exception as e:
            set_text(txt_aud, f"gagal: {e}")

    def act_report():
        written = do_report()
        refresh()
        refresh_audit()
        messagebox.showinfo(APP_NAME, "Laporan sudah diperbarui:\n" + "\n".join(written))

    def act_start():
        messagebox.showinfo(APP_NAME, "Pencatat: " + start_tracker())
        refresh()

    def act_stop():
        messagebox.showinfo(APP_NAME, stop_tracker())
        refresh()

    def act_toggle_autostart():
        if autostart_active():
            autostart_off()
            messagebox.showinfo(APP_NAME, "Autostart dimatikan.\nPencatat tidak nyala sendiri saat login.")
        else:
            p = autostart_on()
            messagebox.showinfo(APP_NAME, "Autostart dipasang:\n" + p)
        refresh()

    def edit_file(path, ensure_default=None):
        tlpaths.ensure_configs()
        if ensure_default and not os.path.exists(path):
            tlpaths.write_json(path, ensure_default)
        if not os.path.exists(path):
            tlpaths.write_json(path, {})
        subprocess.Popen(["notepad.exe", path])

    def act_projs():
        edit_file(tlpaths.cfg_path(), tlpaths.DEFAULT_PROJECTS)

    def act_manual():
        edit_file(tlpaths.manual_path(), tlpaths.DEFAULT_MANUAL)

    def act_excl():
        edit_file(tlpaths.exc_path(), tlpaths.DEFAULT_EXCLUDE)

    def act_detect():
        cfg = tlpaths.read_json(tlpaths.cfg_path(), tlpaths.DEFAULT_PROJECTS)
        found = tlpaths.detect_projects(cfg.get("_roots_scan") or tlpaths.default_roots())
        if not found:
            messagebox.showwarning(APP_NAME, "Tidak ada folder proyek yang terdeteksi.\n"
                                             "Tambahkan folder lewat tombol 'Folder yang Dipantau'.")
            return
        win = tk.Toplevel(root)
        win.title("Hasil deteksi proyek")
        win.geometry("720x520")
        tk.Label(win, text=f"Terdeteksi {len(found)} folder proyek. "
                           f"Boleh diedit dulu (hapus baris yang bukan proyek), "
                           f"lalu klik Simpan.",
                 anchor="w").pack(fill="x", padx=10, pady=8)
        t = tk.Text(win, wrap="none", font=("Consolas", 9))
        t.pack(fill="both", expand=True, padx=10)
        t.insert("1.0", json.dumps(found, indent=2, ensure_ascii=False))

        def save():
            # ambil dari kotak teks supaya bisa diedit dulu
            try:
                items = json.loads(t.get("1.0", "end"))
                if not isinstance(items, list):
                    raise ValueError("bukan list")
            except Exception as e:
                messagebox.showerror(APP_NAME, f"JSON tidak valid: {e}\n\nPerbaiki dulu.")
                return
            have = {p.get("name") for p in cfg.get("projects", [])}
            n = 0
            for p in items:
                if isinstance(p, dict) and p.get("name") and p["name"] not in have:
                    cfg.setdefault("projects", []).append(p)
                    n += 1
            tlpaths.write_json(tlpaths.cfg_path(), cfg)
            win.destroy()
            messagebox.showinfo(APP_NAME, f"{n} proyek ditambahkan ke projects.json\n\n"
                                          f"Daftar bisa kamu rapikan lagi lewat tombol\n"
                                          f"'Daftar Proyek (projects.json)'.")

        tk.Button(win, text="Simpan ke projects.json", command=save,
                  height=2, bg="#dff0d8").pack(fill="x", padx=10, pady=10)

    def act_roots():
        cfg = tlpaths.read_json(tlpaths.cfg_path(), tlpaths.DEFAULT_PROJECTS)
        roots = list(cfg.get("_roots_scan") or tlpaths.default_roots())
        win = tk.Toplevel(root)
        win.title("Folder yang dipantau")
        win.geometry("660x420")
        tk.Label(win, text="TaskLog memindai folder ini untuk mengenali nama proyek.",
                 anchor="w").pack(fill="x", padx=10, pady=8)
        lb = tk.Listbox(win, font=("Consolas", 9))
        lb.pack(fill="both", expand=True, padx=10)
        for r in roots:
            lb.insert("end", r)

        def add():
            d = filedialog.askdirectory(title="Pilih folder proyek")
            if d:
                d = os.path.normpath(d)
                if d not in roots:
                    roots.append(d)
                    lb.insert("end", d)

        def rem():
            for i in reversed(lb.curselection()):
                lb.delete(i)
                roots.pop(i)

        def save():
            cfg["_roots_scan"] = roots
            tlpaths.write_json(tlpaths.cfg_path(), cfg)
            win.destroy()
            messagebox.showinfo(APP_NAME, "Folder tersimpan.\nJalankan 'Cari Proyek' untuk memindai.")

        f = tk.Frame(win)
        f.pack(fill="x", padx=10, pady=8)
        tk.Button(f, text="Tambah folder", command=add, width=16).pack(side="left", padx=3)
        tk.Button(f, text="Hapus", command=rem, width=10).pack(side="left", padx=3)
        tk.Button(f, text="Simpan", command=save, width=12, bg="#dff0d8").pack(side="right", padx=3)

    def act_month():
        month = __import__("datetime").date.today().strftime("%Y-%m")
        do_report(month=month)
        open_path(os.path.join(R.OUT, f"Laporan_Bulanan_{month}.md"))

    def act_setup():
        info = tlpaths.ensure_configs(force_detect=True)
        p = autostart_on() if not autostart_active() else autostart_path()
        msg = start_tracker()
        msg2 = (f"Siap.\n\nConfig dibuat: {', '.join(info['created']) or '-'}\n"
                f"Proyek baru terdeteksi: {info['detected']}\n"
                f"Autostart: {p}\nPencatat: {msg}\n\n"
                f"Data & laporan ada di:\n{tlpaths.data_dir()}")
        messagebox.showinfo(APP_NAME, msg2)
        refresh()

    # baris tombol
    mk(bar, "📄  Buat Laporan Sekarang", act_report, 0, 0)
    mk(bar, "📂  Buka Folder Laporan", lambda: open_path(R.OUT), 1, 0)
    mk(bar, "📅  Laporan Bulanan", act_month, 2, 0)
    mk(bar, "▶  Nyalakan Pencatat", act_start, 0, 1)
    mk(bar, "■  Matikan Pencatat", act_stop, 1, 1)
    mk(bar, "🔁  Autostart: Pasang/Copot", act_toggle_autostart, 2, 1)
    mk(bar, "✏  Daftar Proyek (projects.json)", act_projs, 0, 2)
    mk(bar, "📝  Catatan Manual", act_manual, 1, 2)
    mk(bar, "🚫  Pengecualian (exclude.json)", act_excl, 2, 2)
    mk(bar, "🔍  Cari Proyek Otomatis", act_detect, 0, 3)
    mk(bar, "🗂  Folder yang Dipantau", act_roots, 1, 3)
    mk(bar, "⚙  Siapkan Semua (setup awal)", act_setup, 2, 3)

    info_txt = f"""TaskLog — pencatat tugas harian

CARA PAKAI
  1. Klik "Siapkan Semua" sekali. Selesai.
  2. Kerjakan pekerjaan seperti biasa. TaskLog mencatat window aktif
     tiap 10 detik: aplikasi, folder Explorer, file Excel/Word,
     topik meeting Zoom/Teams/Meet.
  3. Klik "Buat Laporan Sekarang" kapan saja, atau biarkan laporan
     harian otomatis siap tiap sore.

SEMUA FILE BISA KAMU EDIT
  projects.json          daftar proyek + kata kunci
  exclude.json           yang TIDAK dicatat (m-banking, password manager)
  manual_entries.json    pekerjaan di luar laptop (kunjungan, telepon)

LOKASI
  Data & laporan : {tlpaths.data_dir()}
  Database       : {tlpaths.state_dir()}\\activity.db
  Log            : {tlpaths.log_dir()}\\tracker.log

LAPORAN
  Laporan\\Laporan_Harian_<tanggal>.md   laporan lengkap
  Laporan\\Ringkas_<tanggal>.txt         siap tempel ke Google Calendar
  Laporan\\Audit_<tanggal>.txt           telusuri sumber tiap baris
  Laporan\\Laporan_Bulanan_<bulan>.md    rekap 1 bulan

PRIVASI
  Data tersimpan di laptop sendiri. Tidak ada yang dikirim ke internet.
  Aplikasi di exclude.json tidak ditulis ke database sama sekali.

TIDAK MUNCUL / SALAH PROYEK?
  Tab "Audit / Periksa" menjawab asal tiap baris laporan.
"""
    set_text(txt_inf, info_txt)

    refresh()
    refresh_audit()
    root.mainloop()


# ------------------------------------------------------------------ CLI
def say(*parts):
    """Tulis ke layar DAN ke LOG\\app.log (exe noconsole tidak punya stdout)."""
    msg = " ".join(str(p) for p in parts)
    try:
        with open(os.path.join(tlpaths.log_dir(), "app.log"), "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}  {msg}\n")
    except Exception:
        pass
    try:
        print(msg)
    except Exception:
        pass


def main():
    a = sys.argv[1:]
    say("=== TaskLog start args=" + repr(a) + " frozen=" + str(getattr(sys, "frozen", False)))
    if "--tracker" in a:
        import tracker
        tracker.main()
        return
    if "--report" in a:
        day = a[a.index("--day") + 1] if "--day" in a else None
        month = a[a.index("--month") + 1] if "--month" in a else None
        for p in do_report(day=day, month=month, audit=True):
            say("laporan:", p)
        return
    if "--autostart" in a:
        i = a.index("--autostart")
        mode = a[i + 1] if len(a) > i + 1 else "on"
        if str(mode).lower() in ("off", "0", "no"):
            say("autostart dicopot" if autostart_off() else "autostart memang tidak ada")
        else:
            say("autostart dipasang:", autostart_on())
        return
    if "--setup" in a:
        info = tlpaths.ensure_configs(force_detect=True)
        say("config dibuat:", info["created"])
        say("config dipindah dari lokasi lama:", info["moved"])
        say("proyek terdeteksi:", info["detected"])
        say("autostart:", autostart_on())
        say("pencatat:", start_tracker())
        say("data:", tlpaths.data_dir())
        say("state:", tlpaths.state_dir())
        return
    if "--status" in a:
        say("pencatat:", "AKTIF" if tracker_alive() else "MATI")
        say("autostart:", "AKTIF" if autostart_active() else "MATI")
        say("data:", tlpaths.data_dir())
        say("state:", tlpaths.state_dir())
        say("config:", tlpaths.cfg_path(), "ada" if os.path.exists(tlpaths.cfg_path()) else "TIDAK ADA")
        return
    try:
        run_gui()
    except Exception:
        import traceback
        say("=== ERROR GUI ===")
        for line in traceback.format_exc().splitlines():
            say("  " + line)
        try:
            import tkinter.messagebox as mb
            import tkinter as tk
            r = tk.Tk()
            r.withdraw()
            mb.showerror("TaskLog - error",
                         "Aplikasi gagal dibuka.\n\nDetail tersimpan di:\n"
                         + os.path.join(tlpaths.log_dir(), "app.log"))
        except Exception:
            pass


if __name__ == "__main__":
    main()

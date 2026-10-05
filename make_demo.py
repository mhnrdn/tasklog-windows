#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Bikin data contoh + jalankan GUI untuk keperluan screenshot dokumentasi."""
import os
import sys
import json
import sqlite3
import datetime as dt

DEMO = os.path.join(os.environ["LOCALAPPDATA"], "hermes", "cache", "scratch", "demo_gui")
DATA = os.path.join(DEMO, "TaskLog_data")
STATE = os.path.join(DEMO, "state")

PROJECTS = {
    "_comment": "Peta kata kunci -> nama proyek. Urutan = prioritas.",
    "_roots_scan": [r"C:\Users\Public\Documents"],
    "projects": [
        {"name": "Proyek Gedung Kantor", "keys": ["gedung kantor", "kantor pusat", "gk-01"]},
        {"name": "Proyek Sekolah", "keys": ["sekolah", "sd-02"]},
        {"name": "Proyek Jembatan", "keys": ["jembatan", "br-05"]},
        {"name": "Internal - Standar Harga", "keys": ["ahsp", "standar harga", "hsp"]},
    ],
}

ROWS = [
    ("08:10", "10:05", "EXCEL.EXE", r"C:\Users\Public\Documents\GK-01 Gedung Kantor\RAB\RAB Gedung Kantor Rev2.xlsx",
     "Penyusunan / Pengecekan BQ-RAB", "Proyek Gedung Kantor"),
    ("10:05", "10:45", "EXCEL.EXE", r"C:\Users\Public\Documents\GK-01 Gedung Kantor\BQ\BOQ Final.xlsx",
     "Penyusunan / Pengecekan BQ-RAB", "Proyek Gedung Kantor"),
    ("10:45", "11:30", "Zoom.exe", "Weekly Meeting Proyek Gedung Kantor - Zoom Meeting",
     "Meeting", "Proyek Gedung Kantor"),
    ("13:00", "14:20", "EXCEL.EXE", r"C:\Users\Public\Documents\SD-02 Sekolah\VO-03 Volume Tambahan.xlsx",
     "Cek Variation Order", "Proyek Sekolah"),
    ("14:20", "15:35", "EXCEL.EXE", r"C:\Users\Public\Documents\AHSP 2026\AHSP 2026 koefisien.xlsx",
     "Penyusunan / Pengecekan BQ-RAB", "Internal - Standar Harga"),
    ("15:35", "16:00", "msedge.exe", "Harga material besi beton 2026 - Personal - Microsoft Edge",
     "Browsing / Riset", None),
]

MANUAL = {
    "_comment": "Pekerjaan di luar laptop.",
    dt.date.today().isoformat(): [
        {"project": "Proyek Jembatan", "act": "Kunjungan Lapangan / Survey",
         "minutes": 90, "start": "16:10", "end": "17:40", "note": "Survey lokasi jembatan"}
    ],
}


def main():
    os.makedirs(os.path.join(DATA, "Laporan"), exist_ok=True)
    os.makedirs(os.path.join(DATA, "LOG"), exist_ok=True)
    os.makedirs(STATE, exist_ok=True)

    for name, obj in (("projects.json", PROJECTS), ("manual_entries.json", MANUAL)):
        with open(os.path.join(DATA, name), "w", encoding="utf-8") as f:
            json.dump(obj, f, indent=2, ensure_ascii=False)

    db = os.path.join(STATE, "activity.db")
    if os.path.exists(db):
        os.remove(db)
    c = sqlite3.connect(db)
    c.executescript("""
    CREATE TABLE activity (
      id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, start_ts TEXT, end_ts TEXT,
      seconds INTEGER, proc TEXT, title TEXT, context TEXT, kind TEXT, act TEXT, project TEXT);
    """)
    day = dt.date.today().isoformat()
    for s, e, proc, path, act, proj in ROWS:
        secs = (dt.datetime.strptime(e, "%H:%M") - dt.datetime.strptime(s, "%H:%M")).seconds
        c.execute("INSERT INTO activity (day,start_ts,end_ts,seconds,proc,title,context,kind,act,project)"
                  " VALUES (?,?,?,?,?,?,?,?,?,?)",
                  (day, f"{day} {s}:00", f"{day} {e}:00", secs, proc,
                   os.path.basename(path), path, act, act, proj))
    c.commit()
    c.close()
    print("demo siap:", DEMO)


if __name__ == "__main__":
    main()

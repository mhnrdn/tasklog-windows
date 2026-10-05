#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Ambil screenshot jendela TaskLog untuk dokumentasi (atomik, sekali jalan)."""
import os
import sys
import time
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)     # DPI aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import win32gui                                             # noqa: E402
import win32con                                             # noqa: E402
import win32com.client                                      # noqa: E402
from PIL import ImageGrab                                   # noqa: E402


def find_window():
    hits = []
    win32gui.EnumWindows(
        lambda h, _: hits.append(h) if win32gui.GetWindowText(h).startswith("TaskLog") else None, None)
    return hits[0] if hits else None


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "screenshot.png"
    h = find_window()
    if not h:
        print("JENDELA TIDAK DITEMUKAN")
        return 1
    try:
        win32com.client.Dispatch("Shell.Application").MinimizeAll()
        time.sleep(2)
    except Exception:
        pass
    h = find_window() or h
    win32gui.ShowWindow(h, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(h)
    except Exception:
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
        ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
        win32gui.SetForegroundWindow(h)
    win32gui.ShowWindow(h, win32con.SW_MAXIMIZE)
    time.sleep(2.5)

    l, t, r, b = win32gui.GetWindowRect(h)
    print("rect:", l, t, r, b)
    img = ImageGrab.grab(bbox=(max(0, l), max(0, t), r, b), all_screens=True)
    img.thumbnail((1400, 1400))
    img.save(out, optimize=True)
    print("tersimpan:", out, img.size, round(os.path.getsize(out) / 1024, 1), "KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())

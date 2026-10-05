#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""kill_tracker.py - hentikan semua proses TaskLog/tracker (aman, tidak bunuh diri).

Pakai: python kill_tracker.py
"""
import os
import sys
import time

MINE = os.getpid()


def main():
    try:
        import psutil
    except ImportError:
        print("psutil tidak ada")
        return 1
    killed = []
    for p in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            pid = p.info["pid"]
            if pid == MINE or pid == os.getppid():
                continue
            cl = p.info["cmdline"] or []
            name = (p.info["name"] or "").lower()
            is_tracker = "--tracker" in cl
            is_script = any(str(x).lower().endswith("tracker.py") for x in cl)
            is_exe = name.startswith("tasklog")
            if is_tracker or is_script or is_exe:
                p.kill()
                killed.append((pid, p.info["name"]))
        except Exception:
            pass
    time.sleep(1)
    if killed:
        for pid, nm in killed:
            print(f"dihentikan: {pid} {nm}")
    else:
        print("tidak ada proses TaskLog/tracker yang jalan")
    return 0


if __name__ == "__main__":
    sys.exit(main())

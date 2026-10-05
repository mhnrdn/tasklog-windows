#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Bikin icon.ico (32x32 + 16x16) pakai modul bawaan saja."""
import os
import struct


def img(size):
    """Gambar jam + centang sederhana, BGRA."""
    px = []
    c = size / 2.0
    r_out = size * 0.46
    r_in = size * 0.36
    for y in range(size):
        row = []
        for x in range(size):
            dx, dy = x + 0.5 - c, y + 0.5 - c
            d = (dx * dx + dy * dy) ** 0.5
            if d <= r_in:
                row.append((255, 255, 255, 255))          # BGRA putih
            elif d <= r_out:
                row.append((200, 120, 30, 255))           # biru tua
            else:
                row.append((0, 0, 0, 0))                  # transparan
        px.append(row)
    # jarum jam
    for i in range(int(size * 0.22)):
        y = int(c) - i
        if 0 <= y < size:
            px[y][int(c)] = (200, 120, 30, 255)
    for i in range(int(size * 0.18)):
        x = int(c) + i
        if 0 <= x < size:
            px[int(c)][x] = (200, 120, 30, 255)
    return px


def bmp(px, size):
    """DIB header + pixel data, bottom-up, BGRA."""
    hdr = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0,
                      size * size * 4, 0, 0, 0, 0)
    body = b""
    for y in range(size - 1, -1, -1):
        for x in range(size):
            b, g, r, a = px[y][x]
            body += bytes((b, g, r, a))
    return hdr + body


def build(path):
    sizes = [16, 32, 48]
    imgs = [(s, bmp(img(s), s)) for s in sizes]
    out = struct.pack("<HHH", 0, 1, len(imgs))
    offset = 6 + 16 * len(imgs)
    for s, data in imgs:
        out += struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
    for _, data in imgs:
        out += data
    with open(path, "wb") as f:
        f.write(out)
    return path


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    p = build(os.path.join(here, "icon.ico"))
    print(p, os.path.getsize(p), "bytes")

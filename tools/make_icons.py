#!/usr/bin/env python3
"""Ve icon PNG 64x64 (thuan Python) cho Ribbon: customUI/images/pagenum.png"""
import struct, zlib
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
N = 64
px = [[(0, 0, 0, 0)] * N for _ in range(N)]
def rect(x0, y0, x1, y1, c):
    for y in range(y0, y1):
        for x in range(x0, x1):
            px[y][x] = c
BLUE, DARK, WHITE, GREY = (46, 116, 181, 255), (31, 56, 100, 255), (255, 255, 255, 255), (200, 210, 225, 255)
# to giay
rect(12, 4, 52, 60, DARK)
rect(14, 6, 50, 58, WHITE)
# goc gap
for y in range(4, 16):
    for x in range(40, 52):
        if x - 40 >= y - 4 + 0:
            pass
rect(38, 4, 52, 18, (255, 255, 255, 0))
for i in range(14):
    for j in range(14):
        x, y = 38 + j, 4 + i
        if j + i >= 13: px[y][x] = GREY
        if j + i in (13,): px[y][x] = DARK
rect(38, 4, 39, 18, DARK) if False else None
# cac dong chu mo
for y in (22, 28, 34):
    rect(19, y, 45, y + 2, GREY)
# o so trang o chan: nen xanh + chu "1"
rect(20, 40, 44, 54, BLUE)
rect(31, 42, 34, 52, WHITE)      # than so 1
rect(28, 44, 32, 46, WHITE)      # co so 1
rect(28, 51, 37, 53, WHITE)      # chan so 1
raw = b"".join(b"\0" + b"".join(bytes(p) for p in row) for row in px)
def chunk(t, d): return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", N, N, 8, 6, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")
out = ROOT / "customUI" / "images" / "pagenum.png"
out.parent.mkdir(parents=True, exist_ok=True)
out.write_bytes(png)
print("Da tao", out)

#!/usr/bin/env python3
"""
Tao vbaProject.bin (OLE + MS-OVBA) tu dist/*.bas bang Python thuan, khong can Word.
Dung: python3 tools/make_vbaproject.py [-o out.bin]
Sau do: python3 build/build_dotm.py --vba out.bin
"""
import argparse, struct, uuid, math
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FREE, EOC, FATSECT, NONE = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF

# ---------- MS-OVBA compression ----------
def compress_chunk(d: bytes) -> bytes:
    out = bytearray(); pos = 0
    while pos < len(d):
        fp = len(out); out.append(0); flag = 0
        for bit in range(8):
            if pos >= len(d): break
            best_off = best_len = 0
            if pos >= 1 and pos + 3 <= len(d):
                bc = max((pos - 1).bit_length(), 4)
                maxlen = (0xFFFF >> bc) + 3
                s = d.rfind(d[pos:pos + 3], 0, pos + 2)
                if s != -1:
                    l = 0
                    while pos + l < len(d) and l < maxlen and d[s + l] == d[pos + l]:
                        l += 1
                    best_off, best_len = pos - s, l
            if best_len >= 3:
                bc = max((pos - 1).bit_length(), 4)
                tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack("<H", tok); flag |= 1 << bit; pos += best_len
            else:
                out.append(d[pos]); pos += 1
        out[fp] = flag
    return bytes(out)

def compress(data: bytes) -> bytes:
    res = bytearray([1])
    for i in range(0, max(len(data), 1), 4096):
        piece = data[i:i + 4096]
        c = compress_chunk(piece)
        if len(c) + 2 > 4098 or len(piece) < 1:
            res += struct.pack("<H", 4095 | 0x3000) + piece.ljust(4096, b"\0")
        else:
            res += struct.pack("<H", (len(c) + 2 - 3) | 0x3000 | 0x8000) + c
    return bytes(res)

def decompress(b: bytes) -> bytes:
    assert b[0] == 1
    p = 1; out = bytearray()
    while p < len(b):
        h = struct.unpack_from("<H", b, p)[0]; size = (h & 0xFFF) + 3
        end = p + size; start = len(out); q = p + 2
        if h & 0x8000:
            while q < end:
                f = b[q]; q += 1
                for bit in range(8):
                    if q >= end: break
                    if f >> bit & 1:
                        tok = struct.unpack_from("<H", b, q)[0]; q += 2
                        bc = max((len(out) - start - 1).bit_length(), 4)
                        off = (tok >> (16 - bc)) + 1; ln = (tok & (0xFFFF >> bc)) + 3
                        for _ in range(ln): out.append(out[-off])
                    else:
                        out.append(b[q]); q += 1
        else:
            out += b[q:q + 4096]
        p = end
    return bytes(out)

# ---------- VBA streams ----------
def rec(i, data=b""): return struct.pack("<HI", i, len(data)) + data
def u16(s): return s.encode("utf-16-le")

REFS = [
    ("VBA", r"*\G{000204EF-0000-0000-C000-000000000046}#4.2#9#C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL#Visual Basic For Applications"),
    ("Word", r"*\G{00020905-0000-0000-C000-000000000046}#8.7#0#C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB#Microsoft Word 16.0 Object Library"),
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
    ("Normal", None),
    ("MSForms", r"*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\Windows\System32\FM20.DLL#Microsoft Forms 2.0 Object Library"),
    ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.8#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
]

def build_dir(mods):
    d = bytearray()
    d += rec(1, struct.pack("<I", 1))            # SYSKIND win32
    d += rec(2, struct.pack("<I", 0x409))
    d += rec(0x14, struct.pack("<I", 0x409))
    d += rec(3, struct.pack("<H", 1252))
    d += rec(4, b"VBAProject")
    d += rec(5) + rec(0x40)
    d += rec(6) + rec(0x3D)
    d += rec(7, struct.pack("<I", 0))
    d += rec(8, struct.pack("<I", 0))
    d += struct.pack("<HIIH", 9, 4, 1, 0)          # PROJECTVERSION
    d += rec(0x0C) + rec(0x3C)
    for nm, libid in REFS:
        d += rec(0x16, nm.encode("ascii")) + rec(0x3E, u16(nm))
        if libid is None:   # reference to project Normal
            la = b"*\\CNormal"
            d += rec(0x0E, struct.pack("<I", len(la)) + la + struct.pack("<I", len(la)) + la + struct.pack("<IH", 0, 0))
            continue
        lb = libid.encode("ascii")
        d += rec(0x0D, struct.pack("<I", len(lb)) + lb + struct.pack("<IH", 0, 0))
    d += rec(0x0F, struct.pack("<H", len(mods)))
    d += rec(0x13, struct.pack("<H", 0xFFFF))
    for name, kind, _ in mods:
        n = name.encode("ascii")
        d += rec(0x19, n) + rec(0x47, u16(name))
        d += rec(0x1A, n) + rec(0x32, u16(name))
        d += rec(0x1C) + rec(0x48)
        d += rec(0x31, struct.pack("<I", 0))
        d += rec(0x1E, struct.pack("<I", 0))
        d += rec(0x2C, struct.pack("<H", 0xFFFF))
        d += rec(0x21 if kind == "std" else 0x22)
        d += rec(0x2B)
    d += rec(0x10)
    return bytes(d)

THISDOC = ('Attribute VB_Name = "ThisDocument"\r\n'
           'Attribute VB_Base = "1Normal.ThisDocument"\r\n'
           'Attribute VB_GlobalNameSpace = False\r\n'
           'Attribute VB_Creatable = False\r\n'
           'Attribute VB_PredeclaredId = True\r\n'
           'Attribute VB_Exposed = True\r\n'
           'Attribute VB_TemplateDerived = False\r\n'
           'Attribute VB_Customizable = True\r\n')

def vba_streams(bas_files):
    mods = [("ThisDocument", "doc", THISDOC.encode("ascii"))]
    for f in bas_files:
        src = f.read_bytes().replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        mods.append((f.stem, "std", src))
    project = ('ID="{%s}"\r\nDocument=ThisDocument/&H00000000\r\n' % str(uuid.uuid4()).upper())
    for n, k, _ in mods:
        if k == "std": project += "Module=%s\r\n" % n
    project += 'Name="VBAProject"\r\nHelpContextID="0"\r\nVersionCompatible32="393222000"\r\n'
    project += "\r\n[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n"
    wm = b"".join(n.encode("ascii") + b"\0" + u16(n) + b"\0\0" for n, _, _ in mods) + b"\0\0"
    streams = {
        "PROJECT": project.encode("ascii"),
        "PROJECTwm": wm,
        "VBA/_VBA_PROJECT": b"\xCC\x61\xFF\xFF\x00\x00\x00",
        "VBA/dir": compress(build_dir(mods)),
    }
    for n, _, src in mods:
        streams["VBA/" + n] = compress(src)
    return streams, mods

# ---------- CFB writer ----------
def key(n): return (len(n), n.upper())

def build_cfb(streams):
    # entries: 0 root, then VBA storage, then streams
    ents = [{"name": "Root Entry", "type": 5, "children": []}]
    vba = {"name": "VBA", "type": 1, "children": []}
    ents.append(vba)
    root_kids, vba_kids = [1], []
    for path, data in streams.items():
        e = {"name": path.split("/")[-1], "type": 2, "data": data, "children": []}
        ents.append(e); idx = len(ents) - 1
        (vba_kids if path.startswith("VBA/") else root_kids).append(idx)
    def link(parent, kids):
        kids = sorted(kids, key=lambda i: key(ents[i]["name"]))
        def bst(lo, hi):
            if lo >= hi: return NONE
            m = (lo + hi) // 2; i = kids[m]
            ents[i]["left"] = bst(lo, m); ents[i]["right"] = bst(m + 1, hi)
            return i
        ents[parent]["child"] = bst(0, len(kids))
    link(0, root_kids); link(1, vba_kids)
    for e in ents:
        e.setdefault("left", NONE); e.setdefault("right", NONE); e.setdefault("child", NONE)
    # mini stream
    mini = bytearray(); minifat = []
    for e in ents:
        if e["type"] == 2 and len(e["data"]) < 4096:
            e["start"] = len(mini) // 64
            n = math.ceil(len(e["data"]) / 64) or 1
            mini += e["data"].ljust(n * 64, b"\0")
            minifat += list(range(e["start"] + 1, e["start"] + n)) + [EOC]
    S = 512
    nsec = lambda b: math.ceil(b / S)
    n_dir = nsec(len(ents) * 128)
    n_mf = nsec(len(minifat) * 4)
    n_ms = nsec(len(mini))
    big = [e for e in ents if e["type"] == 2 and len(e["data"]) >= 4096]
    n_nonfat = n_dir + n_mf + n_ms + sum(nsec(len(e["data"])) for e in big)
    nfat = 1
    while nfat * 128 < n_nonfat + nfat: nfat += 1
    assert nfat <= 109
    fat = []; cur = 0
    def chain(count):
        nonlocal cur
        start = cur
        fat.extend(list(range(cur + 1, cur + count)) + [EOC]); cur += count
        return start
    dir_start = chain(n_dir)
    mf_start = chain(n_mf) if n_mf else EOC
    ms_start = chain(n_ms) if n_ms else EOC
    for e in big: e["start"] = chain(nsec(len(e["data"])))
    fat_start = cur
    fat += [FATSECT] * nfat
    fat += [FREE] * (nfat * 128 - len(fat))
    ents[0]["start"] = ms_start; ents[0]["size"] = len(mini)
    for e in ents:
        if e["type"] == 1: e["start"] = 0; e["size"] = 0
        if e["type"] == 2: e["size"] = len(e["data"]); e.setdefault("start", EOC)
    # dir bytes
    db = bytearray()
    for e in ents:
        nm = e["name"].encode("utf-16-le") + b"\0\0"
        db += nm.ljust(64, b"\0") + struct.pack("<HBB", len(nm), e["type"], 1)
        db += struct.pack("<III", e["left"], e["right"], e["child"]) + b"\0" * 16 + b"\0" * 20
        db += struct.pack("<IQ", e["start"], e["size"])
    db = db.ljust(n_dir * S, b"\0")
    # header
    difat = [fat_start + i for i in range(nfat)] + [FREE] * (109 - nfat)
    hdr = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\0" * 16
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
    hdr += struct.pack("<IIIIIIIII", 0, nfat, dir_start, 0, 4096, mf_start if n_mf else EOC, n_mf, EOC, 0)
    hdr += struct.pack("<109I", *difat)
    body = bytearray(db)
    body += struct.pack("<%dI" % len(minifat), *minifat).ljust(n_mf * S, b"\0") if n_mf else b""
    body += bytes(mini).ljust(n_ms * S, b"\0")
    for e in big: body += e["data"].ljust(nsec(len(e["data"])) * S, b"\0")
    body += struct.pack("<%dI" % len(fat), *fat)
    return hdr + bytes(body)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default=str(ROOT / "build" / "vbaProject.bin"))
    a = ap.parse_args()
    files = sorted((ROOT / "dist").glob("*.bas"))
    streams, mods = vba_streams(files)
    # tu kiem tra nen/giai nen
    for n, _, src in mods:
        assert decompress(streams["VBA/" + n]) == src, n
    blob = build_cfb(streams)
    Path(a.output).write_bytes(blob)
    print("Da tao %s (%d byte, %d module)" % (a.output, len(blob), len(mods)))

if __name__ == "__main__":
    main()

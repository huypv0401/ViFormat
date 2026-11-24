#!/usr/bin/env python3
"""
Chuyển src/*.bas (UTF-8, có tiếng Việt) thành dist/*.bas thuần ASCII, CRLF.

- Chuỗi chứa ký tự không phải ASCII  ->  U("...\\uXXXX...")   (hàm U trong modPlatform)
- Chú thích (') chứa dấu                ->  bỏ dấu
Lý do: VBA lưu mã nguồn theo ANSI codepage của máy nên Unicode trong source sẽ hỏng.
"""
import sys
import unicodedata
from pathlib import Path


def strip_marks(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.encode("ascii", "replace").decode("ascii")


def esc(buf: str) -> str:
    return "".join(c if ord(c) < 128 else "\\u%04x" % ord(c) for c in buf)


def flush(buf: str) -> str:
    if buf.isascii():
        return '"' + buf + '"'
    return 'U("' + esc(buf) + '")'


def process_line(line: str) -> str:
    out = []
    i, n = 0, len(line)
    in_str = False
    buf = ""
    while i < n:
        c = line[i]
        if in_str:
            if c == '"':
                if i + 1 < n and line[i + 1] == '"':
                    buf += '""'
                    i += 2
                    continue
                in_str = False
                out.append(flush(buf))
            else:
                buf += c
        else:
            if c == '"':
                in_str = True
                buf = ""
            elif c == "'":
                out.append(strip_marks(line[i:]))
                break
            else:
                out.append(c)
        i += 1
    if in_str:
        raise ValueError("Chuỗi không đóng: " + line)
    return "".join(out)


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    src, dist = root / "src", root / "dist"
    dist.mkdir(exist_ok=True)
    for f in dist.glob("*.bas"):
        f.unlink()
    count = 0
    for f in sorted(src.glob("*.bas")):
        text = f.read_text(encoding="utf-8")
        lines = [process_line(l) for l in text.splitlines()]
        data = "\r\n".join(lines) + "\r\n"
        if not data.isascii():
            print("LỖI: còn ký tự không phải ASCII trong", f.name)
            return 1
        (dist / f.name).write_bytes(data.encode("ascii"))
        count += 1
    print(f"OK: {count} module -> {dist}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

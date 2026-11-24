#!/usr/bin/env python3
"""
Lint tĩnh cho dist/*.bas. Không thay thế việc chạy thật trong Word, nhưng bắt được:
- API/đối tượng chỉ có trên Windows (sẽ lỗi trên Mac)
- thiếu Option Explicit
- Sub/Function/If/With/Select/For/Do không đóng cặp
- ký tự không phải ASCII còn sót
- callback trong customUI14.xml không có Sub tương ứng
"""
import re
import sys
from pathlib import Path
import xml.etree.ElementTree as ET

root = Path(__file__).resolve().parent.parent
dist = root / "dist"
problems = []

FORBIDDEN = [
    r"Scripting\.", r"FileSystemObject", r"CreateObject", r"GetObject",
    r"\bDeclare\b", r"user32", r"kernel32", r"\bShell\b", r"MacScript",
    r"WScript", r"ActiveX", r"MSForms", r"UserForm", r"\bEnviron\b",
]


def strip_strings_comments(line: str) -> str:
    out, in_s = [], False
    i = 0
    while i < len(line):
        c = line[i]
        if in_s:
            if c == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                in_s = False
        else:
            if c == '"':
                in_s = True
            elif c == "'":
                break
            else:
                out.append(c)
        i += 1
    return "".join(out)


def join_continuations(lines):
    res, cur = [], ""
    for l in lines:
        s = l.rstrip()
        if s.endswith(" _"):
            cur += s[:-1]
        else:
            res.append(cur + s)
            cur = ""
    return res


subs = set()
for f in sorted(dist.glob("*.bas")):
    raw = f.read_bytes()
    if not raw.isascii():
        problems.append(f"{f.name}: còn ký tự không phải ASCII")
    if b"\r\n" not in raw:
        problems.append(f"{f.name}: không dùng CRLF")
    text = raw.decode("ascii", "replace")
    lines = join_continuations(text.splitlines())
    if not any(l.strip() == "Option Explicit" for l in lines):
        problems.append(f"{f.name}: thiếu Option Explicit")

    stack = []
    for no, l in enumerate(lines, 1):
        code = strip_strings_comments(l).strip()
        if not code:
            continue
        for pat in FORBIDDEN:
            if re.search(pat, code):
                problems.append(f"{f.name}:{no}: dùng '{pat}' (không chạy được trên Mac)")
        m = re.match(r"(?:Public |Private )?(Sub|Function)\s+(\w+)", code)
        if m and not code.startswith("End "):
            stack.append((m.group(1), no))
            subs.add(m.group(2))
            continue
        if re.match(r"End (Sub|Function)\b", code):
            if not stack or stack[-1][0] != code.split()[1]:
                problems.append(f"{f.name}:{no}: {code} không khớp")
            else:
                stack.pop()
            continue
        if re.match(r"(?:Public |Private )?Type\s+\w+", code):
            stack.append(("Type", no))
        elif code == "End Type":
            if stack and stack[-1][0] == "Type":
                stack.pop()
        elif code.startswith("#If "):
            stack.append(("#If", no))
        elif code.startswith("#End If"):
            if stack and stack[-1][0] == "#If":
                stack.pop()
        elif re.match(r"With\b", code):
            stack.append(("With", no))
        elif code == "End With":
            if stack and stack[-1][0] == "With":
                stack.pop()
            else:
                problems.append(f"{f.name}:{no}: End With thừa")
        elif re.match(r"Select Case\b", code):
            stack.append(("Select", no))
        elif code == "End Select":
            if stack and stack[-1][0] == "Select":
                stack.pop()
            else:
                problems.append(f"{f.name}:{no}: End Select thừa")
        elif re.match(r"If\b.*\bThen$", code):
            stack.append(("If", no))
        elif code == "End If":
            if stack and stack[-1][0] == "If":
                stack.pop()
            else:
                problems.append(f"{f.name}:{no}: End If thừa")
        elif re.match(r"For\b", code):
            stack.append(("For", no))
        elif re.match(r"Next\b", code):
            if stack and stack[-1][0] == "For":
                stack.pop()
            else:
                problems.append(f"{f.name}:{no}: Next thừa")
        elif re.match(r"Do\b", code):
            stack.append(("Do", no))
        elif re.match(r"Loop\b", code):
            if stack and stack[-1][0] == "Do":
                stack.pop()
            else:
                problems.append(f"{f.name}:{no}: Loop thừa")
    for kind, no in stack:
        if kind != "#If":
            problems.append(f"{f.name}:{no}: {kind} không được đóng")

# callback Ribbon
ui = root / "customUI" / "customUI14.xml"
tree = ET.parse(ui)
cb_attrs = {"onLoad", "onAction", "getItemCount", "getItemID", "getItemLabel", "getSelectedItemID"}
for el in tree.iter():
    for a, v in el.attrib.items():
        if a in cb_attrs and v not in subs:
            problems.append(f"customUI14.xml: callback '{v}' ({a}) chưa có Sub trong VBA")

if problems:
    print("\n".join(problems))
    print(f"\n{len(problems)} vấn đề")
    sys.exit(1)
print(f"Lint OK: {len(list(dist.glob('*.bas')))} module, {len(subs)} Sub/Function, callback Ribbon đều có Sub")

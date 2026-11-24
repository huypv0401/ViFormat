#!/bin/bash
# Cai ViFormat cho Word tren macOS: copy ViFormat.dotm vao thu muc Startup cua Word.
# Chay: nhap dup file nay (hoac: bash install.command)
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
SRC="$HERE/ViFormat.dotm"
[ -f "$SRC" ] || SRC="$HERE/../ViFormat.dotm"
[ -f "$SRC" ] || { echo "Khong thay ViFormat.dotm"; exit 1; }

if pgrep -x "Microsoft Word" >/dev/null; then
  echo "Hay thoat Word truoc khi cai, roi chay lai."
  exit 1
fi

DEST="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
mkdir -p "$DEST"
cp -f "$SRC" "$DEST/ViFormat.dotm"
xattr -d com.apple.quarantine "$DEST/ViFormat.dotm" 2>/dev/null || true
# Script hien thong bao tieng Viet (Unicode) cho Word Mac
AS_DIR="$HOME/Library/Application Scripts/com.microsoft.Word"
mkdir -p "$AS_DIR"
cp -f "$HERE/ViFormat.js" "$AS_DIR/ViFormat.js"
osacompile -o "$AS_DIR/ViFormat.scpt" "$HERE/ViFormat.applescript" && echo "Da cai script thong bao vao $AS_DIR"
echo "Da cai ViFormat vao:"
echo "$DEST"
echo "Mo Word, tab ViFormat se xuat hien. Lan dau co the can bat macro khi Word hoi."

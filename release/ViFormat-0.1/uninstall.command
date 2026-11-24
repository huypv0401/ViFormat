#!/bin/bash
# Go ViFormat khoi Word tren macOS.
# Chay: bash uninstall.command
if pgrep -x "Microsoft Word" >/dev/null; then
  echo "Hay thoat Word truoc khi go, roi chay lai."
  exit 1
fi
W="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
A="$HOME/Library/Application Scripts/com.microsoft.Word"
rm -f "$W/ViFormat.dotm" "$W/~\$Format.dotm" "$A/ViFormat.scpt" "$A/ViFormat.js"
echo "Da go ViFormat. Mo lai Word, tab ViFormat se khong con."

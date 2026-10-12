#!/bin/bash
# Mac 더블클릭: ~/Documents/DesktopOrganizer/Folder 1 에 PNG 저장
set -euo pipefail
DEST="$HOME/Documents/DesktopOrganizer/Folder 1"
PNG="새 Mac 앱 구성 (Mole 제외, 체험일 표기).png"
RAW_BASE="https://github.com/For1916/1916/raw/cursor/mac-app-infographic-no-mole-de41/DesktopOrganizer/Folder%201"
ENC_PNG=$(python3 -c "import urllib.parse; print(urllib.parse.quote('$PNG'))")

mkdir -p "$DEST"
curl -fsSL "$RAW_BASE/$ENC_PNG" -o "$DEST/$PNG"
open -R "$DEST/$PNG"
osascript -e "display notification \"문서 폴더에 저장됨\" with title \"$PNG\""

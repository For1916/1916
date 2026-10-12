#!/bin/bash
# Mac에서 더블클릭: 이 스크립트는 **DesktopOrganizer/Folder 1** 안에 있어야 합니다.
set -euo pipefail
DEST="$(cd "$(dirname "$0")" && pwd)"
PNG_NAME="새 Mac 앱 구성 (Mole 제외, 체험일 표기).png"
RAW="https://github.com/For1916/1916/raw/cursor/mac-app-infographic-no-mole-de41/DesktopOrganizer/Folder%201/$(python3 -c "import urllib.parse; print(urllib.parse.quote('$PNG_NAME'))")"

if [[ "$(basename "$DEST")" != "Folder 1" ]] || [[ "$(basename "$(dirname "$DEST")")" != "DesktopOrganizer" ]]; then
  osascript -e 'display alert "경로 오류" message "이 파일은 Mac의 DesktopOrganizer/Folder 1 폴더에 두고 실행하세요." as critical'
  exit 1
fi

curl -fsSL "$RAW" -o "$DEST/$PNG_NAME"
cp "$DEST/mac-app-config.html" "$DEST/mac-app-config.html" 2>/dev/null || true
open -R "$DEST/$PNG_NAME"
osascript -e "display notification \"저장 완료: $DEST/$PNG_NAME\" with title \"Mac 앱 구성 PNG\""

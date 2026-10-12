#!/bin/bash
# Run on your Mac (BetterTouchTool installed) to print trial dates for the infographic.
set -euo pipefail
PLIST="$HOME/Library/Preferences/com.folivora.BetterTouchTool.plist"
if [[ ! -f "$PLIST" ]]; then
  echo "BetterTouchTool preferences not found: $PLIST" >&2
  exit 1
fi
echo "=== BetterTouchTool (from plist) ==="
/usr/bin/plutil -p "$PLIST" 2>/dev/null | grep -iE 'trial|license|expir|date|start|end|valid' || plutil -p "$PLIST" | head -80
echo ""
echo "Open app: BetterTouchTool → Settings → License (screenshot that screen for exact end date)."

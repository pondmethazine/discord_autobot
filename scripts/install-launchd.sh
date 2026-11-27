#!/bin/bash
# Run the bot as a macOS LaunchAgent: starts at login, restarts if it crashes.
#   ./scripts/install-launchd.sh            install / reinstall and start
#   ./scripts/install-launchd.sh uninstall  stop and remove
# Logs: data/bot.log (stdout+stderr)
set -euo pipefail

LABEL="com.pondmethazine.work-assistant"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
DIR="$(cd "$(dirname "$0")/.." && pwd)"
NODE="$(command -v node)"

if [[ "${1:-}" == "uninstall" ]]; then
  launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "removed $LABEL"
  exit 0
fi

[[ -f "$DIR/.env" ]] || { echo "missing $DIR/.env (copy .env.example and fill it in)"; exit 1; }
grep -q '^DISCORD_TOKEN=.\+' "$DIR/.env" || { echo "DISCORD_TOKEN is empty in .env"; exit 1; }
mkdir -p "$DIR/data" "$HOME/Library/LaunchAgents"

# launchd starts with a bare PATH; bake in the current one so git, gcloud, firebase, claude resolve.
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array><string>$NODE</string><string>$DIR/src/index.js</string></array>
  <key>WorkingDirectory</key><string>$DIR</string>
  <key>EnvironmentVariables</key>
  <dict>
    <key>PATH</key><string>$PATH</string>
    <key>HOME</key><string>$HOME</string>
  </dict>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><dict><key>SuccessfulExit</key><false/></dict>
  <key>ThrottleInterval</key><integer>30</integer>
  <key>StandardOutPath</key><string>$DIR/data/bot.log</string>
  <key>StandardErrorPath</key><string>$DIR/data/bot.log</string>
</dict>
</plist>
EOF

launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"
echo "started $LABEL — tail -f $DIR/data/bot.log"

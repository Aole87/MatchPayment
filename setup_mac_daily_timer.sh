#!/bin/bash
# สคริปต์ติดตั้งระบบตั้งเวลาทำงานอัตโนมัติบน Mac (LaunchAgent) ทุก 06:00 น.
PLIST_PATH="$HOME/Library/LaunchAgents/com.matchpayment.daily.plist"
PROJECT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
PYTHON_BIN="$PROJECT_DIR/.venv/bin/python3"

echo "⏰ กำลังติดตั้งระบบตั้งเวลาทำงาน 06:00 น. ทุกวันบน Mac..."

mkdir -p "$HOME/Library/LaunchAgents"

cat <<EOF > "$PLIST_PATH"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.matchpayment.daily</string>
    <key>ProgramArguments</key>
    <array>
        <string>$PYTHON_BIN</string>
        <string>$PROJECT_DIR/match_bot.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>$PROJECT_DIR</string>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>6</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>$PROJECT_DIR/daily_cron.log</string>
    <key>StandardErrorPath</key>
    <string>$PROJECT_DIR/daily_cron.log</string>
</dict>
</plist>
EOF

launchctl unload "$PLIST_PATH" 2>/dev/null
launchctl load "$PLIST_PATH"

echo "✅ ติดตั้งเรียบร้อย! ระบบจะทำงานทุกวันเวลา 06:00 น. ตรง แม้ปิดหน้าจอ เมื่อเครื่องตื่นจะทำงานต่อทันที"


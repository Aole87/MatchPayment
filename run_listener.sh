#!/bin/bash
# สคริปต์เปิดรัน Telegram Bot Listener ให้พร้อมรับคำสั่ง "Match Data" ตลอดเวลา
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "🤖 กำลังเปิดระบบบอทรับคำสั่ง Telegram (Match Data)..."
source .venv/bin/activate
python match_bot.py --listen

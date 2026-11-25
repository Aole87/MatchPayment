import os
import threading
from flask import Flask, request, jsonify
from match_bot import (
    TELEGRAM_TOKEN,
    TELEGRAM_CHAT_ID,
    send_telegram,
    send_chat_action,
    extract_two_files_from_text,
    scan_and_group_files,
    resolve_target_date,
    run_matching_job,
    check_files_status,
    download_drive_folder,
    logger
)

app = Flask(__name__)


def process_telegram_update(data):
    """ฟังก์ชันประมวลผลคำสั่งจาก Telegram Webhook ใน Background Thread"""
    try:
        msg = data.get("message") or data.get("edited_message")
        if not msg:
            return

        chat = msg.get("chat", {})
        chat_id = chat.get("id")
        text = str(msg.get("text", "")).strip()
        if not text:
            return

        logger.info(f"ได้รับข้อความ Webhook จาก Chat ID {chat_id}: '{text}'")
        text_clean = text.lower().strip()
        send_chat_action(chat_id, "typing")

        # 1. ตรวจสอบว่าระบุชื่อ 2 ไฟล์โดยตรงหรือไม่
        file_pair = extract_two_files_from_text(text)
        if not file_pair and (".xls" in text_clean or "report" in text_clean):
            download_drive_folder()
            file_pair = extract_two_files_from_text(text)

        if file_pair:
            rep_f, mast_f = file_pair
            rep_name = os.path.basename(rep_f)
            mast_name = os.path.basename(mast_f)
            send_telegram(
                f"⏳ *น้องบิ๊นท์ได้รับคำสั่ง จับคู่ไฟล์ตามที่ระบุเรียบร้อยแล้วค่ะ!*\n"
                f"📁 *Report:* `{rep_name}`\n"
                f"📁 *Master:* `{mast_name}`\n\n"
                f"กำลังประมวลผล กรุณารอสักครู่นะคะ...",
                chat_id=chat_id
            )
            run_matching_job(report_file=rep_f, master_file=mast_f, chat_id=chat_id)
            return

        # 2. คำสั่ง: ตรวจสอบไฟล์
        if any(k in text_clean for k in ("ตรวจสอบไฟล์", "ตรวจไฟล์", "เช็คไฟล์", "status", "check", "update", "อัปเดต", "อัพเดต", "ไฟล์ใหม่", "ดูไฟล์", "มีไฟล์")):
            send_telegram("⏳ *น้องบิ๊นท์กำลังตรวจสอบไฟล์ใน Google Drive กรุณารอสักครู่นะคะ...*", chat_id=chat_id)
            check_files_status(chat_id=chat_id)
            return

        # 3. คำสั่ง: จับคู่เอกสาร
        if any(k in text_clean for k in ("จับคู่เอกสาร", "จับคู่ข้อมูล", "จับคู่", "match data", "matchdata", "match", "/match")):
            temp_groups = scan_and_group_files()
            req_date = resolve_target_date(text, temp_groups)
            date_text = f" ของวันที่ {req_date}" if req_date else " (รอบล่าสุด)"
            send_telegram(
                f"⏳ *น้องบิ๊นท์ได้รับคำสั่ง จับคู่เอกสาร{date_text} เรียบร้อยแล้วค่ะ!*\n"
                "กำลังดาวน์โหลดไฟล์และตรวจสอบความพร้อม กรุณารอสักครู่นะคะ...",
                chat_id=chat_id
            )
            run_matching_job(chat_id=chat_id, target_date=req_date)
            return

        # 4. คำสั่ง Help / Start
        if text_clean in ("/start", "/help", "help", "คำสั่ง", "เมนู", "menu"):
            send_telegram(
                "👋 *สวัสดีค่ะ! น้องบิ๊นท์ ยินดีต้อนรับสู่ระบบตรวจยอดบิล (PJ, V)*\n\n"
                "💡 *คำสั่งภาษาไทยที่รองรับ:*\n"
                "• *ตรวจสอบไฟล์* ➔ ให้น้องบิ๊นท์ตรวจเช็คว่าใน Google Drive มีไฟล์วันไหนเข้ามาบ้าง และคู่ไหนพร้อมจับคู่แล้ว\n"
                "• *จับคู่เอกสาร* ➔ เริ่มการตรวจสอบและจับคู่ข้อมูลยอดขายอัตโนมัติ (เฉพาะบิล PJ และ V)\n"
                "• *จับคู่เอกสาร [วันที่]* (เช่น `จับคู่เอกสาร 20-11` หรือ `จับคู่ 17-19/11`) ➔ ระบุวันที่ที่ต้องการจับคู่ได้ค่ะ\n"
                "• *พิมพ์ชื่อ 2 ไฟล์* (เช่น `20-11.xls Report_20261120_103902.xlsx`) ➔ เลือกจับคู่ 2 ไฟล์ที่ต้องการได้ทันทีค่ะ",
                chat_id=chat_id
            )
        else:
            send_telegram(
                "👋 *สวัสดีค่ะ น้องบิ๊นท์ ได้รับข้อความแล้วค่ะ*\n\n"
                "💡 *สามารถสั่งงานได้ดังนี้เลยนะคะ:*\n"
                "• พิมพ์ *ตรวจสอบไฟล์* ➔ ให้น้องบิ๊นท์ตรวจเช็คไฟล์ล่าสุดใน Google Drive\n"
                "• พิมพ์ *จับคู่เอกสาร* ➔ ทำการตรวจสอบและสรุปยอดบิลส่งไฟล์ Excel ให้ทันทีค่ะ\n"
                "• พิมพ์ *ชื่อ 2 ไฟล์* (เช่น `20-11.xls Report_20261120_103902.xlsx`) ➔ เลือกจับคู่ 2 ไฟล์ที่ต้องการได้ทันทีค่ะ",
                chat_id=chat_id
            )

    except Exception as e:
        logger.error(f"เกิดข้อผิดพลาดในการประมวลผล Webhook: {e}", exc_info=True)


@app.route("/", methods=["GET"])
def health_check():
    return "N'Binto Cloud Run Webhook is Active!", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json(force=True, silent=True)
    if data:
        # ตอบกลับ HTTP 200 ทันที แล้วปล่อยให้ Thread แยกประมวลผลต่อ เพื่อไม่ให้ Telegram timeout
        threading.Thread(target=process_telegram_update, args=(data,), daemon=True).start()
    return jsonify({"ok": True}), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

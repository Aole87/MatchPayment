import os
import sys
import re
import time
import glob
import logging
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xlrd
import requests
import gdown

# --- การตั้งค่าคอนฟิก (ดึงค่าจาก Environment Variables หรือใช้ค่า Default) ---
DEFAULT_DRIVE_FOLDER = "https://drive.google.com/drive/folders/17eSUsgGmTwRexYmyBRSW_OaanuwLYR3g?usp=sharing"
GDRIVE_FOLDER_URL = (os.environ.get("GDRIVE_FOLDER_URL", "").strip()) or DEFAULT_DRIVE_FOLDER

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8999661395:AAGMuIUlp0WU4rq5nk_WQs6cOblkB38D270").strip()
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "7603648423").strip()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def download_drive_folder(folder_url=None, output_dir="downloads"):
    """ดาวน์โหลดไฟล์ทั้งหมดจาก Google Drive Folder สาธารณะ"""
    if not folder_url:
        folder_url = GDRIVE_FOLDER_URL or DEFAULT_DRIVE_FOLDER
    os.makedirs(output_dir, exist_ok=True)
    logger.info(f"กำลังดาวน์โหลดไฟล์จาก Google Drive: {folder_url}")
    try:
        downloaded = gdown.download_folder(
            url=folder_url,
            output=output_dir,
            quiet=False,
            use_cookies=False
        )
        logger.info(f"ดาวน์โหลดสำเร็จ ได้รับไฟล์: {downloaded}")
        return output_dir
    except Exception as e:
        logger.warning(f"gdown.download_folder API error: {e}, กำลังลองด้วย gdown CLI...")
        try:
            import subprocess
            cmd = [sys.executable, "-m", "gdown", "--folder", folder_url, "-O", output_dir]
            logger.info(f"รันคำสั่ง CLI: {' '.join(cmd)}")
            subprocess.run(cmd, check=True)
            return output_dir
        except Exception as e2:
            logger.error(f"เกิดข้อผิดพลาดในการดาวน์โหลดด้วย gdown CLI: {e2}")
            return output_dir


def find_target_files(search_dirs=None):
    """
    ค้นหาไฟล์ Report ล่าสุด และไฟล์ Master (เช่น 5-11.xls)
    โดยรองรับการเปลี่ยนชื่อไฟล์ทุกวันอัตโนมัติ
    """
    # ตรวจสอบโฟลเดอร์ downloads ก่อนเป็นอันดับแรก หากมีไฟล์ที่เพิ่งดาวน์โหลดมาให้ใช้โฟลเดอร์นี้เท่านั้น
    if os.path.exists("downloads") and any(f.lower().endswith((".xlsx", ".xls")) for f in os.listdir("downloads")):
        search_dirs = ["downloads"]
    elif search_dirs is None:
        search_dirs = ["downloads", "."]

    all_files = []
    for d in search_dirs:
        if os.path.exists(d):
            for f in os.listdir(d):
                fp = os.path.join(d, f)
                if os.path.isfile(fp):
                    name_lower = f.lower()
                    if name_lower.endswith((".xlsx", ".xls")) and not f.startswith("~$") and not name_lower.startswith("match_result"):
                        all_files.append(fp)

    if not all_files:
        raise FileNotFoundError("ไม่พบไฟล์ Excel ในโฟลเดอร์สำหรับประมวลผล")

    # แยกไฟล์ Report และไฟล์ Master
    report_candidates = [f for f in all_files if "report" in os.path.basename(f).lower()]
    master_candidates = [f for f in all_files if "report" not in os.path.basename(f).lower()]

    if not report_candidates:
        raise FileNotFoundError("ไม่พบไฟล์ประเภท Report_*.xlsx")
    if not master_candidates:
        raise FileNotFoundError("ไม่พบไฟล์ประเภท Master (เช่น 5-11.xls)")

    def sort_key(filepath):
        """จัดเรียงตามเวลาแก้ไขไฟล์ และชื่อไฟล์ที่มีตัวเลขวันที่กำกับ"""
        fname = os.path.basename(filepath)
        mtime = os.path.getmtime(filepath)
        nums = re.findall(r"\d+", fname)
        num_score = int("".join(nums)) if nums else 0
        return (mtime, num_score, fname)

    # เลือกไฟล์ล่าสุด
    latest_report = max(report_candidates, key=sort_key)
    latest_master = max(master_candidates, key=sort_key)

    logger.info(f"ไฟล์ Report ที่เลือก: {latest_report}")
    logger.info(f"ไฟล์ Master ที่เลือก: {latest_master}")
    return latest_report, latest_master


def normalize_doc_no(val):
    """
    แปลงค่ารหัสเอกสารให้เป็นมาตรฐาน PJXXXXXXX
    ถ้าในคอลัมน์ไม่มีเลข หรือเป็นค่าว่าง จะไม่พยายามสร้างค่าขึ้นมาเองเด็ดขาด
    """
    if val is None:
        return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("nan", "none", "-", ""):
        return ""

    # ค้นหาคำว่า PJ ตามด้วยตัวเลข เช่น PJ1547564
    match_pj = re.search(r"PJ\s*(\d+)", val_str, re.IGNORECASE)
    if match_pj:
        return f"PJ{match_pj.group(1)}"

    # ถ้ามีตัวเลขชัดเจนในหมายเหตุ เช่น 1547564
    match_num = re.search(r"(\d{4,})", val_str)
    if match_num:
        return f"PJ{match_num.group(1)}"

    return ""


def parse_amount(val):
    """แปลงค่ายอดเงินเป็น float ปลอดภัย"""
    if val is None or pd.isna(val):
        return 0.0
    val_str = str(val).replace(",", "").replace(" ", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def read_report_file(filepath):
    """
    อ่านไฟล์ Report_*.xlsx
    สกัด: วันเวลา, รหัสรายการ, ยอดเงิน (Col D), หมายเหตุ (Col I)
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    rows_data = []
    for r in range(2, ws.max_row + 1):
        time_val = ws.cell(r, 1).value
        code_val = ws.cell(r, 2).value
        type_val = ws.cell(r, 3).value
        amount_val = ws.cell(r, 4).value
        status_val = ws.cell(r, 8).value
        note_val = ws.cell(r, 9).value

        # กรองแถวว่าง
        if time_val is None and code_val is None and amount_val is None and note_val is None:
            continue

        # กรองรายการปิดสิ้นวัน/เติมเงินทอน/นำเงินออก ที่ไม่ใช่บิลขายและไม่มีหมายเหตุบิล
        type_str = str(type_val or "").strip()
        note_str = str(note_val or "").strip()

        if type_str in ("ปิดสิ้นวัน", "เติมเงินทอน", "ถอนเงิน", "นำเงินออกระหว่างวัน") and not note_str:
            continue

        doc_from_note = normalize_doc_no(note_val)
        # ตรวจสอบ Col B เฉพาะกรณีที่มีคำว่า PJ ชัดเจนเท่านั้น (ห้ามแปลงวันที่ในรหัส SAL_ เด็ดขาด)
        doc_from_code = normalize_doc_no(code_val) if (code_val and "PJ" in str(code_val).upper()) else ""

        pj_doc = doc_from_note or doc_from_code

        # ถ้าในคอลัมน์ไม่มีเลข ไม่ต้องพยายามจับคู่ ข้ามรายการนี้ไป
        if not pj_doc:
            continue

        rows_data.append({
            "report_row": r,
            "datetime": str(time_val or "").strip(),
            "tx_code": str(code_val or "").strip(),
            "tx_type": str(type_val or "").strip(),
            "amount": parse_amount(amount_val),
            "status": str(status_val or "").strip(),
            "note_raw": str(note_val or "").strip(),
            "pj_doc": pj_doc
        })

    logger.info(f"อ่านไฟล์ Report สำเร็จ: {len(rows_data)} แถว")
    return rows_data


def read_master_file(filepath):
    """
    อ่านไฟล์ 5-11.xls (หรือไฟล์ Master ประจำวัน)
    รองรับการถอดรหัสภาษาไทย CP874 จาก Excel 97-2003
    สกัด: เลขที่เอกสาร (Col B - PJ...), ชื่อลูกหนี้ (Col D), มูลค่าสุทธิ (Col M)
    """
    if filepath.lower().endswith(".xls"):
        try:
            wb = xlrd.open_workbook(filepath, encoding_override="cp874")
        except Exception:
            wb = xlrd.open_workbook(filepath)
        sh = wb.sheet_by_index(0)

        master_data = {}
        for r in range(sh.nrows):
            if sh.ncols < 2:
                continue

            doc_cell = sh.cell_value(r, 1)  # Col B
            doc_str = str(doc_cell).strip()
            norm_doc = normalize_doc_no(doc_str)

            if norm_doc.startswith("PJ"):
                cust_name = str(sh.cell_value(r, 3)).strip() if sh.ncols > 3 else ""
                val_before = parse_amount(sh.cell_value(r, 8)) if sh.ncols > 8 else 0.0
                val_net = parse_amount(sh.cell_value(r, 12)) if sh.ncols > 12 else 0.0

                master_data[norm_doc] = {
                    "master_row": r + 1,
                    "doc_no_raw": doc_str,
                    "customer_name": cust_name,
                    "val_before": val_before,
                    "val_net": val_net
                }
    else:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        master_data = {}
        for r in range(1, ws.max_row + 1):
            doc_cell = ws.cell(r, 2).value  # Col B
            doc_str = str(doc_cell or "").strip()
            norm_doc = normalize_doc_no(doc_str)

            if norm_doc.startswith("PJ"):
                cust_name = str(ws.cell(r, 4).value or "").strip()
                val_before = parse_amount(ws.cell(r, 9).value)
                val_net = parse_amount(ws.cell(r, 13).value)

                master_data[norm_doc] = {
                    "master_row": r,
                    "doc_no_raw": doc_str,
                    "customer_name": cust_name,
                    "val_before": val_before,
                    "val_net": val_net
                }

    logger.info(f"อ่านไฟล์ Master สำเร็จ: พบเอกสาร PJ ทั้งหมด {len(master_data)} รายการ")
    return master_data


def process_matching(report_rows, master_dict):
    """
    ทำการ Match ข้อมูลระหว่าง Report และ Master:
    - Match เลขที่เอกสาร PJ
    - ตรวจสอบยอดเงิน (Report ยอดเงิน Col D == Master มูลค่าสุทธิ Col M)
    """
    matched_list = []
    mismatch_list = []
    not_found_report = []

    matched_master_keys = set()

    for item in report_rows:
        pj_key = item["pj_doc"]

        if not pj_key:
            continue

        if pj_key in master_dict:
            m_info = master_dict[pj_key]
            matched_master_keys.add(pj_key)

            rep_amt = item["amount"]
            mast_amt = m_info["val_net"]
            diff = round(rep_amt - mast_amt, 2)

            detail = {
                "Report แถว": item["report_row"],
                "วันเวลาใน Report": item["datetime"],
                "เลขที่บิล (PJ)": pj_key,
                "หมายเหตุ Report": item["note_raw"],
                "ยอดเงิน Report (Col D)": rep_amt,
                "Master แถว": m_info["master_row"],
                "เลขที่เอกสาร Master (Col B)": m_info["doc_no_raw"],
                "ชื่อลูกค้า Master (Col D)": m_info["customer_name"],
                "มูลค่าสุทธิ Master (Col M)": mast_amt,
                "ผลต่าง (Report - Master)": diff,
            }

            if abs(diff) < 0.01:
                detail["สถานะ"] = "ตรงกัน"
                matched_list.append(detail)
            else:
                detail["สถานะ"] = "ยอดเงินไม่ตรงกัน"
                detail["หมายเหตุปัญหา"] = f"ยอดเงินใน Report ({rep_amt:,.2f}) ไม่ตรงกับ Master ({mast_amt:,.2f})"
                mismatch_list.append(detail)
        else:
            detail = {
                "Report แถว": item["report_row"],
                "วันเวลาใน Report": item["datetime"],
                "เลขที่บิลที่สกัดได้": pj_key,
                "หมายเหตุ Report (ดิบ)": item["note_raw"],
                "ยอดเงิน Report (Col D)": item["amount"],
                "สถานะ": "ไม่พบใน Master",
                "หมายเหตุปัญหา": f"ไม่พบรหัส {pj_key} ในไฟล์ Master"
            }
            not_found_report.append(detail)

    # ค้นหาเอกสารใน Master ที่ไม่มีใน Report
    not_in_report = []
    for doc_no, m_info in master_dict.items():
        if doc_no not in matched_master_keys:
            not_in_report.append({
                "Master แถว": m_info["master_row"],
                "เลขที่เอกสาร Master (Col B)": m_info["doc_no_raw"],
                "ชื่อลูกค้า Master (Col D)": m_info["customer_name"],
                "มูลค่าสุทธิ Master (Col M)": m_info["val_net"],
                "สถานะ": "ไม่มีใน Report"
            })

    return {
        "matched": matched_list,
        "mismatched": mismatch_list,
        "not_found_report": not_found_report,
        "not_in_report": not_in_report,
        "total_report_pj": len(matched_list) + len(mismatch_list) + len(not_found_report),
        "total_master_pj": len(master_dict)
    }


def style_excel_sheet(ws, title_color="1F4E79"):
    """จัดแต่งหน้าตาตาราง Excel ให้สวยงาม อ่านง่ายระดับมืออาชีพ"""
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=title_color, end_color=title_color, fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    # Style Header Row
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(1, col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    # Style Data Rows
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 20
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border

            if isinstance(cell.value, (int, float)):
                header_name = str(ws.cell(1, c).value or "")
                if "แถว" in header_name or "ลำดับ" in header_name:
                    cell.alignment = align_center
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = align_right
                    cell.number_format = "#,##0.00"
            else:
                cell.alignment = align_left

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or "")
            length = sum(1.5 if ord(char) > 128 else 1 for char in val)
            if length > max_len:
                max_len = length
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)


def create_result_excel(result, report_name, master_name, output_path="match_result.xlsx"):
    """สร้างไฟล์ match_result.xlsx ที่มีหลายชีต ครอบคลุมทั้งสรุปผลและรายละเอียด"""
    wb = openpyxl.Workbook()

    # --- Sheet 1: สรุปภาพรวม (Summary) ---
    ws_summary = wb.active
    ws_summary.title = "สรุปภาพรวม"
    ws_summary.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_summary.merge_cells("A1:D1")
    title_cell = ws_summary["A1"]
    title_cell.value = "📊 รายงานสรุปผลการ Match ข้อมูลยอดชำระเงินประจำวัน"
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 35

    summary_rows = [
        ("วันที่ประมวลผล", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("ไฟล์ Report (รายการรับชำระเงิน)", os.path.basename(report_name)),
        ("ไฟล์ Master (ยอดขาย/ใบแจ้งหนี้)", os.path.basename(master_name)),
        ("──────────────────────────────", "──────────────────"),
        ("1. ข้อมูลที่ตรงกันสมบูรณ์ (Matched)", len(result["matched"])),
        ("2. ข้อมูลที่ยอดเงินไม่ตรงกัน (Mismatch)", len(result["mismatched"])),
        ("3. ข้อมูลใน Report ที่ไม่พบใน Master", len(result["not_found_report"])),
        ("รวมรายการ Report ที่ตรวจเช็คทั้งหมด", result["total_report_pj"]),
        ("──────────────────────────────", "──────────────────"),
        ("4. บิลใน Master ที่ยังไม่มีใน Report", len(result["not_in_report"])),
        ("รวมบิล PJ ใน Master ทั้งหมด", result["total_master_pj"]),
    ]

    for idx, (label, val) in enumerate(summary_rows, start=3):
        ws_summary.cell(idx, 1, label)
        c2 = ws_summary.cell(idx, 2, val)
        ws_summary.cell(idx, 1).font = Font(name="Segoe UI", size=11, bold=True)
        c2.font = Font(name="Segoe UI", size=11, bold=isinstance(val, int))
        if isinstance(val, int):
            c2.alignment = Alignment(horizontal="right")
            c2.number_format = "#,##0"

    ws_summary.cell(7, 2).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    ws_summary.cell(8, 2).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    ws_summary.cell(9, 2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    ws_summary.column_dimensions["A"].width = 42
    ws_summary.column_dimensions["B"].width = 35

    def add_data_sheet(sheet_name, data_list, title_color):
        ws = wb.create_sheet(sheet_name)
        if not data_list:
            ws.cell(1, 1, "ไม่มีรายการ")
            return
        df = pd.DataFrame(data_list)
        df.insert(0, "ลำดับ", range(1, len(df) + 1))

        headers = list(df.columns)
        for col_idx, h in enumerate(headers, start=1):
            ws.cell(1, col_idx, h)

        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(r_idx, c_idx, val)

        style_excel_sheet(ws, title_color=title_color)

    add_data_sheet("ข้อมูลตรงกัน", result["matched"], "2E75B6")
    add_data_sheet("ยอดเงินไม่ตรงกัน", result["mismatched"], "C65911")
    add_data_sheet("ไม่พบใน Master", result["not_found_report"], "BF8F00")
    add_data_sheet("Master ไม่มีใน Report", result["not_in_report"], "595959")

    wb.save(output_path)
    logger.info(f"บันทึกผลลัพธ์ลงไฟล์ Excel เรียบร้อยแล้ว: {output_path}")
    return output_path


def send_telegram(summary_text, file_path=None, chat_id=None):
    """ส่งข้อความสรุปผลและส่งไฟล์รายงานเข้า Telegram"""
    target_chat = str(chat_id or TELEGRAM_CHAT_ID).strip()
    if not TELEGRAM_TOKEN or not target_chat:
        logger.warning("⚠️ ไม่พบ TELEGRAM_TOKEN หรือ TELEGRAM_CHAT_ID ในระบบ")
        return False

    # 1. ส่งข้อความสรุป
    msg_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": target_chat,
        "text": summary_text,
        "parse_mode": "Markdown"
    }
    try:
        res = requests.post(msg_url, data=payload, timeout=20)
        if res.status_code != 200:
            payload.pop("parse_mode", None)
            res = requests.post(msg_url, data=payload, timeout=20)
        logger.info(f"ผลการส่งข้อความ Telegram: {res.status_code}")
    except Exception as e:
        logger.error(f"เกิดข้อผิดพลาดในการส่งข้อความ Telegram: {e}")
        return False

    # 2. ส่งไฟล์ Excel (ถ้ามี)
    if file_path and os.path.exists(file_path):
        doc_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
        try:
            with open(file_path, "rb") as f:
                res_doc = requests.post(
                    doc_url,
                    data={
                        "chat_id": target_chat,
                        "caption": f"📄 รายงานฉบับเต็มประจำวัน: {os.path.basename(file_path)}"
                    },
                    files={"document": f},
                    timeout=60
                )
            logger.info(f"ผลการส่งไฟล์ Telegram: {res_doc.status_code}")
        except Exception as e:
            logger.error(f"เกิดข้อผิดพลาดในการส่งไฟล์ Telegram: {e}")

    return True


def upload_to_google_drive(file_path):
    """
    อัปโหลดไฟล์ผลลัพธ์ไปยัง Google Drive Folder
    รองรับทั้ง Google Apps Script Web App (GDRIVE_UPLOAD_URL)
    และ Google Service Account (GDRIVE_SERVICE_ACCOUNT_JSON)
    """
    upload_url = os.environ.get("GDRIVE_UPLOAD_URL", "").strip()
    service_account_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON", "").strip()
    filename = os.path.basename(file_path)

    # 1. วิธีที่ 1: ผ่าน Google Apps Script Web App (วิธีที่สะดวกและง่ายที่สุด)
    if upload_url:
        logger.info(f"กำลังอัปโหลด {filename} ไปยัง Google Drive ผ่าน Web App...")
        try:
            import base64
            with open(file_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            resp = requests.post(upload_url, data={"filename": filename, "file": encoded}, timeout=60)
            if resp.status_code == 200:
                logger.info(f"✅ อัปโหลดไฟล์ {filename} ไปยัง Google Drive สำเร็จ!")
                return True
            else:
                logger.error(f"อัปโหลดล้มเหลว (Status {resp.status_code}): {resp.text}")
        except Exception as e:
            logger.error(f"เกิดข้อผิดพลาดในการอัปโหลดผ่าน Web App: {e}")

    # 2. วิธีที่ 2: ผ่าน Google Service Account JSON
    elif service_account_json:
        logger.info(f"กำลังอัปโหลด {filename} ไปยัง Google Drive ผ่าน Service Account...")
        try:
            import json
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaFileUpload

            creds_dict = json.loads(service_account_json)
            creds = service_account.Credentials.from_service_account_info(
                creds_dict,
                scopes=["https://www.googleapis.com/auth/drive.file"]
            )
            service = build("drive", "v3", credentials=creds)

            folder_id = "17eSUsgGmTwRexYmyBRSW_OaanuwLYR3g"
            file_metadata = {
                "name": filename,
                "parents": [folder_id]
            }
            media = MediaFileUpload(
                file_path,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
            logger.info(f"✅ อัปโหลดไฟล์ {filename} ไปยัง Google Drive สำเร็จ! ID: {file.get('id')}")
            return True
        except Exception as e:
            logger.error(f"เกิดข้อผิดพลาดในการอัปโหลดผ่าน Service Account: {e}")
    else:
        logger.info(
            f"ℹ️ ไฟล์ '{filename}' ถูกบันทึกไว้ในเครื่องเรียบร้อยแล้ว\n"
            f"(หากต้องการให้อัปโหลดเข้า Google Drive อัตโนมัติ สามารถเพิ่ม GDRIVE_UPLOAD_URL ใน Secrets)"
        )

    return False


def run_matching_job(chat_id=None):
    """ฟังก์ชันประมวลผลการ Match ข้อมูลหลักแบบรอบเดียว (One-off Job)"""
    logger.info("=== เริ่มต้นกระบวนการ Match ข้อมูล ===")
    target_chat = chat_id or TELEGRAM_CHAT_ID

    # 1. ดาวน์โหลดไฟล์จาก Google Drive
    download_dir = download_drive_folder()

    # 2. ค้นหาไฟล์ Report และไฟล์ Master ล่าสุด
    report_file, master_file = find_target_files([download_dir, "."])

    # 3. อ่านและเตรียมข้อมูล
    report_rows = read_report_file(report_file)
    master_dict = read_master_file(master_file)

    # 4. ทำการ Match ข้อมูล
    result = process_matching(report_rows, master_dict)

    # 5. สร้างไฟล์ Excel สรุปผล โดยตั้งชื่อเป็น match_result_ตามด้วยวันที่.xlsx
    import shutil
    date_str = datetime.now().strftime("%Y%m%d")
    output_excel = f"match_result_{date_str}.xlsx"
    create_result_excel(result, report_file, master_file, output_excel)

    # ทำสำเนาเป็น match_result.xlsx ไว้ด้วย
    shutil.copyfile(output_excel, "match_result.xlsx")

    # 6. อัปโหลดไฟล์ผลลัพธ์เข้า Google Drive
    upload_to_google_drive(output_excel)

    # 7. เตรียมข้อความสรุป
    n_match = len(result["matched"])
    n_mismatch = len(result["mismatched"])
    n_not_found = len(result["not_found_report"])
    n_total_report = result["total_report_pj"]
    n_not_in_rep = len(result["not_in_report"])

    summary_msg = (
        f"📊 *รายงานสรุปการ Match ข้อมูลประจำวัน*\n"
        f"📅 ประมวลผลเมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        f"────────────────────────\n"
        f"📁 *Report:* `{os.path.basename(report_file)}`\n"
        f"📁 *Master:* `{os.path.basename(master_file)}`\n\n"
        f"📈 *สรุปผลการตรวจสอบ:*\n"
        f"✅ *ข้อมูลตรงกัน:* {n_match} แถว\n"
        f"⚠️ *ยอดเงินไม่ตรงกัน:* {n_mismatch} แถว\n"
        f"🔍 *ไม่พบใน Master:* {n_not_found} แถว\n"
        f"📋 *รวมรายการที่ตรวจเช็ค:* {n_total_report} แถว\n\n"
        f"📌 *บิล Master ที่ไม่มีใน Report:* {n_not_in_rep} แถว\n"
        f"────────────────────────\n"
        f"📎 *แนบไฟล์รายงานฉบับสมบูรณ์ `{output_excel}` มาด้วยแล้วครับ*"
    )

    print("\n" + summary_msg + "\n")

    # 8. ส่งผลลัพธ์เข้า Telegram
    send_telegram(summary_msg, output_excel, chat_id=target_chat)
    logger.info("=== เสร็จสิ้นกระบวนการทำงานเรียบร้อย ===")
    return result


def start_telegram_listener():
    """
    โหมดรันเป็น Bot Listener ตลอดเวลา:
    คอยรับข้อความใน Telegram เมื่อผู้ใช้พิมพ์คำว่า "Match Data" (หรือ /match)
    จะสั่งรันประมวลผลและตอบกลับรายงานพร้อมไฟล์ Excel ทันที
    """
    logger.info("🤖 บอท Telegram กำลังออนไลน์และพร้อมรับคำสั่ง 'Match Data'...")
    offset = None

    # ส่งข้อความแจ้งผู้ใช้ว่าบอทออนไลน์แล้ว
    send_telegram(
        "🤖 *บอท Ploy(Accounting) ออนไลน์แล้วครับ!*\n\n"
        "พิมพ์คำว่า *Match Data* เพื่อสั่งให้ตรวจสอบและจับคู่ข้อมูลยอดเงินทันที",
        chat_id=TELEGRAM_CHAT_ID
    )

    while True:
        try:
            params = {"timeout": 25}
            if offset is not None:
                params["offset"] = offset

            url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates"
            res = requests.get(url, params=params, timeout=30)
            if res.status_code != 200:
                time.sleep(3)
                continue

            data = res.json()
            if not data.get("ok"):
                time.sleep(3)
                continue

            for update in data.get("result", []):
                update_id = update["update_id"]
                offset = update_id + 1

                msg = update.get("message")
                if not msg:
                    continue

                chat = msg.get("chat", {})
                chat_id = chat.get("id")
                text = str(msg.get("text", "")).strip()

                logger.info(f"ได้รับข้อความจาก Chat ID {chat_id}: '{text}'")

                text_lower = text.lower()
                # ตรวจจับคำสั่ง Match Data
                if text_lower in ("match data", "matchdata", "match", "/match"):
                    send_telegram(
                        "⏳ *ได้รับคำสั่ง Match Data เรียบร้อยแล้ว!*\n"
                        "กำลังดาวน์โหลดไฟล์จาก Google Drive และประมวลผลข้อมูล กรุณารอสักครู่...",
                        chat_id=chat_id
                    )
                    try:
                        run_matching_job(chat_id=chat_id)
                    except Exception as err:
                        logger.error(f"เกิดข้อผิดพลาดในการประมวลผล: {err}", exc_info=True)
                        send_telegram(f"❌ *เกิดข้อผิดพลาดในการประมวลผล:*\n`{str(err)}`", chat_id=chat_id)

                elif text_lower in ("/start", "/help"):
                    send_telegram(
                        "👋 *สวัสดีครับ! ยินดีต้อนรับสู่ระบบตรวจยอดบิล*\n\n"
                        "💡 *คำสั่งที่รองรับ:*\n"
                        "• พิมพ์คำว่า *Match Data* เพื่อเริ่มการตรวจสอบข้อมูลทันที\n"
                        "• หรือกดปุ่มรันผ่าน GitHub Actions ได้ตลอดเวลาครับ",
                        chat_id=chat_id
                    )

        except requests.exceptions.RequestException as re_err:
            logger.warning(f"Connection error: {re_err}, retrying in 5s...")
            time.sleep(5)
        except Exception as e:
            logger.error(f"Listener error: {e}", exc_info=True)
            time.sleep(3)


def main():
    # ตรวจสอบว่าผู้ใช้สั่งให้รันโหมด Listener หรือรันแบบรอบเดียว
    if "--listen" in sys.argv or "-l" in sys.argv:
        start_telegram_listener()
    else:
        run_matching_job()


if __name__ == "__main__":
    main()

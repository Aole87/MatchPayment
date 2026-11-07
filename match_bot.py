import os
import sys
import re
import time
import glob
import logging
import shutil
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
    ค้นหาไฟล์ Report ล่าสุด และไฟล์ Master (เช่น 5-11.xls / 6-11.xls)
    โดยรองรับการเปลี่ยนชื่อไฟล์ทุกวันอัตโนมัติ
    """
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

    report_candidates = [f for f in all_files if "report" in os.path.basename(f).lower()]
    master_candidates = [f for f in all_files if "report" not in os.path.basename(f).lower()]

    if not report_candidates:
        raise FileNotFoundError("ไม่พบไฟล์ประเภท Report_*.xlsx")
    if not master_candidates:
        raise FileNotFoundError("ไม่พบไฟล์ประเภท Master (เช่น 5-11.xls)")

    def sort_key(filepath):
        fname = os.path.basename(filepath)
        mtime = os.path.getmtime(filepath)
        nums = re.findall(r"\d+", fname)
        num_score = int("".join(nums)) if nums else 0
        return (mtime, num_score, fname)

    latest_report = max(report_candidates, key=sort_key)
    latest_master = max(master_candidates, key=sort_key)

    logger.info(f"ไฟล์ Report ที่เลือก: {latest_report}")
    logger.info(f"ไฟล์ Master ที่เลือก: {latest_master}")
    return latest_report, latest_master


def parse_amount(val):
    """แปลงค่ายอดเงินเป็น float ปลอดภัย"""
    if val is None or pd.isna(val):
        return 0.0
    val_str = str(val).replace(",", "").replace(" ", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def read_master_file(filepath):
    """
    อ่านไฟล์ 5-11.xls (หรือไฟล์ Master ประจำวัน)
    แยกตารางค้นหาเป็น 2 กลุ่มอย่างชัดเจน:
    1. กลุ่ม PJ (ใบขายสด/ใบเสร็จรับเงิน)
    2. กลุ่ม V และ *V (ใบกำกับภาษี)
    """
    master_rows = []
    pj_master = {}
    v_master = {}

    def index_master_doc(info):
        master_rows.append(info)
        doc_raw = info["doc_no_raw"]
        u = doc_raw.upper()

        if u.startswith("PJ"):
            clean = u.replace("PJ", "").strip()
            pj_master[doc_raw] = info
            pj_master[u] = info
            pj_master[clean] = info
            m_dig = re.search(r"(\d+)", clean)
            if m_dig:
                d = m_dig.group(1)
                pj_master[d] = info
                pj_master[f"PJ{d}"] = info
        elif u.startswith("V") or u.startswith("*V"):
            clean_star = u.replace("*", "")
            v_master[doc_raw] = info
            v_master[u] = info
            v_master[clean_star] = info
            m_slash = re.search(r"(\d+/\d+)", doc_raw)
            if m_slash:
                v_master[m_slash.group(1)] = info
            m_v_num = re.search(r"V(\d+)", clean_star)
            if m_v_num:
                v_master[f"V{m_v_num.group(1)}"] = info
                v_master[f"*V{m_v_num.group(1)}"] = info

    if filepath.lower().endswith(".xls"):
        try:
            wb = xlrd.open_workbook(filepath, encoding_override="cp874")
        except Exception:
            wb = xlrd.open_workbook(filepath)
        sh = wb.sheet_by_index(0)

        for r in range(sh.nrows):
            if sh.ncols < 2:
                continue

            doc_cell = sh.cell_value(r, 1)  # Col B
            doc_raw = str(doc_cell).strip()
            if not doc_raw or doc_raw == "เลขที่เอกสาร":
                continue

            u = doc_raw.upper()
            if u.startswith("PJ") or u.startswith("V") or u.startswith("*V"):
                cust_name = str(sh.cell_value(r, 3)).strip() if sh.ncols > 3 else ""
                val_before = parse_amount(sh.cell_value(r, 8)) if sh.ncols > 8 else 0.0
                val_net = parse_amount(sh.cell_value(r, 12)) if sh.ncols > 12 else 0.0

                index_master_doc({
                    "master_row": r + 1,
                    "doc_no_raw": doc_raw,
                    "customer_name": cust_name,
                    "val_before": val_before,
                    "val_net": val_net
                })
    else:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        for r in range(1, ws.max_row + 1):
            doc_cell = ws.cell(r, 2).value  # Col B
            doc_raw = str(doc_cell or "").strip()
            if not doc_raw or doc_raw == "เลขที่เอกสาร":
                continue

            u = doc_raw.upper()
            if u.startswith("PJ") or u.startswith("V") or u.startswith("*V"):
                cust_name = str(ws.cell(r, 4).value or "").strip()
                val_before = parse_amount(ws.cell(r, 9).value)
                val_net = parse_amount(ws.cell(r, 13).value)

                index_master_doc({
                    "master_row": r,
                    "doc_no_raw": doc_raw,
                    "customer_name": cust_name,
                    "val_before": val_before,
                    "val_net": val_net
                })

    logger.info(f"อ่านไฟล์ Master สำเร็จ: พบเอกสาร PJ ({len(pj_master)} คีย์), V/*V ({len(v_master)} คีย์) รวม {len(master_rows)} รายการ")
    return master_rows, pj_master, v_master


def read_report_file(filepath):
    """
    อ่านไฟล์ Report_*.xlsx
    สกัด: วันเวลา, รหัสรายการ, ยอดเงิน (Col D), หมายเหตุ (Col I)
    และกรองเฉพาะแถวที่มีการระบุเลขที่บิลขายจริง
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

        type_str = str(type_val or "").strip()
        note_str = str(note_val or "").strip()
        code_str = str(code_val or "").strip()

        # กรองรายการปิดสิ้นวัน/เติมเงินทอน/นำเงินออก ที่ไม่ใช่บิลขายและไม่มีหมายเหตุบิล
        if type_str in ("ปิดสิ้นวัน", "เติมเงินทอน", "ถอนเงิน", "นำเงินออกระหว่างวัน") and not note_str:
            continue

        # ถ้าไม่มีหมายเหตุ และรหัสรายการไม่ได้ขึ้นต้นด้วย PJ, V, *V ให้ข้ามไป
        u_code = code_str.upper()
        if not note_str and not any(u_code.startswith(p) for p in ("PJ", "V", "*V")):
            continue

        # ตรวจสอบว่าใน note หรือ code มีตัวเลขบิลหรือไม่
        if not re.search(r"\d{4,}", note_str) and not any(u_code.startswith(p) for p in ("PJ", "V", "*V")):
            continue

        rows_data.append({
            "report_row": r,
            "datetime": str(time_val or "").strip(),
            "tx_code": code_str,
            "tx_type": type_str,
            "amount": parse_amount(amount_val),
            "status": str(status_val or "").strip(),
            "note_raw": note_str
        })

    logger.info(f"อ่านไฟล์ Report สำเร็จ: พบรายการที่มีเลขที่บิล {len(rows_data)} แถว")
    return rows_data


def process_matching(report_rows, master_rows, pj_master, v_master):
    """
    ทำการ Match ข้อมูลระหว่าง Report และ Master:
    แยกประเภทตามแบบที่ 1:
    - ถ้าใน Report มีเครื่องหมาย '/' หรือมีตัวอักษร 'V', '*V' -> จับคู่เฉพาะกลุ่ม V และ *V
    - ถ้าเป็นตัวเลขล้วน ไม่มี '/' หรือมี 'PJ' -> จับคู่เฉพาะกลุ่ม PJ เท่านั้น (ไม่จับคู่ข้ามกลุ่ม)
    - ตรวจสอบยอดเงิน (Report ยอดเงิน Col D == Master มูลค่าสุทธิ Col M)
    """
    matched_list = []
    mismatch_list = []
    not_found_report = []

    matched_master_docs = set()

    for item in report_rows:
        note = item["note_raw"]
        code = item["tx_code"]
        amt = item["amount"]
        r_num = item["report_row"]
        u_code = code.upper()

        # แยกประเภท: ถ้ามีเครื่องหมาย '/' หรือ 'v', '*v' ให้ค้นหาเฉพาะในกลุ่ม V/*V
        is_v_group = ("/" in note) or ("v" in note.lower()) or u_code.startswith("V") or u_code.startswith("*V")
        target_dict = v_master if is_v_group else pj_master

        candidates = []
        if note:
            candidates.append(note)
            candidates.append(note.upper())
            for t in re.split(r"[\s,]+", note):
                t = t.strip()
                if not t:
                    continue
                candidates.append(t)
                candidates.append(t.upper())
                m_dig = re.search(r"(\d+)", t)
                if m_dig:
                    d = m_dig.group(1)
                    candidates.append(d)
                    if is_v_group:
                        candidates.append(f"V{d}")
                        candidates.append(f"*V{d}")
                    else:
                        candidates.append(f"PJ{d}")
        if code:
            candidates.append(code)

        m_info = None
        for cand in candidates:
            if cand in target_dict:
                m_info = target_dict[cand]
                break

        if m_info:
            doc_key = m_info["doc_no_raw"]
            matched_master_docs.add(doc_key)
            diff = round(amt - m_info["val_net"], 2)

            detail = {
                "Report แถว": r_num,
                "วันเวลาใน Report": item["datetime"],
                "เลขที่บิลที่สกัดได้": doc_key,
                "หมายเหตุ Report": note,
                "ยอดเงิน Report (Col D)": amt,
                "Master แถว": m_info["master_row"],
                "เลขที่เอกสาร Master (Col B)": doc_key,
                "ชื่อลูกค้า Master (Col D)": m_info["customer_name"],
                "มูลค่าสุทธิ Master (Col M)": m_info["val_net"],
                "ผลต่าง (Report - Master)": diff,
            }

            if abs(diff) < 0.01:
                detail["สถานะ"] = "ตรงกัน"
                matched_list.append(detail)
            else:
                detail["สถานะ"] = "ยอดเงินไม่ตรงกัน"
                detail["หมายเหตุปัญหา"] = f"ยอดเงินใน Report ({amt:,.2f}) ไม่ตรงกับ Master ({m_info['val_net']:,.2f})"
                mismatch_list.append(detail)
        else:
            if re.search(r"\d{4,}", note):
                disp_doc = note.split()[0]
                if not is_v_group and not disp_doc.upper().startswith("PJ"):
                    disp_doc = f"PJ{disp_doc}"
                not_found_report.append({
                    "Report แถว": r_num,
                    "วันเวลาใน Report": item["datetime"],
                    "เลขที่บิลที่สกัดได้": disp_doc,
                    "หมายเหตุ Report (ดิบ)": note,
                    "ยอดเงิน Report (Col D)": amt,
                    "สถานะ": "ไม่พบใน Master",
                    "หมายเหตุปัญหา": f"ไม่พบรหัส {disp_doc} ในไฟล์ Master"
                })

    not_in_report = []
    for m_info in master_rows:
        doc_no = m_info["doc_no_raw"]
        if doc_no not in matched_master_docs:
            not_in_report.append({
                "Master แถว": m_info["master_row"],
                "เลขที่เอกสาร Master (Col B)": doc_no,
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
        "total_master_pj": len(master_rows)
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

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(1, col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

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
    """สร้างไฟล์ Excel สรุปผล ที่มีหลายชีต ครอบคลุมทั้งสรุปผลและรายละเอียด"""
    wb = openpyxl.Workbook()

    # --- Sheet 1: สรุปภาพรวม ---
    ws_summary = wb.active
    ws_summary.title = "สรุปภาพรวม"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:D1")
    title_cell = ws_summary["A1"]
    title_cell.value = "📊 รายงานสรุปผลการ Match ข้อมูลยอดชำระเงิน (PJ, V, *V)"
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 35

    summary_rows = [
        ("วันที่ประมวลผล", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("ไฟล์ Report (รายการรับชำระเงิน)", os.path.basename(report_name)),
        ("ไฟล์ Master (ยอดขาย/ใบแจ้งหนี้ PJ, V, *V)", os.path.basename(master_name)),
        ("──────────────────────────────", "──────────────────"),
        ("1. ข้อมูลที่ตรงกันสมบูรณ์ (Matched)", len(result["matched"])),
        ("2. ข้อมูลที่ยอดเงินไม่ตรงกัน (Mismatch)", len(result["mismatched"])),
        ("3. ข้อมูลใน Report ที่ไม่พบใน Master", len(result["not_found_report"])),
        ("รวมรายการ Report ที่ตรวจเช็คทั้งหมด", result["total_report_pj"]),
        ("──────────────────────────────", "──────────────────"),
        ("4. บิลใน Master ที่ยังไม่มีใน Report", len(result["not_in_report"])),
        ("รวมบิล Master (PJ, V, *V) ทั้งหมด", result["total_master_pj"]),
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

    ws_summary.column_dimensions["A"].width = 45
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

    if file_path and os.path.exists(file_path):
        doc_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
        try:
            with open(file_path, "rb") as f:
                res_doc = requests.post(
                    doc_url,
                    data={
                        "chat_id": target_chat,
                        "caption": f"📄 รายงานฉบับเต็ม: {os.path.basename(file_path)}"
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
            f"ℹ️ ไฟล์ '{filename}' ถูกบันทึกไว้ในระบบเรียบร้อยแล้ว\n"
            f"(หากต้องการให้อัปโหลดเข้า Google Drive อัตโนมัติ สามารถเพิ่ม GDRIVE_UPLOAD_URL ใน Secrets)"
        )

    return False


def run_matching_job(report_file=None, master_file=None, chat_id=None):
    """ฟังก์ชันประมวลผลการ Match ข้อมูลหลัก (รองรับ PJ, V, *V)"""
    logger.info("=== เริ่มต้นกระบวนการ Match ข้อมูล (PJ, V, *V) ===")
    target_chat = chat_id or TELEGRAM_CHAT_ID

    # 1. ดาวน์โหลดไฟล์จาก Google Drive
    download_dir = download_drive_folder()

    # 2. ค้นหาไฟล์ Report และไฟล์ Master
    if not report_file or not master_file:
        report_file, master_file = find_target_files([download_dir, "."])

    # 3. อ่านและเตรียมข้อมูล
    report_rows = read_report_file(report_file)
    master_rows, pj_master, v_master = read_master_file(master_file)

    # 4. ทำการ Match ข้อมูล (PJ, V, *V ตามแบบที่ 1)
    result = process_matching(report_rows, master_rows, pj_master, v_master)

    # 5. สร้างไฟล์ Excel สรุปผล
    date_str = datetime.now().strftime("%Y%m%d")
    output_excel = f"match_result_{date_str}.xlsx"
    create_result_excel(result, report_file, master_file, output_excel)

    shutil.copyfile(output_excel, "match_result.xlsx")

    # 6. อัปโหลดเข้า Google Drive (ถ้ามีการตั้งค่า)
    upload_to_google_drive(output_excel)

    # 7. เตรียมข้อความสรุป
    n_match = len(result["matched"])
    n_mismatch = len(result["mismatched"])
    n_not_found = len(result["not_found_report"])
    n_total_report = result["total_report_pj"]
    n_not_in_rep = len(result["not_in_report"])

    summary_msg = (
        f"📊 *รายงานสรุปการ Match ข้อมูลประจำวัน (PJ, V, *V)*\n"
        f"📅 ประมวลผลเมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        f"────────────────────────\n"
        f"📁 *Report:* `{os.path.basename(report_file)}`\n"
        f"📁 *Master:* `{os.path.basename(master_file)}`\n\n"
        f"📈 *สรุปผลการตรวจสอบ:*\n"
        f"✅ *ข้อมูลตรงกัน:* {n_match} แถว\n"
        f"⚠️ *ยอดเงินไม่ตรงกัน:* {n_mismatch} แถว\n"
        f"🔍 *ไม่พบใน Master:* {n_not_found} แถว\n"
        f"📋 *รวมรายการที่ตรวจเช็ค:* {n_total_report} แถว\n\n"
        f"📌 *บิล Master (PJ/V/*V) ที่ไม่มีใน Report:* {n_not_in_rep} แถว\n"
        f"────────────────────────\n"
        f"📎 *แนบไฟล์รายงานฉบับสมบูรณ์ `{output_excel}` มาด้วยแล้วครับ*"
    )

    print("\n" + summary_msg + "\n")

    # 8. ส่งผลลัพธ์เข้า Telegram
    send_telegram(summary_msg, output_excel, chat_id=target_chat)
    logger.info("=== เสร็จสิ้นกระบวนการทำงานเรียบร้อย ===")
    return result


def start_telegram_listener():
    """โหมดรันเป็น Bot Listener ตลอดเวลา รองรับคำสั่ง Match Data"""
    logger.info("🤖 บอท Telegram กำลังออนไลน์และพร้อมรับคำสั่ง 'Match Data' (PJ, V, *V)...")
    offset = None

    send_telegram(
        "🤖 *บอท Ploy(Accounting) ออนไลน์แล้วครับ!*\n\n"
        "พิมพ์คำว่า *Match Data* เพื่อสั่งให้ตรวจสอบและจับคู่ข้อมูลยอดเงิน (PJ, V, *V) ทันที",
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
                if text_lower in ("match data", "matchdata", "match", "/match"):
                    send_telegram(
                        "⏳ *ได้รับคำสั่ง Match Data เรียบร้อยแล้ว!*\n"
                        "กำลังดาวน์โหลดไฟล์และประมวลผลข้อมูล (PJ, V, *V) กรุณารอสักครู่...",
                        chat_id=chat_id
                    )
                    try:
                        run_matching_job(chat_id=chat_id)
                    except Exception as err:
                        logger.error(f"เกิดข้อผิดพลาดในการประมวลผล: {err}", exc_info=True)
                        send_telegram(f"❌ *เกิดข้อผิดพลาดในการประมวลผล:*\n`{str(err)}`", chat_id=chat_id)

                elif text_lower in ("/start", "/help"):
                    send_telegram(
                        "👋 *สวัสดีครับ! ยินดีต้อนรับสู่ระบบตรวจยอดบิล (PJ, V, *V)*\n\n"
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
    if "--listen" in sys.argv or "-l" in sys.argv:
        start_telegram_listener()
    else:
        # ตรวจสอบว่ามีการระบุชื่อไฟล์ผ่าน argument หรือไม่
        rep_file = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else None
        mast_file = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("-") else None
        run_matching_job(report_file=rep_file, master_file=mast_file)


if __name__ == "__main__":
    main()

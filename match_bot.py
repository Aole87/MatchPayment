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

# --- โหลดตัวแปรจากไฟล์ .env (ถ้ามี) สำหรับการรันในเครื่องอย่างปลอดภัย ---
def load_dotenv(filepath=".env"):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        k, v = k.strip(), v.strip().strip('"').strip("'")
                        if k and k not in os.environ:
                            os.environ[k] = v
        except Exception:
            pass

load_dotenv()

# --- การตั้งค่าคอนฟิก (ดึงค่าจาก Environment Variables) ---
DEFAULT_DRIVE_FOLDER = "https://drive.google.com/drive/folders/17eSUsgGmTwRexYmyBRSW_OaanuwLYR3g?usp=sharing"
GDRIVE_FOLDER_URL = (os.environ.get("GDRIVE_FOLDER_URL", "").strip()) or DEFAULT_DRIVE_FOLDER

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "7603648423").strip()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def safe_print(msg):
    """พิมพ์ข้อความลง stdout อย่างปลอดภัย ไม่ให้ crash ด้วย Errno 5 เมื่อรันเป็น background daemon"""
    try:
        sys.stdout.write(str(msg) + "\n")
        sys.stdout.flush()
    except (IOError, OSError):
        pass


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


class MissingPairError(Exception):
    """ข้อยกเว้นเมื่อไฟล์ของวันล่าสุดยังมาไม่ครบทั้ง 2 ระบบ"""
    pass


def get_file_data_date(filepath):
    """
    ตรวจสอบวันที่จริงของข้อมูลในไฟล์:
    - ถ้าเป็น Report: อ่านวันที่จากคอลัมน์ A (รองรับทั้งวันเดียว เช่น '05/11' และช่วงวัน เช่น '06-08/11')
    - ถ้าเป็น Master: อ่านวันที่จากชื่อไฟล์ (เช่น 5-11.xls -> '05/11' หรือ 6-8-11.xls -> '06-08/11') หรือจากเนื้อหา
    คืนค่าเป็น (date_str, sort_key) เช่น ('06-08/11', (11, 8))
    """
    fname = os.path.basename(filepath)
    is_report = "report" in fname.lower()

    if is_report:
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            dates = []
            for r in range(2, ws.max_row + 1):
                val = ws.cell(r, 1).value
                if val:
                    m = re.search(r"(\d{1,2})/(\d{1,2})/\d{2,4}", str(val))
                    if m:
                        d, mth = int(m.group(1)), int(m.group(2))
                        dates.append((mth, d))
            if dates:
                dates.sort()
                min_mth, min_d = dates[0]
                max_mth, max_d = dates[-1]
                if min_mth == max_mth and min_d == max_d:
                    return f"{min_d:02d}/{min_mth:02d}", (max_mth, max_d)
                elif min_mth == max_mth:
                    return f"{min_d:02d}-{max_d:02d}/{min_mth:02d}", (max_mth, max_d)
                else:
                    return f"{min_d:02d}/{min_mth:02d}-{max_d:02d}/{max_mth:02d}", (max_mth, max_d)
        except Exception as e:
            logger.warning(f"ไม่สามารถอ่านวันที่จากเซลล์ใน {fname}: {e}")

        # สำรอง: ถอดจากชื่อไฟล์ Report_YYYYMMDD...
        m_fn = re.search(r"Report_(\d{4})(\d{2})(\d{2})", fname, re.IGNORECASE)
        if m_fn:
            d, mth = int(m_fn.group(3)), int(m_fn.group(2))
            return f"{d:02d}/{mth:02d}", (mth, d)

    else:
        # Master: ลองอ่านจากเนื้อหาตาราง Express ก่อน (เช่น แถว 'ระหว่างวันที่ 06/11/2569 ถึง 08/11/2569')
        if filepath.lower().endswith(".xls"):
            try:
                sh = xlrd.open_workbook(filepath, encoding_override="cp874").sheet_by_index(0)
                for r in range(min(sh.nrows, 20)):
                    for c in range(min(sh.ncols, 5)):
                        val = str(sh.cell_value(r, c))
                        m_period = re.search(r"(\d{1,2})/(\d{1,2})/\d{2,4}\s+ถึง\s+(\d{1,2})/(\d{1,2})/\d{2,4}", val)
                        if m_period:
                            d1, m1, d2, m2 = int(m_period.group(1)), int(m_period.group(2)), int(m_period.group(3)), int(m_period.group(4))
                            if d1 == d2 and m1 == m2:
                                return f"{d1:02d}/{m1:02d}", (m2, d2)
                            elif m1 == m2:
                                return f"{d1:02d}-{d2:02d}/{m1:02d}", (m2, d2)
                            else:
                                return f"{d1:02d}/{m1:02d}-{d2:02d}/{m2:02d}", (m2, d2)

                        m2 = re.search(r"(\d{1,2})/(\d{1,2})/\d{2,4}", val)
                        if m2:
                            d, mth = int(m2.group(1)), int(m2.group(2))
                            return f"{d:02d}/{mth:02d}", (mth, d)
            except Exception as e:
                logger.warning(f"ไม่สามารถอ่านวันที่จากเซลล์ใน {fname}: {e}")

        # สำรอง Master: ตรวจสอบช่วงวันที่ในชื่อไฟล์ เช่น 6-8-11.xls
        m_rng = re.search(r"(\d{1,2})-(\d{1,2})-(\d{1,2})", fname)
        if m_rng:
            d1, d2, mth = int(m_rng.group(1)), int(m_rng.group(2)), int(m_rng.group(3))
            return f"{d1:02d}-{d2:02d}/{mth:02d}", (mth, d2)

        # สำรอง Master: ตรวจสอบวันเดี่ยว เช่น 5-11.xls, 6-11.xls
        m_single = re.search(r"(\d{1,2})[-_](\d{1,2})", fname)
        if m_single:
            d, mth = int(m_single.group(1)), int(m_single.group(2))
            return f"{d:02d}/{mth:02d}", (mth, d)

    return None, (0, 0)


def scan_and_group_files(search_dirs=None):
    """
    สแกนโฟลเดอร์และจัดกลุ่มไฟล์แยกตาม 'วันที่ของข้อมูล'
    คืนค่าเป็น:
      date_groups: dict เรียงจากวันล่าสุดไปเก่าสุด
                   {'06/11': {'report': None, 'master': '...', 'sort_key': (11, 6)}, ...}
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

    date_groups = {}
    for fp in all_files:
        fname = os.path.basename(fp)
        is_rep = "report" in fname.lower()
        d_str, s_key = get_file_data_date(fp)
        if not d_str:
            d_str = "ไม่ทราบวันที่"
            s_key = (0, 0)

        if d_str not in date_groups:
            date_groups[d_str] = {"report": None, "master": None, "sort_key": s_key}

        if is_rep:
            curr = date_groups[d_str]["report"]
            if not curr or os.path.getmtime(fp) > os.path.getmtime(curr):
                date_groups[d_str]["report"] = fp
        else:
            curr = date_groups[d_str]["master"]
            if not curr or os.path.getmtime(fp) > os.path.getmtime(curr):
                date_groups[d_str]["master"] = fp

    # เรียงลำดับตามวันที่ล่าสุด (sort_key = (month, day))
    sorted_groups = dict(sorted(date_groups.items(), key=lambda item: item[1]["sort_key"], reverse=True))
    return sorted_groups


def check_files_status(chat_id=None):
    """
    ตรวจสอบสถานะความพร้อมของไฟล์ใน Google Drive
    แล้วส่งรายงานสรุปให้ผู้ใช้ทราบว่าวันไหนพร้อม วันไหนขาดไฟล์อะไร
    """
    download_dir = download_drive_folder()
    date_groups = scan_and_group_files([download_dir, "."])

    if not date_groups:
        msg = "⚠️ *ไม่พบไฟล์ Excel ในระบบเลยครับ* กรุณาอัปโหลดไฟล์เข้า Google Drive"
        send_telegram(msg, chat_id=chat_id)
        return msg

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    lines = [
        "📂 *ผลการตรวจสอบไฟล์ใน Google Drive*",
        f"🕒 ตรวจสอบเมื่อ: {now_str}",
        "────────────────────────"
    ]

    has_ready = False
    for d_str, info in date_groups.items():
        rep = info["report"]
        mast = info["master"]
        lines.append(f"📅 *วันที่ข้อมูล: {d_str}*")

        if mast:
            lines.append(f"  • ฝั่ง Master: `{os.path.basename(mast)}` ✅")
        else:
            lines.append("  • ฝั่ง Master: ❌ *ยังไม่ส่งมา*")

        if rep:
            lines.append(f"  • ฝั่ง Report: `{os.path.basename(rep)}` ✅")
        else:
            lines.append("  • ฝั่ง Report: ❌ *ยังไม่ส่งมา*")

        if mast and rep:
            lines.append("  ➔ *สถานะ:* ✅ *พร้อมจับคู่เอกสาร*")
            has_ready = True
        elif mast and not rep:
            lines.append("  ➔ *สถานะ:* ⏳ *รอไฟล์ Report เพื่อทำการจับคู่*")
        elif rep and not mast:
            lines.append("  ➔ *สถานะ:* ⏳ *รอไฟล์ Master เพื่อทำการจับคู่*")
        lines.append("")

    lines.append("────────────────────────")
    if has_ready:
        lines.append("💡 *พิมพ์คำว่า 'จับคู่เอกสาร' เพื่อเริ่มประมวลผลได้ทันทีครับ*")
    else:
        lines.append("💡 *เมื่ออีกระบบอัปโหลดไฟล์เข้ามาครบแล้ว สามารถพิมพ์ 'จับคู่เอกสาร' ได้ทันทีครับ*")

    full_msg = "\n".join(lines)
    safe_print("\n" + full_msg + "\n")
    send_telegram(full_msg, chat_id=chat_id)
    return full_msg


def find_target_files(search_dirs=None, target_date=None):
    """
    ค้นหาคู่ไฟล์ที่ตรงกันตาม 'วันที่ของข้อมูล' อย่างปลอดภัย:
    1. หากระบุ target_date (เช่น '05/11') -> ดึงคู่วันที่นั้น
    2. หากไม่ระบุ -> ดึงวันที่ล่าสุดที่สมบูรณ์ทั้ง 2 ฝั่ง
       หากวันล่าสุดยังมีไฟล์แค่ระบบเดียว จะแจ้งเตือนให้รออย่างชัดเจน
    """
    date_groups = scan_and_group_files(search_dirs)

    if not date_groups:
        raise FileNotFoundError("ไม่พบไฟล์ Excel ในโฟลเดอร์สำหรับประมวลผล")

    # กรณีระบุวันที่เจาะจง
    if target_date:
        m_t = re.search(r"(\d{1,2})[/-](\d{1,2})", target_date)
        if m_t:
            norm_date = f"{int(m_t.group(1)):02d}/{int(m_t.group(2)):02d}"
        else:
            norm_date = target_date

        if norm_date not in date_groups:
            raise FileNotFoundError(f"ไม่พบข้อมูลไฟล์ของวันที่ {target_date} ในระบบ")

        info = date_groups[norm_date]
        if not info["report"] or not info["master"]:
            missing = "Report" if not info["report"] else "Master"
            raise MissingPairError(f"ไฟล์ของวันที่ {norm_date} ยังไม่ครบ (ขาดไฟล์ {missing})")
        return info["report"], info["master"], norm_date

    # กรณีไม่ระบุวันที่: ตรวจสอบวันล่าสุด
    dates_list = list(date_groups.keys())
    latest_date = dates_list[0]
    latest_info = date_groups[latest_date]

    # ถ้าวันล่าสุดมีครบทั้ง 2 ไฟล์ -> ดำเนินการได้เลย
    if latest_info["report"] and latest_info["master"]:
        logger.info(f"พบคู่ไฟล์สมบูรณ์ของวันล่าสุด {latest_date}: Report={latest_info['report']}, Master={latest_info['master']}")
        return latest_info["report"], latest_info["master"], latest_date

    # ถ้าวันล่าสุดมีแค่ 1 ระบบ -> แจ้งเตือนว่ากำลังรออีกระบบ
    present = "Master" if latest_info["master"] else "Report"
    missing = "Report" if latest_info["master"] else "Master"
    present_file = os.path.basename(latest_info["master"] or latest_info["report"])

    # ตรวจสอบดูว่ามีวันก่อนหน้าที่มีไฟล์ครบหรือไม่
    ready_dates = [d for d, inf in date_groups.items() if inf["report"] and inf["master"]]
    ready_hint = f"\n*(💡 วันที่พร้อมจับคู่ตอนนี้คือ {ready_dates[0]} สามารถพิมพ์ 'จับคู่เอกสาร {ready_dates[0]}' ได้ครับ)*" if ready_dates else ""

    err_msg = (
        f"พบไฟล์ {present} วันที่ {latest_date} (`{present_file}`) แล้ว\n"
        f"แต่ยังขาดไฟล์ {missing} ของวันที่ {latest_date} ครับ\n"
        f"⏳ กรุณารออีกระบบอัปโหลดไฟล์ {missing} แล้วส่งคำสั่ง 'จับคู่เอกสาร' อีกครั้ง{ready_hint}"
    )
    raise MissingPairError(err_msg)


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
    อ่านไฟล์ Master ประจำวัน (เช่น 5-11.xls, 6-11.xls, 8-11.xls)
    รับประกันว่าอ่านครบทุกแถวข้อมูล ไม่ให้มีแถวใดตกหล่น (รวมถึงแถวที่เป็น *PJ เช่น แถว 533):
    1. กลุ่ม PJ และ *PJ: นำมาบันทึกใน master_rows และ pj_master เพื่อจับคู่กับบิลขายสด
    2. กลุ่ม V: นำมาบันทึกใน master_rows และ v_master เพื่อจับคู่กับบิลภาษี
    3. กลุ่ม *V: บันทึกใน master_rows (เพื่อให้รายงานครบทุกแถว) แต่ไม่ใส่ใน v_master เพื่อไม่ให้จับคู่กับยอดเงินตามคำสั่ง
    """
    master_rows = []
    pj_master = {}
    v_master = {}

    def index_master_doc(info):
        master_rows.append(info)
        doc_raw = info["doc_no_raw"]
        u = doc_raw.upper()

        if u.startswith("PJ") or u.startswith("*PJ"):
            clean = u.replace("*", "").replace("PJ", "").strip()
            pj_master[doc_raw] = info
            pj_master[u] = info
            pj_master[clean] = info
            m_dig = re.search(r"(\d+)", clean)
            if m_dig:
                d = m_dig.group(1)
                pj_master[d] = info
                pj_master[f"PJ{d}"] = info
                pj_master[f"*PJ{d}"] = info

        elif u.startswith("V"):
            v_master[doc_raw] = info
            v_master[u] = info
            m_slash = re.search(r"(\d+/\d+)", doc_raw)
            if m_slash:
                v_master[m_slash.group(1)] = info
            m_v_num = re.search(r"V(\d+)", u)
            if m_v_num:
                v_master[f"V{m_v_num.group(1)}"] = info
                v_master[m_v_num.group(1)] = info

        elif u.startswith("*V"):
            # เอกสาร *V อยู่ใน master_rows เพื่อให้ครบทุกแถว แต่ไม่ใส่ใน v_master จึงไม่ถูกนำมา match
            info["is_cancelled_v"] = True

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
            if u.startswith(("PJ", "*PJ", "V", "*V")):
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
            if u.startswith(("PJ", "*PJ", "V", "*V")):
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

    logger.info(f"อ่านไฟล์ Master สำเร็จ: ครบทั้งหมด {len(master_rows)} รายการ (PJ/*PJ: {len(pj_master)} คีย์, V: {len(v_master)} คีย์)")
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

        # ถ้าไม่มีหมายเหตุ และรหัสรายการไม่ได้ขึ้นต้นด้วย PJ, V ให้ข้ามไป
        u_code = code_str.upper()
        if not note_str and not any(u_code.startswith(p) for p in ("PJ", "V")):
            continue

        # ตรวจสอบว่าใน note หรือ code มีตัวเลขบิลหรือไม่
        if not re.search(r"\d{4,}", note_str) and not any(u_code.startswith(p) for p in ("PJ", "V")):
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
    ทำการ Match ข้อมูลระหว่าง Report และ Master (เหลือแค่ PJ กับ V ไม่รวม *V):
    - ถ้าใน Report มีเครื่องหมาย '/' (เช่น 594378/69) หรือมีตัวอักษร V -> จับคู่เฉพาะกลุ่ม V
    - ถ้าใน Report เป็นตัวเลขล้วน ไม่มี '/' (เช่น 594399, 1547564) หรือมี PJ -> จับคู่เฉพาะกลุ่ม PJ เท่านั้น (ไม่จับคู่ข้ามกับ V.../69)
    - เอกสารที่เป็น *V (เอกสารยกเลิกใน Express) ถูกตัดออก ไม่นำมา match
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

        # ข้ามหากมีการระบุเป็น *V ใน Report
        if note.strip().upper().startswith("*V") or u_code.startswith("*V"):
            continue

        # แยกประเภท: ถ้ามีเครื่องหมาย '/' หรือ 'v' ให้ค้นหาเฉพาะในกลุ่ม V
        is_v_group = ("/" in note) or ("v" in note.lower()) or u_code.startswith("V")
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
            st = "ไม่มีใน Report (*V เอกสารยกเลิก - ไม่นำมา match)" if m_info.get("is_cancelled_v") else "ไม่มีใน Report"
            not_in_report.append({
                "Master แถว": m_info["master_row"],
                "เลขที่เอกสาร Master (Col B)": doc_no,
                "ชื่อลูกค้า Master (Col D)": m_info["customer_name"],
                "มูลค่าสุทธิ Master (Col M)": m_info["val_net"],
                "สถานะ": st
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
    title_cell.value = "📊 รายงานสรุปผลการจับคู่ข้อมูลยอดชำระเงิน (น้องบิ๊นท์)"
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 35

    summary_rows = [
        ("วันที่ประมวลผล", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("ผู้จัดทำรายงาน", "น้องบิ๊นท์ (Accounting Bot)"),
        ("ไฟล์ Report (รายการรับชำระเงิน)", os.path.basename(report_name)),
        ("ไฟล์ Master (ยอดขาย/ใบแจ้งหนี้ PJ, V)", os.path.basename(master_name)),
        ("──────────────────────────────", "──────────────────"),
        ("1. ข้อมูลที่ตรงกันสมบูรณ์ (Matched)", len(result["matched"])),
        ("2. ข้อมูลที่ยอดเงินไม่ตรงกัน (Mismatch)", len(result["mismatched"])),
        ("3. ข้อมูลใน Report ที่ไม่พบใน Master", len(result["not_found_report"])),
        ("รวมรายการ Report ที่ตรวจเช็คทั้งหมด", result["total_report_pj"]),
        ("──────────────────────────────", "──────────────────"),
        ("4. บิลใน Master ที่ยังไม่มีใน Report", len(result["not_in_report"])),
        ("รวมบิล Master ทั้งหมด (ครบทุกแถว)", result["total_master_pj"]),
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


def run_matching_job(report_file=None, master_file=None, chat_id=None, target_date=None):
    """ฟังก์ชันประมวลผลการ Match ข้อมูลหลัก (รองรับเฉพาะ PJ และ V ตามกฎใหม่ ไม่รวม *V)"""
    logger.info("=== เริ่มต้นกระบวนการ Match ข้อมูล (PJ, V) ===")
    target_chat = chat_id or TELEGRAM_CHAT_ID

    # 1. ดาวน์โหลดไฟล์จาก Google Drive
    download_dir = download_drive_folder()

    # 2. ค้นหาไฟล์ Report และไฟล์ Master ตามวันที่ของข้อมูลจริง
    matched_date_str = None
    if not report_file or not master_file:
        try:
            report_file, master_file, matched_date_str = find_target_files([download_dir, "."], target_date=target_date)
        except MissingPairError as me:
            logger.warning(f"ไฟล์ยังไม่พร้อมจับคู่: {me}")
            warn_msg = (
                f"⚠️ *ยังไม่สามารถจับคู่เอกสารได้*\n\n"
                f"{str(me)}\n\n"
                f"💡 *สามารถพิมพ์ 'ตรวจสอบไฟล์' เพื่อดูสถานะไฟล์ทั้งหมดใน Google Drive ได้ครับ*"
            )
            send_telegram(warn_msg, chat_id=target_chat)
            return None
        except Exception as fe:
            logger.error(f"ไม่สามารถค้นหาคู่ไฟล์ได้: {fe}")
            send_telegram(f"❌ *ไม่สามารถค้นหาคู่ไฟล์ได้:*\n`{str(fe)}`", chat_id=target_chat)
            return None

    # 3. อ่านและเตรียมข้อมูล
    report_rows = read_report_file(report_file)
    master_rows, pj_master, v_master = read_master_file(master_file)

    # 4. ทำการ Match ข้อมูล (เฉพาะ PJ และ V)
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

    date_tag = f" (ข้อมูลวันที่ {matched_date_str})" if matched_date_str else ""
    summary_msg = (
        f"📊 *รายงานสรุปการจับคู่เอกสารโดย น้องบิ๊นท์ (PJ, V)*{date_tag}\n"
        f"📅 ประมวลผลเมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        f"────────────────────────\n"
        f"📁 *Report:* `{os.path.basename(report_file)}`\n"
        f"📁 *Master:* `{os.path.basename(master_file)}`\n\n"
        f"📈 *สรุปผลการตรวจสอบ:*\n"
        f"✅ *ข้อมูลตรงกัน:* {n_match} แถว\n"
        f"⚠️ *ยอดเงินไม่ตรงกัน:* {n_mismatch} แถว\n"
        f"🔍 *ไม่พบใน Master:* {n_not_found} แถว\n"
        f"📋 *รวมรายการที่ตรวจเช็ค:* {n_total_report} แถว\n\n"
        f"📌 *บิล Master (PJ/V) ที่ไม่มีใน Report:* {n_not_in_rep} แถว\n"
        f"────────────────────────\n"
        f"📎 *แนบไฟล์รายงานฉบับสมบูรณ์ `{output_excel}` มาด้วยแล้วค่ะ*"
    )

    safe_print("\n" + summary_msg + "\n")

    # 8. ส่งผลลัพธ์เข้า Telegram
    send_telegram(summary_msg, output_excel, chat_id=target_chat)
    logger.info("=== เสร็จสิ้นกระบวนการทำงานเรียบร้อย ===")
    return result


def start_telegram_listener():
    """โหมดรันเป็น Bot Listener ตลอดเวลา รองรับคำสั่ง 'ตรวจสอบไฟล์' และ 'จับคู่เอกสาร'"""
    logger.info("🤖 น้องบิ๊นท์ กำลังออนไลน์และพร้อมรับคำสั่ง...")
    offset = None

    send_telegram(
        "🤖 *สวัสดีค่ะ น้องบิ๊นท์ พร้อมช่วยตรวจสอบยอดบิลแล้วค่ะ!*\n\n"
        "💡 *คำสั่งภาษาไทยที่รองรับ:*\n"
        "• *ตรวจสอบไฟล์* ➔ ให้น้องบิ๊นท์เช็คดูว่าใน Google Drive มีไฟล์วันไหนบ้าง และระบบไหนส่งมาแล้วบ้าง\n"
        "• *จับคู่เอกสาร* ➔ สั่งให้น้องบิ๊นท์จับคู่และส่งรายงาน Excel เข้า Telegram ทันที\n"
        "• *จับคู่เอกสาร [วันที่]* ➔ เจาะจงจับคู่ของวันที่ต้องการ (เช่น `จับคู่เอกสาร 05/11`)",
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
                text_clean = text.lower().strip()

                # คำสั่ง: ตรวจสอบไฟล์ (status, check)
                if any(k in text_clean for k in ("ตรวจสอบไฟล์", "ตรวจไฟล์", "เช็คไฟล์", "status", "check")):
                    send_telegram("⏳ *น้องบิ๊นท์กำลังตรวจสอบไฟล์ใน Google Drive กรุณารอสักครู่นะคะ...*", chat_id=chat_id)
                    try:
                        check_files_status(chat_id=chat_id)
                    except Exception as err:
                        logger.error(f"เกิดข้อผิดพลาดในการตรวจสอบไฟล์: {err}", exc_info=True)
                        send_telegram(f"❌ *เกิดข้อผิดพลาดในการตรวจสอบไฟล์:*\n`{str(err)}`", chat_id=chat_id)

                # คำสั่ง: จับคู่เอกสาร (match data, match)
                elif any(k in text_clean for k in ("จับคู่เอกสาร", "จับคู่ข้อมูล", "จับคู่", "match data", "matchdata", "match", "/match")):
                    # ตรวจสอบว่ามีการระบุวันที่หรือไม่ เช่น "จับคู่เอกสาร 05/11" หรือ "จับคู่ 5-11"
                    m_d = re.search(r"(\d{1,2})[/-](\d{1,2})", text)
                    req_date = f"{int(m_d.group(1)):02d}/{int(m_d.group(2)):02d}" if m_d else None

                    date_text = f" ของวันที่ {req_date}" if req_date else ""
                    send_telegram(
                        f"⏳ *น้องบิ๊นท์ได้รับคำสั่ง จับคู่เอกสาร{date_text} เรียบร้อยแล้วค่ะ!*\n"
                        "กำลังดาวน์โหลดไฟล์และตรวจสอบความพร้อม กรุณารอสักครู่นะคะ...",
                        chat_id=chat_id
                    )
                    try:
                        run_matching_job(chat_id=chat_id, target_date=req_date)
                    except Exception as err:
                        logger.error(f"เกิดข้อผิดพลาดในการประมวลผล: {err}", exc_info=True)
                        send_telegram(f"❌ *เกิดข้อผิดพลาดในการประมวลผล:*\n`{str(err)}`", chat_id=chat_id)

                elif text_clean in ("/start", "/help", "help", "คำสั่ง"):
                    send_telegram(
                        "👋 *สวัสดีค่ะ! น้องบิ๊นท์ ยินดีต้อนรับสู่ระบบตรวจยอดบิล (PJ, V)*\n\n"
                        "💡 *คำสั่งภาษาไทยที่รองรับ:*\n"
                        "• *ตรวจสอบไฟล์* ➔ ให้น้องบิ๊นท์ตรวจเช็คว่าใน Google Drive มีไฟล์วันไหนเข้ามาบ้าง และคู่ไหนพร้อมจับคู่แล้ว\n"
                        "• *จับคู่เอกสาร* ➔ เริ่มการตรวจสอบและจับคู่ข้อมูลยอดขายอัตโนมัติ (เฉพาะบิล PJ และ V)\n"
                        "• *จับคู่เอกสาร [วันที่]* (เช่น `จับคู่เอกสาร 05/11`) ➔ ระบุวันที่ที่ต้องการจับคู่ได้ค่ะ",
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
    elif "--check" in sys.argv or "--status" in sys.argv or "-c" in sys.argv:
        check_files_status()
    else:
        # ตรวจสอบว่ามีการระบุชื่อไฟล์ผ่าน argument หรือไม่
        rep_file = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") and os.path.exists(sys.argv[1]) else None
        mast_file = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("-") and os.path.exists(sys.argv[2]) else None
        target_d = None
        for arg in sys.argv[1:]:
            m_d = re.search(r"^(\d{1,2})[/-](\d{1,2})$", arg)
            if m_d:
                target_d = f"{int(m_d.group(1)):02d}/{int(m_d.group(2)):02d}"
                break
        run_matching_job(report_file=rep_file, master_file=mast_file, target_date=target_d)


if __name__ == "__main__":
    main()

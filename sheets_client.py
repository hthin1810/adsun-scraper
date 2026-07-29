"""
Ket noi Google Sheets qua Service Account (gspread + google-auth).
Can file google_credentials.json (tai ve tu Google Cloud Console) va sheet_id
cua Google Sheet da duoc share quyen Editor cho email service account.
"""

import json
import re
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

CONFIG_PATH = Path(__file__).parent / "config.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ZONES_TAB = "Zones"
ZONES_HEADER = ["name", "points_json", "created_at"]

PAIR_RULES_TAB = "PairRules"
PAIR_RULES_HEADER = ["vung_a", "vung_b", "created_at"]


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_client():
    config = load_config()
    creds_path = Path(config["credentials_path"])
    if not creds_path.is_absolute():
        # Duong dan tuong doi (vd khi trien khai len server khac, Linux) —
        # tinh theo vi tri file config.json, khong phu thuoc thu muc dang chay.
        creds_path = CONFIG_PATH.parent / creds_path
    if not creds_path.exists():
        raise RuntimeError(
            f"Không tìm thấy file credentials tại {creds_path}. "
            "Cần tải file service account JSON từ Google Cloud Console."
        )
    creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet():
    config = load_config()
    sheet_id = config.get("sheet_id")
    if not sheet_id:
        raise RuntimeError("Chưa cấu hình sheet_id trong config.json.")
    client = get_client()
    return client.open_by_key(sheet_id)


def get_report_spreadsheet():
    """Spreadsheet RIÊNG chỉ chứa báo cáo hành trình (6 cột: STT, Thời điểm
    A/B, Biển số xe, Vùng A/B) — tách khỏi spreadsheet chứa tab Zones."""
    config = load_config()
    sheet_id = config.get("report_sheet_id")
    if not sheet_id:
        raise RuntimeError("Chưa cấu hình report_sheet_id trong config.json.")
    client = get_client()
    return client.open_by_key(sheet_id)


def create_report_spreadsheet(title, share_email):
    """Tạo mới 1 spreadsheet (chạy 1 lần) rồi share quyền Editor cho email
    người dùng — service account tạo file sẽ là chủ sở hữu mặc định, cần
    share thì tài khoản Google thật của người dùng mới mở/sửa được."""
    client = get_client()
    ss = client.create(title)
    ss.share(share_email, perm_type="user", role="writer")
    return ss.id


def get_or_create_worksheet(spreadsheet, title, header=None, rows=1000, cols=20):
    try:
        ws = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
        if header:
            ws.append_row(header)
    return ws


def list_zones():
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    records = ws.get_all_records()
    zones = []
    for r in records:
        try:
            points = json.loads(r["points_json"])
        except Exception:
            continue
        zones.append({"name": r["name"], "points": points, "created_at": r.get("created_at", "")})
    return zones


def add_zone(name, points, created_at):
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    ws.append_row([name, json.dumps(points, ensure_ascii=False), created_at])


def delete_zone(name):
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    cell = ws.find(name, in_column=1)
    if cell:
        ws.delete_rows(cell.row)
        return True
    return False


def _autofit_column_requests(ws, dataframe, char_px=9, padding_px=28, min_px=70, max_px=420):
    """Tinh do rong cot THU CONG theo do dai ky tu dai nhat (ke ca header),
    thay vi dung autoResizeDimensions cua Google — API do doi khi uoc luong
    thieu vai px voi chuoi co dau tieng Viet, gay cat chu (vd "LÁI" mat chu
    "I" cuoi). Tinh tay + padding du rong tranh loi nay."""
    requests = []
    for idx, col in enumerate(dataframe.columns):
        max_len = len(str(col))
        if len(dataframe) > 0:
            max_len = max(max_len, dataframe[col].astype(str).map(len).max())
        width = min(max(max_len * char_px + padding_px, min_px), max_px)
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": idx,
                        "endIndex": idx + 1,
                    },
                    "properties": {"pixelSize": int(width)},
                    "fields": "pixelSize",
                }
            }
        )
    return requests


def list_pair_rules():
    """Danh sach cap (Vung A, Vung B) do nguoi dung tu dinh nghia — hieu la:
    heo gap Vung A ngay truoc Vung B trong du lieu, LUON tao 1 dong chuyen
    vung (A->B), bat ke vai tro mac dinh cua 2 vung do la gi (xem
    adsun_daily_report.build_transition_rows)."""
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, PAIR_RULES_TAB, header=PAIR_RULES_HEADER)
    records = ws.get_all_records()
    rules = []
    for r in records:
        a = str(r.get("vung_a", "")).strip()
        b = str(r.get("vung_b", "")).strip()
        if a and b:
            rules.append((a, b))
    return rules


def add_pair_rule(vung_a, vung_b):
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, PAIR_RULES_TAB, header=PAIR_RULES_HEADER)
    ws.append_row([vung_a, vung_b, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])


def delete_pair_rule(vung_a, vung_b):
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, PAIR_RULES_TAB, header=PAIR_RULES_HEADER)
    records = ws.get_all_records()
    for idx, r in enumerate(records):
        if str(r.get("vung_a", "")).strip() == vung_a and str(r.get("vung_b", "")).strip() == vung_b:
            ws.delete_rows(idx + 2)  # +2: header row + 1-indexed
            return True
    return False


def _write_full_tab(ss, ws, dataframe):
    ws.clear()
    values = [list(dataframe.columns)] + dataframe.astype(str).values.tolist()
    ws.update(values)
    ss.batch_update({"requests": _autofit_column_requests(ws, dataframe)})


def upsert_range_report(
    tab_name,
    range_start,
    range_end,
    new_dataframe,
    time_column="Thời điểm A",
    plate_column="Biển số xe",
    only_replace_plates=None,
):
    """Cap nhat 1 tab (thuong dat ten theo thang, vd '2026-07') bang cach CHI
    THAY du lieu co time_column nam trong [range_start, range_end) — giu
    nguyen toan bo du lieu cac ngay/gio khac da co san trong tab do. Phu hop
    de chay lap lai (moi gio) ma khong can quet/ghi lai toan bo thang.

    only_replace_plates (tuy chon, mot tap bien so xe): neu duoc truyen, CHI
    thay du lieu cua NHUNG XE DO trong khoang thoi gian — du lieu cua cac xe
    KHAC (vd 1 xe bi loi khi quet, khong nam trong tap nay) van duoc GIU
    NGUYEN du cung nam trong khoang [range_start, range_end). Bat buoc phai
    dung tham so nay bat cu khi nao co kha nang 1 vai xe bi loi/bo qua giua
    chung — neu khong, du lieu cu cua nhung xe do se bi XOA MAT thay vi giu
    nguyen (day chinh la nguyen nhan gay mat du lieu that su xay ra ngay
    27/7 — 1 xe loi khi quet backfill ca thang lam xoa mat du lieu ca thang
    cua xe do).

    Viec doc + ghi lai ca tab la 1 API call re (vai tram/nghin dong van nhanh);
    phan TON THOI GIAN THAT (quet Adsun bang trinh duyet) chi gioi han trong
    khoang [range_start, range_end) truyen vao, khong lien quan kich thuoc
    ca thang.
    """
    ss = get_report_spreadsheet()
    try:
        ws = ss.worksheet(tab_name)
        existing_records = ws.get_all_records()
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(
            title=tab_name,
            rows=max(len(new_dataframe) + 10, 100),
            cols=max(len(new_dataframe.columns) + 2, 10),
        )
        existing_records = []

    kept_rows = []
    for r in existing_records:
        try:
            t = pd.to_datetime(r[time_column])
        except Exception:
            continue
        in_range = range_start <= t < range_end
        if not in_range:
            kept_rows.append(r)
        elif only_replace_plates is not None and r.get(plate_column) not in only_replace_plates:
            # Xe nay KHONG nam trong danh sach duoc phep thay lan nay (vd bi
            # loi khi quet) — giu nguyen dong cu, khong xoa.
            kept_rows.append(r)

    if kept_rows:
        existing_df = pd.DataFrame(kept_rows)
        existing_df[time_column] = pd.to_datetime(existing_df[time_column])
        combined = pd.concat([existing_df, new_dataframe], ignore_index=True)
    else:
        combined = new_dataframe.copy()

    combined = combined.sort_values(time_column).reset_index(drop=True)
    if "STT" in combined.columns:
        combined["STT"] = range(1, len(combined) + 1)

    _write_full_tab(ss, ws, combined)
    return ws.id


REFRESH_TAB = "RefreshRequest"
REFRESH_HEADER = ["requested_at", "status", "completed_at", "range_start", "range_end", "mode"]


def request_refresh(range_start=None, range_end=None, mode="scrape"):
    """Ghi 1 yeu cau 'Cap nhat ngay' — bat ke nguoi dung dang mo web qua link
    local (Cloudflare Tunnel) hay link PythonAnywhere, ca 2 deu ghi vao CHUNG
    1 Google Sheet nay. Tien trinh theo doi tren may local (refresh_watcher.py,
    chay moi 1 phut qua Task Scheduler) se phat hien va kich hoat quet ngay,
    vi viec quet Adsun (Playwright) chi chay duoc tren may local co trinh
    duyet, khong chay duoc tren PythonAnywhere.

    range_start/range_end (chuoi "YYYY-MM-DD HH:MM", tuy chon): neu co, chi
    lam moi DUNG khoang thoi gian do (nhanh hon, giup sua dung cho nhung
    khoang bi loi/thieu ma khong can quet lai ca ngay/ca thang). Neu bo trong,
    mac dinh lam moi theo cua so cap nhat thong thuong (xem adsun_daily_report.py).

    mode: "scrape" (mac dinh) — quet du lieu MOI tu Adsun. "recompute" — KHONG
    quet Adsun, chi tinh lai Vung A/B cho toan bo du lieu DA CO san trong cache
    bang danh sach vung/cap vung MOI NHAT (dung khi vua them/sua vung tren ban
    do, nhanh hon nhieu vi khong can dang nhap/quet gi ca)."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, REFRESH_TAB, header=REFRESH_HEADER)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.update("A2:F2", [[now, "pending", "", range_start or "", range_end or "", mode or "scrape"]])
    return now


def get_refresh_status():
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, REFRESH_TAB, header=REFRESH_HEADER)
    values = ws.get_all_values()
    if len(values) < 2:
        return {
            "requested_at": None,
            "status": None,
            "completed_at": None,
            "range_start": None,
            "range_end": None,
            "mode": None,
        }
    row = values[1] + ["", "", "", "", "", ""]
    return {
        "requested_at": row[0] or None,
        "status": row[1] or None,
        "completed_at": row[2] or None,
        "range_start": row[3] or None,
        "range_end": row[4] or None,
        "mode": row[5] or None,
    }


def set_refresh_status(status, completed_at=None):
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, REFRESH_TAB, header=REFRESH_HEADER)
    values = ws.get_all_values()
    requested_at = values[1][0] if len(values) > 1 and len(values[1]) > 0 else ""
    ws.update("A2:C2", [[requested_at, status, completed_at or ""]])


ANNOTATIONS_TAB = "ReportAnnotations"
ANNOTATIONS_HEADER = ["row_key", "ticked", "note", "cell_colors_json", "updated_at"]


def get_annotations():
    """Tra ve dict {row_key: {"ticked": bool, "note": str, "colors": {ten_cot: ma_mau}}}
    — chu thich (tick/ghi chu/mau o) tren bang bao cao, luu chung tren Sheets
    de moi nguoi trong cong ty thay giong nhau."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, ANNOTATIONS_TAB, header=ANNOTATIONS_HEADER)
    records = ws.get_all_records()
    result = {}
    for r in records:
        key = str(r.get("row_key", "")).strip()
        if not key:
            continue
        try:
            colors = json.loads(r.get("cell_colors_json") or "{}")
        except Exception:
            colors = {}
        result[key] = {
            "ticked": str(r.get("ticked", "")).strip().upper() == "TRUE",
            "note": r.get("note", ""),
            "colors": colors,
        }
    return result


def set_annotation(row_key, ticked=None, note=None, colors=None):
    """Cap nhat 1 hoac nhieu truong chu thich cho 1 row_key (tao moi neu chua
    co). `colors` la dict {ten_cot: ma_mau_hoac_None} — gop vao mau da co san,
    None nghia la xoa mau cot do.

    Dung ws.find() de tim DUNG 1 dong can sua thay vi doc toan bo tab (nhu
    truoc) — vi ham nay chay MOI LAN nguoi dung tick/to mau/go ghi chu 1 o
    tren web, doc ca ngan dong chi de sua 1 dong se cham dan khi tab lon len
    theo thoi gian."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, ANNOTATIONS_TAB, header=ANNOTATIONS_HEADER)

    cell = ws.find(row_key, in_column=1)
    current = {"ticked": False, "note": "", "colors": {}}
    if cell:
        existing = ws.row_values(cell.row)
        if len(existing) > 1:
            current["ticked"] = str(existing[1]).strip().upper() == "TRUE"
        if len(existing) > 2:
            current["note"] = existing[2]
        if len(existing) > 3:
            try:
                current["colors"] = json.loads(existing[3] or "{}")
            except Exception:
                current["colors"] = {}

    if ticked is not None:
        current["ticked"] = bool(ticked)
    if note is not None:
        current["note"] = note
    if colors:
        current["colors"].update(colors)
        current["colors"] = {k: v for k, v in current["colors"].items() if v}

    row_values = [
        row_key,
        "TRUE" if current["ticked"] else "FALSE",
        current["note"],
        json.dumps(current["colors"], ensure_ascii=False),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ]

    if cell is None:
        ws.append_row(row_values)
    else:
        ws.update(f"A{cell.row}:E{cell.row}", [row_values])
    return current


MONTH_TAB_RE = re.compile(r"^\d{4}-\d{2}$")


def get_all_report_data():
    """Doc toan bo du lieu tu TAT CA cac tab THANG (dang "YYYY-MM") trong
    spreadsheet bao cao, gop lai thanh 1 danh sach dict, them cot "Tháng"
    (= ten tab) vao moi dong de phan biet nguon. Bo qua cac tab khac nhu
    ReportAnnotations (chu thich tick/ghi chu/mau) hay tab mac dinh trong
    Google Sheets. Dung cho bang tong hop tren web app."""
    ss = get_report_spreadsheet()
    all_rows = []
    for ws in ss.worksheets():
        if not MONTH_TAB_RE.match(ws.title):
            continue
        records = ws.get_all_records()
        for r in records:
            r["Tháng"] = ws.title
            all_rows.append(r)
    return all_rows


def list_report_tabs():
    """Tra ve danh sach (ten_tab, gid) cua CAC TAB THANG (dang "YYYY-MM"),
    doc tu spreadsheet bao cao rieng (report_sheet_id) — bo qua tab
    ReportAnnotations va cac tab khac khong phai du lieu thang."""
    ss = get_report_spreadsheet()
    tabs = []
    for ws in ss.worksheets():
        if not MONTH_TAB_RE.match(ws.title):
            continue
        tabs.append({"name": ws.title, "gid": ws.id})
    tabs.sort(key=lambda t: t["name"], reverse=True)
    return tabs


def spreadsheet_url():
    config = load_config()
    return f"https://docs.google.com/spreadsheets/d/{config['sheet_id']}"


def report_spreadsheet_url():
    config = load_config()
    return f"https://docs.google.com/spreadsheets/d/{config['report_sheet_id']}"

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
ZONES_HEADER = ["name", "points_json", "created_at", "easy_pass"]

PAIR_RULES_TAB = "PairRules"
PAIR_RULES_HEADER = ["vung_a", "vung_b", "created_at"]


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


_client_cache = None
_spreadsheet_cache = {}  # sheet_id -> Spreadsheet, giu lai suot vong doi tien trinh


def get_client():
    global _client_cache
    if _client_cache is not None:
        return _client_cache
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
    _client_cache = gspread.authorize(creds)
    return _client_cache


def _open_by_key_cached(sheet_id):
    """Mo spreadsheet theo id va GIU LAI trong bo nho suot vong doi tien
    trinh, khong mo lai tu dau moi lan goi. Truoc day get_client() +
    open_by_key() bi goi lai tu dau O MOI HAM trong file nay — day chinh la
    nguyen nhan khien moi thao tac (vd luu 1 dong tick/ghi chu) ton 2-3 giay
    du chi doc/ghi 1 dong rat nho, vi phan xac thuc + mo spreadsheet ton
    nhieu thoi gian hon han chinh thao tac do. Voi script chay 1 lan roi
    thoat (backfill_month.py, refresh_range.py...) cache nay van co loi (vi
    ho cung goi nhieu ham trong cung 1 lan chay), khong phat sinh rui ro gi
    them (tab/spreadsheet chi bi doc lai object, khong anh huong du lieu)."""
    if sheet_id not in _spreadsheet_cache:
        client = get_client()
        _spreadsheet_cache[sheet_id] = client.open_by_key(sheet_id)
    return _spreadsheet_cache[sheet_id]


def get_spreadsheet():
    config = load_config()
    sheet_id = config.get("sheet_id")
    if not sheet_id:
        raise RuntimeError("Chưa cấu hình sheet_id trong config.json.")
    return _open_by_key_cached(sheet_id)


def get_report_spreadsheet():
    """Spreadsheet RIÊNG chỉ chứa báo cáo hành trình (6 cột: STT, Thời điểm
    A/B, Biển số xe, Vùng A/B) — tách khỏi spreadsheet chứa tab Zones."""
    config = load_config()
    sheet_id = config.get("report_sheet_id")
    if not sheet_id:
        raise RuntimeError("Chưa cấu hình report_sheet_id trong config.json.")
    return _open_by_key_cached(sheet_id)


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
        zones.append({
            "name": r["name"],
            "points": points,
            "created_at": r.get("created_at", ""),
            "easy_pass": str(r.get("easy_pass", "")).strip().upper() == "TRUE",
        })
    return zones


def add_zone(name, points, created_at, easy_pass=False):
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    ws.append_row([name, json.dumps(points, ensure_ascii=False), created_at, "TRUE" if easy_pass else "FALSE"])


def delete_zone(name):
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    cell = ws.find(name, in_column=1)
    if cell:
        ws.delete_rows(cell.row)
        return True
    return False


def update_zone_points(name, points):
    """Ghi de lai hinh dang (points) cua 1 vung DA CO SAN — dung khi nguoi
    dung keo/sua lai bien dang vung tren ban do (nut sua cua Leaflet.Draw).
    Truoc day tinh nang sua hinh CHUA CO ham nay — chinh sua tren ban do chi
    doi giao dien local, khong bao gio duoc gui len server nen mat het sau
    khi tai lai trang."""
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    cell = ws.find(name, in_column=1)
    if not cell:
        return False
    ws.update_cell(cell.row, 2, json.dumps(points, ensure_ascii=False))
    return True


def set_zone_easy_pass(name, easy_pass):
    """Bat/tat co "vung de di qua" cho 1 vung DA CO SAN (khong can xoa-tao
    lai). Vung de di qua: khi tinh Vung A/B, neu 1 nhom diem lien tiep cung
    vai tro (vd Vung B) co chua vung nay VA CON co it nhat 1 vung KHAC khong
    phai de-di-qua, thuat toan se uu tien chon vung KHAC do lam diem den
    thuc su (xem _pick_group_representative trong adsun_daily_report.py) —
    dung cho truong hop xe phai di ngang qua 1 vung co dinh de toi vung dich
    that su roi quay lai (vd Fico Cat Lai -> Yamaken Cat Lai -> Fico Cat Lai)."""
    ss = get_spreadsheet()
    ws = get_or_create_worksheet(ss, ZONES_TAB, header=ZONES_HEADER)
    cell = ws.find(name, in_column=1)
    if not cell:
        return False
    ws.update_cell(cell.row, 4, "TRUE" if easy_pass else "FALSE")
    return True


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


SCRAPE_LOCK_TAB = "ScrapeLock"
SCRAPE_LOCK_HEADER = ["locked_by", "locked_at"]
SCRAPE_LOCK_STALE_MINUTES = 20  # qua thoi gian nay coi nhu tien trinh cu da chet/treo, cho giao lai khoa


def acquire_scrape_lock(owner_id):
    """Thu 'khoa' quet Adsun TOAN CUC qua Google Sheets — dung CHUNG giua may
    local (AdsunHourlyReport, AdsunRefreshWatcher) VA GitHub Actions (chay
    theo lich rieng + workflow_dispatch tu web app), vi ca 2 phia deu co the
    dang nhap CUNG 1 tai khoan Adsun DOC LAP voi nhau. Da gap thuc te NHIEU
    LAN (vd sang 7/8: 4 lan workflow_dispatch tren GitHub Actions chay chong
    cheo dung luc voi cac lan chay local) khien server Adsun tra nham du lieu
    giua cac phien dang nhap dong thoi — xe MOI duoc chon nhung file xuat ra
    lai la cua 1 xe dang duoc XU LY BOI PHIEN KHAC (xem adsun_common.
    export_vehicle_report — co che fingerprint chi bat duoc loi "dinh du lieu
    xe truoc" TRONG CUNG 1 vong lap, KHONG the phat hien nhiem xuyen tu 1
    tien trinh hoan toan khac dang chay song song).

    Tra ve True neu thu duoc khoa (da ghi ten minh vao), False neu dang co
    tien trinh KHAC giu khoa va con moi (chua qua SCRAPE_LOCK_STALE_MINUTES)."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, SCRAPE_LOCK_TAB, header=SCRAPE_LOCK_HEADER)
    values = ws.get_all_values()
    if len(values) > 1 and len(values[1]) > 0 and values[1][0]:
        locked_by = values[1][0]
        locked_at_str = values[1][1] if len(values[1]) > 1 else ""
        try:
            locked_at = datetime.strptime(locked_at_str, "%Y-%m-%d %H:%M:%S")
            age_minutes = (datetime.now() - locked_at).total_seconds() / 60
        except Exception:
            age_minutes = None
        if age_minutes is None or age_minutes < SCRAPE_LOCK_STALE_MINUTES:
            return False
        print(
            f"[scrape_lock] Khoa cu cua '{locked_by}' da qua {age_minutes:.0f} phut "
            "(coi nhu tien trinh do da chet) — giao lai khoa."
        )
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.update("A2:B2", [[owner_id, now]])
    return True


def release_scrape_lock(owner_id):
    """Tra khoa — CHI xoa neu dung minh dang giu (tranh vo tinh xoa mat khoa
    cua 1 tien trinh KHAC da gianh duoc no sau khi minh het han/bi loi)."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, SCRAPE_LOCK_TAB, header=SCRAPE_LOCK_HEADER)
    values = ws.get_all_values()
    if len(values) > 1 and len(values[1]) > 0 and values[1][0] == owner_id:
        ws.update("A2:B2", [["", ""]])


SCHEDULE_STATE_TAB = "ScheduleState"
SCHEDULE_STATE_HEADER = ["last_success_at"]


def get_last_schedule_checkpoint():
    """Tra ve thoi diem KET THUC (datetime) cua lan chay DINH KY (khong phai
    'Cap nhat ngay') THANH CONG GAN NHAT — dung de tu dong mo rong cua so cap
    nhat khi phat hien khoang trong lon hon binh thuong (xem
    adsun_daily_report.compute_window). Tra ve None neu chua co lan nao (vd
    lan dau chay) hoac doc loi."""
    try:
        ss = get_report_spreadsheet()
        ws = get_or_create_worksheet(ss, SCHEDULE_STATE_TAB, header=SCHEDULE_STATE_HEADER)
        values = ws.get_all_values()
        if len(values) < 2 or not values[1] or not values[1][0]:
            return None
        return pd.to_datetime(values[1][0]).to_pydatetime()
    except Exception:
        return None


def set_last_schedule_checkpoint(when):
    """Ghi lai thoi diem KET THUC cua lan chay dinh ky THANH CONG — goi SAU
    KHI da day xong len Google Sheets, o CUOI moi lan chay dinh ky thanh
    cong (khong goi cho 'Cap nhat ngay'/'Cap nhat theo khoang' rieng le, chi
    goi cho luong chinh adsun_daily_report.py, de phan anh dung "lan cap
    nhat dinh ky gan nhat" — dung lam moc tu dong mo rong cua so lan sau)."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, SCHEDULE_STATE_TAB, header=SCHEDULE_STATE_HEADER)
    ws.update("A2", [[when.strftime("%Y-%m-%d %H:%M:%S")]])


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


_annotation_row_cache = {}  # row_key -> so dong tren sheet (1-indexed), cache trong bo nho tien trinh


def set_annotation(row_key, ticked=False, note="", colors=None):
    """Ghi DE TOAN BO 1 dong chu thich (khong doc-hop-nhat truoc nua) — trinh
    duyet (web) da tu gop du lieu day du (ticked/note/colors) o phia client
    truoc khi goi ham nay. Lan luu DAU TIEN cho 1 row_key can 1 lan doc
    (ws.find) + 1 lan ghi; TU LAN THU 2 tro di cho CUNG row_key (vd go ghi
    chu roi xoa ngay sau do), nho vi tri dong da cache tu lan truoc nen CHI
    can 1 lan ghi — giam gan mot nua do tre so voi truoc (3 lan goi: find +
    row_values + update). Cache an toan vi tab nay CHI append/update tai
    cho, khong bao gio xoa/sap xep lai dong."""
    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, ANNOTATIONS_TAB, header=ANNOTATIONS_HEADER)

    colors = {k: v for k, v in (colors or {}).items() if v}
    row_values = [
        row_key,
        "TRUE" if ticked else "FALSE",
        note or "",
        json.dumps(colors, ensure_ascii=False),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ]

    cached_row = _annotation_row_cache.get(row_key)
    if cached_row is not None:
        ws.update(f"A{cached_row}:E{cached_row}", [row_values])
        return {"ticked": bool(ticked), "note": note or "", "colors": colors}

    cell = ws.find(row_key, in_column=1)
    if cell is None:
        resp = ws.append_row(row_values)
        try:
            updated_range = resp["updates"]["updatedRange"]
            row_num = int(re.search(r"(\d+)", updated_range.split("!")[-1]).group(1))
            _annotation_row_cache[row_key] = row_num
        except Exception:
            pass
    else:
        ws.update(f"A{cell.row}:E{cell.row}", [row_values])
        _annotation_row_cache[row_key] = cell.row
    return {"ticked": bool(ticked), "note": note or "", "colors": colors}


def set_annotations_batch(items):
    """Ghi NHIEU dong chu thich CUNG LUC bang so luot goi Sheets API IT NHAT
    co the (toi da 3 luot: 1 doc + 1 batch-update + 1 batch-append, BAT KE
    dang xoa/dan bao nhieu o) — dung khi web goi hang loat (vd Backspace tren
    1 vung chon lon, hoac dan nhieu o cung luc). Neu goi set_annotation()
    rieng le cho tung dong trong truong hop nay, hang chuc request chay gan
    nhu dong thoi se de vuot gioi han so luot GHI/phut cua Google Sheets
    API, khien 1 phan thay doi bi loi am tham — day chinh la nguyen nhan gay
    ra loi thuc te "xoa het, mot hoi sau van hien lai" (vi thuc ra chua bao
    gio luu duoc, chi la giao dien hien thi da xoa tam thoi).

    items: list cac dict {row_key, ticked, note, colors}."""
    if not items:
        return

    ss = get_report_spreadsheet()
    ws = get_or_create_worksheet(ss, ANNOTATIONS_TAB, header=ANNOTATIONS_HEADER)

    existing_values = ws.get_all_values()
    row_of_key = {row[0]: i + 1 for i, row in enumerate(existing_values) if row}  # 1-indexed

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    range_updates = []
    new_rows = []
    next_new_row_num = len(existing_values) + 1

    for item in items:
        row_key = str(item.get("row_key") or "").strip()
        if not row_key:
            continue
        colors = {k: v for k, v in (item.get("colors") or {}).items() if v}
        row_values = [
            row_key,
            "TRUE" if item.get("ticked") else "FALSE",
            item.get("note") or "",
            json.dumps(colors, ensure_ascii=False),
            now,
        ]
        row_num = row_of_key.get(row_key)
        if row_num:
            range_updates.append({"range": f"A{row_num}:E{row_num}", "values": [row_values]})
        else:
            new_rows.append(row_values)
            # Phong khi 2 item trong CUNG 1 batch trung row_key (vd du lieu
            # thua) — gan tam vi tri de item sau dung update thay vi them
            # them 1 dong moi nua.
            row_of_key[row_key] = next_new_row_num
            next_new_row_num += 1

    if range_updates:
        ws.batch_update(range_updates)
    if new_rows:
        ws.append_rows(new_rows)


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

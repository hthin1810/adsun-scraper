"""
Ban sao cache du lieu GPS tho (tung ngay/tung xe) tren Google Drive — dung
de GitHub Actions (may ao tam thoi, khong co o dia rieng giua cac lan chay)
cung doc/ghi duoc cache nhu may local, thay vi phai quet lai tu dau moi lan.

QUAN TRONG: KHONG dung service account de ghi (service account KHONG CO dung
luong luu tru rieng tren Drive — Google chan viec service account tu tao/so
huu file, ke ca khi duoc share quyen Editor vao 1 thu muc — chi Google
Workspace tra phi (Shared Drive) moi cho phep). Thay vao do dung OAuth2 uy
quyen TAI KHOAN CA NHAN cua nguoi dung (xem drive_oauth_setup.py — chay 1
lan duy nhat, mo trinh duyet de dang nhap + cap quyen, sau do luu refresh
token vao config.json muc "drive_oauth" — tu dong lam moi access token moi
lan can, khong can dang nhap lai). File luu vao Drive se tinh vao dung luong
CUA NGUOI DUNG (khong phai cua Google Cloud project).

Cau truc tren Drive: nguoi dung tu tao 1 thu muc ten dung "AdsunRawCache" o
Drive ca nhan cua ho — moi bien so xe la 1 thu muc con, moi ngay la 1 file:
    AdsunRawCache/{bien_so}/{YYYY-MM-DD}.xlsx

Truoc khi dung duoc, can BAT API "Google Drive API" cho project Google Cloud
dang dung (Google Cloud Console -> APIs & Services -> Enable APIs -> tim
"Google Drive API" -> Enable).
"""

import io
import json
import os
from datetime import date, timedelta

import pandas as pd
import requests
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request

import sheets_client

TIME_COLUMN = "Thời điểm"
ROOT_FOLDER_NAME = "AdsunRawCache"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
API_FILES = "https://www.googleapis.com/drive/v3/files"
API_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files"
TOKEN_URI = "https://oauth2.googleapis.com/token"

_creds = None
_folder_id_cache = {}  # "" (goc) hoac ten bien so -> id thu muc tren Drive
_plate_files_cache = {}  # ten bien so -> {ten_file: file_id} (cache trong 1 lan chay)


def _load_oauth_config():
    """Doc thong tin OAuth tu config.json (may local/PythonAnywhere) HOAC tu
    bien moi truong DRIVE_OAUTH_JSON (GitHub Actions — dat qua repo secret
    cung ten, xem .github/workflows/adsun_scrape.yml)."""
    config = sheets_client.load_config()
    oauth = config.get("drive_oauth")
    if oauth:
        return oauth
    env_json = os.environ.get("DRIVE_OAUTH_JSON")
    if env_json:
        return json.loads(env_json)
    return None


def _get_credentials():
    global _creds
    if _creds is None:
        oauth = _load_oauth_config()
        if not oauth:
            raise RuntimeError(
                'Chưa cấu hình "drive_oauth" trong config.json (hoặc biến môi trường '
                "DRIVE_OAUTH_JSON) — chạy drive_oauth_setup.py 1 lần trên máy local để lấy "
                "refresh token trước."
            )
        _creds = UserCredentials(
            token=None,
            refresh_token=oauth["refresh_token"],
            client_id=oauth["client_id"],
            client_secret=oauth["client_secret"],
            token_uri=TOKEN_URI,
            scopes=DRIVE_SCOPES,
        )
    if not _creds.valid:
        _creds.refresh(Request())
    return _creds


def _token():
    return _get_credentials().token


def _headers():
    return {"Authorization": f"Bearer {_token()}"}


def _find_folder(name, parent_id=None):
    q = f"name = '{name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    r = requests.get(
        API_FILES,
        headers=_headers(),
        params={"q": q, "fields": "files(id,name)", "includeItemsFromAllDrives": "true", "supportsAllDrives": "true"},
        timeout=30,
    )
    r.raise_for_status()
    files = r.json().get("files", [])
    return files[0]["id"] if files else None


def _create_folder(name, parent_id=None):
    metadata = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    if parent_id:
        metadata["parents"] = [parent_id]
    r = requests.post(
        API_FILES,
        headers={**_headers(), "Content-Type": "application/json"},
        params={"supportsAllDrives": "true"},
        json=metadata,
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["id"]


def _get_or_create_folder(name, parent_id=None):
    fid = _find_folder(name, parent_id)
    if fid:
        return fid
    return _create_folder(name, parent_id)


def _root_folder_id():
    """Tim thu muc goc "AdsunRawCache" — KHONG tu tao, vi service account
    khong co dung luong luu tru rieng (Google gioi han, chi Google Workspace
    tra phi moi tao/so huu file duoc). Nguoi dung phai tu tao thu muc nay
    tren Drive CA NHAN cua ho, roi Share (quyen Editor) cho email service
    account — luc do file moi tinh vao dung luong cua nguoi dung, service
    account chi co quyen doc/ghi."""
    if "" not in _folder_id_cache:
        fid = _find_folder(ROOT_FOLDER_NAME)
        if fid is None:
            raise RuntimeError(
                f'Chưa tìm thấy thư mục "{ROOT_FOLDER_NAME}" trên Google Drive. '
                f'Tạo 1 thư mục tên đúng "{ROOT_FOLDER_NAME}" trong Drive cá nhân của bạn, '
                f"sau đó bấm chuột phải → Share → chia sẻ cho email service account (quyền Editor)."
            )
        _folder_id_cache[""] = fid
    return _folder_id_cache[""]


def _plate_folder_id(plate, create=True):
    if plate not in _folder_id_cache:
        root = _root_folder_id()
        fid = _find_folder(plate, root)
        if fid is None:
            if not create:
                return None
            fid = _create_folder(plate, root)
        _folder_id_cache[plate] = fid
    return _folder_id_cache[plate]


def _list_plate_files(plate):
    """Tra ve dict {ten_file: file_id} cho toan bo file trong thu muc cua 1
    bien so — 1 lan goi API duy nhat, cache lai trong bo nho suot qua trinh
    chay (dung cho ca get_cached_dates va load_cached_range/save_day_cache
    de tranh goi lai API nhieu lan khong can thiet)."""
    if plate in _plate_files_cache:
        return _plate_files_cache[plate]
    folder_id = _plate_folder_id(plate, create=False)
    if folder_id is None:
        _plate_files_cache[plate] = {}
        return {}
    files = {}
    page_token = None
    while True:
        params = {
            "q": f"'{folder_id}' in parents and trashed = false",
            "fields": "nextPageToken, files(id,name)",
            "pageSize": 1000,
            "includeItemsFromAllDrives": "true",
            "supportsAllDrives": "true",
        }
        if page_token:
            params["pageToken"] = page_token
        r = requests.get(API_FILES, headers=_headers(), params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        for f in data.get("files", []):
            files[f["name"]] = f["id"]
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    _plate_files_cache[plate] = files
    return files


def save_day_cache(plate, day: date, day_df: pd.DataFrame):
    """Ghi cache cho 1 (xe, ngay) len Drive NEU CHUA CO SAN — cung ngu nghia
    voi raw_cache.save_day_cache() ban local."""
    file_name = f"{day:%Y-%m-%d}.xlsx"
    existing = _list_plate_files(plate)
    if file_name in existing:
        return False

    folder_id = _plate_folder_id(plate, create=True)
    buf = io.BytesIO()
    day_df.to_excel(buf, index=False)
    buf.seek(0)

    boundary = "adsun_cache_boundary"
    metadata = f'{{"name": "{file_name}", "parents": ["{folder_id}"]}}'
    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{metadata}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    ).encode("utf-8") + buf.read() + f"\r\n--{boundary}--".encode("utf-8")

    r = requests.post(
        API_UPLOAD,
        headers={**_headers(), "Content-Type": f"multipart/related; boundary={boundary}"},
        params={"uploadType": "multipart", "supportsAllDrives": "true"},
        data=body,
        timeout=60,
    )
    r.raise_for_status()
    existing[file_name] = r.json()["id"]
    return True


def cache_complete_past_days(plate, df: pd.DataFrame, today: date, time_column=TIME_COLUMN):
    cached_count = 0
    if df.empty:
        return 0
    for d, day_df in df.groupby(df[time_column].dt.date):
        if d < today:
            if save_day_cache(plate, d, day_df.reset_index(drop=True)):
                cached_count += 1
    return cached_count


def get_cached_dates(plate):
    files = _list_plate_files(plate)
    result = set()
    for name in files:
        stem = name[:-5] if name.endswith(".xlsx") else name
        try:
            result.add(date.fromisoformat(stem))
        except ValueError:
            continue
    return result


def load_cached_range(plate, start_date: date, end_date_exclusive: date):
    """Doc cache tren Drive trong [start_date, end_date_exclusive). Cung tra
    ve (DataFrame gop, set cac ngay CON THIEU) giong ham cung ten trong
    raw_cache.py (ban local)."""
    files = _list_plate_files(plate)
    frames = []
    missing = set()
    d = start_date
    while d < end_date_exclusive:
        file_name = f"{d:%Y-%m-%d}.xlsx"
        file_id = files.get(file_name)
        if file_id is None:
            missing.add(d)
        else:
            try:
                r = requests.get(
                    f"{API_FILES}/{file_id}",
                    headers=_headers(),
                    params={"alt": "media", "supportsAllDrives": "true"},
                    timeout=60,
                )
                r.raise_for_status()
                frames.append(pd.read_excel(io.BytesIO(r.content)))
            except Exception:
                missing.add(d)
        d += timedelta(days=1)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=[TIME_COLUMN])
    return combined, missing

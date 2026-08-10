"""
Cache du lieu GPS tho (tung ngay/tung xe) ra file Excel, vi du lieu cua 1
NGAY DA QUA la vinh vien khong doi. Nho vay khi can tinh lai Vung A/B (vi vua
them vung moi), neu ca thang da co san trong cache thi CHI can doc lai file
Excel + chay Python thuan (khong can mo trinh duyet, khong can doi Adsun xu
ly) — nhanh hon rat nhieu so voi quet lai tu dau.

Cau truc (may local): raw_cache/{bien_so}/{YYYY-MM-DD}.xlsx (moi file la du
lieu tho cua DUNG 1 xe, DUNG 1 ngay). Ngay "hom nay" (chua ket thuc, con thay
doi) khong bao gio duoc cache.

KHI CHAY TREN GITHUB ACTIONS (bien moi truong GITHUB_ACTIONS="true", GitHub
tu dong thiet lap): moi lan chay la 1 may ao HOAN TOAN MOI, khong con o dia
rieng giu lai giua cac lan chay — nen cac ham duoi day TU DONG chuyen sang
doc/ghi qua Google Drive (xem drive_cache.py, dung chung service account voi
Google Sheets) thay vi o dia local. Cac noi goi (recompute_zones.py,
refresh_range.py, backfill_month.py, adsun_daily_report.py) khong can doi gi
ca, van goi dung cac ham nay nhu binh thuong."""

import os
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

CACHE_DIR = Path(__file__).parent / "raw_cache"
TIME_COLUMN = "Thời điểm"

_USE_DRIVE = os.environ.get("GITHUB_ACTIONS", "").strip().lower() == "true"


def cache_path(plate, day: date) -> Path:
    return CACHE_DIR / plate / f"{day:%Y-%m-%d}.xlsx"


def split_by_calendar_day(df: pd.DataFrame, time_column=TIME_COLUMN):
    """Tach 1 DataFrame (nhieu ngay) thanh dict {date: DataFrame con dung 1 ngay}."""
    if df.empty:
        return {}
    groups = {}
    for d, sub in df.groupby(df[time_column].dt.date):
        groups[d] = sub.reset_index(drop=True)
    return groups


def save_day_cache(plate, day: date, day_df: pd.DataFrame):
    """Ghi cache cho 1 (xe, ngay) NEU CHUA CO SAN — du lieu ngay da qua la
    khong doi nen khong can ghi de lai."""
    if _USE_DRIVE:
        import drive_cache
        return drive_cache.save_day_cache(plate, day, day_df)
    path = cache_path(plate, day)
    if path.exists():
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    day_df.to_excel(path, index=False)
    return True


def overwrite_day_cache(plate, day: date, day_df: pd.DataFrame):
    """GHI DE cache cho 1 (xe, ngay) DU DA CO SAN — dung khi biet chac cache
    cu bi sai (vd loi "dinh" du lieu xe khac luc quet, xem adsun_common.
    export_vehicle_report) va can sua lai thu cong. Khac save_day_cache (chi
    ghi khi chua co) — ham nay LUON ghi de."""
    if _USE_DRIVE:
        import drive_cache
        return drive_cache.overwrite_day_cache(plate, day, day_df)
    path = cache_path(plate, day)
    path.parent.mkdir(parents=True, exist_ok=True)
    day_df.to_excel(path, index=False)
    return True


def cache_complete_past_days(plate, df: pd.DataFrame, today: date, time_column=TIME_COLUMN):
    """Tach df theo ngay va cache tat ca cac ngay DA QUA (< today). Goi ham
    nay sau moi lan quet Adsun (kha nang xay thanh cache dan theo thoi gian
    ma khong ton them chi phi quet rieng)."""
    cached_count = 0
    for d, day_df in split_by_calendar_day(df, time_column).items():
        if d < today:
            if save_day_cache(plate, d, day_df):
                cached_count += 1
    return cached_count


def list_cached_plates():
    """Tra ve danh sach ten cac bien so xe DA CO cache (thu muc con)."""
    if _USE_DRIVE:
        import drive_cache
        return drive_cache.list_cached_plates()
    if not CACHE_DIR.exists():
        return []
    return sorted(p.name for p in CACHE_DIR.iterdir() if p.is_dir())


def get_cached_dates(plate):
    if _USE_DRIVE:
        import drive_cache
        return drive_cache.get_cached_dates(plate)
    plate_dir = CACHE_DIR / plate
    if not plate_dir.exists():
        return set()
    return {
        date.fromisoformat(p.stem) for p in plate_dir.glob("*.xlsx") if _is_valid_date(p.stem)
    }


def _is_valid_date(s):
    try:
        date.fromisoformat(s)
        return True
    except ValueError:
        return False


def load_cached_range(plate, start_date: date, end_date_exclusive: date):
    """Doc toan bo cache cua 1 xe trong [start_date, end_date_exclusive).
    Tra ve (DataFrame gop, set cac ngay CON THIEU trong khoang do)."""
    if _USE_DRIVE:
        import drive_cache
        return drive_cache.load_cached_range(plate, start_date, end_date_exclusive)
    frames = []
    missing = set()
    d = start_date
    while d < end_date_exclusive:
        path = cache_path(plate, d)
        if path.exists():
            try:
                frames.append(pd.read_excel(path))
            except Exception:
                missing.add(d)
        else:
            missing.add(d)
        d += timedelta(days=1)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=[TIME_COLUMN])
    return combined, missing

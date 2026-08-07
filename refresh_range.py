"""
Lam moi du lieu cho DUNG 1 KHOANG THOI GIAN CU THE (khong phai ca thang) —
dung khi biet chinh xac lich su/khoang bi loi hoac thieu, sua nhanh hon
nhieu so voi backfill_month.py (chi xu ly dung so ngay lien quan, khong
phai toan bo 31 ngay).

CHAY:
    python refresh_range.py --from "2026-07-24 08:00" --to "2026-07-24 10:00"
"""

import argparse
import sys
from datetime import date, datetime, timedelta

import pandas as pd
from playwright.sync_api import sync_playwright

import raw_cache
import sheets_client
import zone_matching
from adsun_common import (
    export_vehicle_report,
    get_credentials,
    goto_report_page,
    list_all_plates,
    login,
    make_scrape_lock_owner_id,
    wait_for_scrape_lock,
)
from adsun_daily_report import (
    FINAL_COLUMNS,
    HEADER_ROW_INDEX,
    process_vehicle_dataframe,
    upsert_split_by_month,
)


def parse_dt(s):
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"Không đọc được thời điểm: {s!r} (định dạng cần: YYYY-MM-DD HH:MM)")


def pick_preset(earliest_day: date, today: date):
    """Chon preset nho nhat du phu tu earliest_day den hom nay — cang gan
    cang nhanh (Adsun xu ly khoang ngan nhanh hon han khoang dai)."""
    days_back = (today - earliest_day).days + 1
    for key in ("2", "3", "7", "14", "31"):
        if days_back <= int(key):
            return key
    return "31"


def fetch_vehicle_raw(page, plate, range_preset, prev_fingerprint=None):
    """Wrapper mong quanh adsun_common.export_vehicle_report — xem docstring
    ham do ve co che chong "dinh" du lieu xe truoc (da gap thuc te nhieu lan,
    co lan toi 7 xe bi ghi cung 1 du lieu). Tra ve (df, fingerprint_moi)."""
    return export_vehicle_report(page, plate, range_preset, HEADER_ROW_INDEX, prev_fingerprint)


def collect_vehicle_data(page, plate, day_start: date, day_end_inclusive: date, preset, prev_fingerprint=None):
    """Tra ve DataFrame gop du lieu cho vehicle trong khoang ngay lich (ca
    hai dau), uu tien dung cache, chi quet Adsun neu con thieu ngay nao
    (khong tinh 'hom nay' — luon phai quet tuoi vi con thay doi).

    QUAN TRONG: neu thieu cache, se quet tuoi de BOI VAO CACHE, nhung sau do
    LUON DOC LAI TU CACHE (khong dung truc tiep ket qua quet gop nhieu ngay
    lam du lieu bao cao). Ly do: ban quet gop 1 lan cho nhieu ngay lien tuc
    co the bi Adsun tra ve du lieu gop/cat sai (da gap thuc te ngay 27/7 —
    xe 51M51978 bi gop nham 3-4 ngay lien tiep thanh 1 "chuyen di" vo ly khi
    dung truc tiep ban quet 31 ngay thay vi doc lai tung ngay tu cache)."""
    today = date.today()
    range_end_exclusive = day_end_inclusive + timedelta(days=1)
    cached_df, missing = raw_cache.load_cached_range(plate, day_start, range_end_exclusive)
    missing_excluding_today = missing - {today}

    if missing_excluding_today:
        fresh_df, prev_fingerprint = fetch_vehicle_raw(page, plate, preset, prev_fingerprint)
        if fresh_df is not None:
            n = raw_cache.cache_complete_past_days(plate, fresh_df, today)
            if n:
                print(f"  [{plate}] Đã bồi thêm {n} ngày vào cache.")
        cached_df, missing = raw_cache.load_cached_range(plate, day_start, range_end_exclusive)

    frames = [cached_df] if not cached_df.empty else []
    if today < range_end_exclusive:
        fresh, prev_fingerprint = fetch_vehicle_raw(page, plate, "hom-nay", prev_fingerprint)
        if fresh is not None:
            frames.append(fresh)
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, prev_fingerprint


def main():
    parser = argparse.ArgumentParser(description="Làm mới dữ liệu cho 1 khoảng thời gian cụ thể")
    parser.add_argument("--from", dest="range_from", required=True, help='VD: "2026-07-24 08:00"')
    parser.add_argument("--to", dest="range_to", required=True, help='VD: "2026-07-24 10:00"')
    parser.add_argument("--show", action="store_true")
    args = parser.parse_args()

    range_start = parse_dt(args.range_from)
    range_end = parse_dt(args.range_to)
    if range_end <= range_start:
        print("Lỗi: thời điểm kết thúc phải sau thời điểm bắt đầu.", file=sys.stderr)
        sys.exit(1)

    day_start = range_start.date()
    day_end_inclusive = range_end.date()
    today = date.today()
    preset = pick_preset(day_start, today)

    print(f"Làm mới khoảng: {range_start} -> {range_end} (preset quét dự phòng: {preset} ngày)")

    username, password = get_credentials()
    plate_to_df = {}
    failed_plates = set()

    # Khoa quet Adsun toan cuc (chung voi may local + GitHub Actions) — xem
    # adsun_common.acquire_scrape_lock.
    lock_owner = make_scrape_lock_owner_id()
    if not wait_for_scrape_lock(lock_owner):
        print("Đang có tiến trình khác quét Adsun, bỏ qua lần chạy này.", file=sys.stderr)
        try:
            sheets_client.set_refresh_status("error", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        except Exception:
            pass
        sys.exit(1)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=not args.show)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            try:
                login(page, username, password)
                print("Đăng nhập thành công.")
                goto_report_page(page)
                plates = list_all_plates(page)
                print(f"Tìm thấy {len(plates)} xe: {', '.join(plates)}")
                prev_fingerprint = None  # xem adsun_common.export_vehicle_report
                for plate in plates:
                    try:
                        plate_to_df[plate], prev_fingerprint = collect_vehicle_data(
                            page, plate, day_start, day_end_inclusive, preset, prev_fingerprint
                        )
                    except Exception as e:
                        print(f"  [{plate}] LỖI khi lấy dữ liệu: {e}", file=sys.stderr)
                        plate_to_df[plate] = pd.DataFrame()
                        failed_plates.add(plate)
            except Exception as e:
                print(f"Lỗi: {e}", file=sys.stderr)
                try:
                    sheets_client.set_refresh_status("error", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                except Exception:
                    pass
                sys.exit(1)
            finally:
                browser.close()
    finally:
        try:
            sheets_client.release_scrape_lock(lock_owner)
        except Exception:
            pass

    try:
        zone_polygons = zone_matching.load_zone_polygons()
        print(f"Đã tải {len(zone_polygons)} vùng từ Google Sheets.")
    except Exception as e:
        zone_polygons = None
        print(f"Không tải được vùng ({e}).", file=sys.stderr)

    try:
        pair_rules = sheets_client.list_pair_rules()
    except Exception:
        pair_rules = []

    try:
        easy_pass_zones = zone_matching.load_easy_pass_zones()
    except Exception:
        easy_pass_zones = set()

    all_rows = []
    for plate, df in plate_to_df.items():
        if df is None or df.empty:
            continue
        all_rows.extend(
            process_vehicle_dataframe(plate, df, range_start, range_end, zone_polygons, pair_rules, easy_pass_zones)
        )

    if not all_rows:
        print("Không có dữ liệu hành trình nào trong khoảng thời gian này.")
        try:
            sheets_client.set_refresh_status("done", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        except Exception:
            pass
        return

    merged = pd.DataFrame(all_rows)
    merged = merged.sort_values("Thời điểm A").reset_index(drop=True)
    merged.insert(0, "STT", range(1, len(merged) + 1))
    merged = merged[FINAL_COLUMNS]

    successful_plates = set(plate_to_df.keys()) - failed_plates
    upsert_split_by_month(merged, range_start, range_end, only_replace_plates=successful_plates)
    print(f"Hoàn tất — đã làm mới {len(merged)} dòng trong khoảng {range_start} -> {range_end}.")
    if failed_plates:
        print(
            f"(Xe {', '.join(sorted(failed_plates))} bị lỗi khi quét lần này — dữ liệu cũ của các xe "
            "đó vẫn giữ nguyên, chưa được làm mới.)",
            file=sys.stderr,
        )

    try:
        sheets_client.set_refresh_status("done", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    except Exception:
        pass


if __name__ == "__main__":
    main()

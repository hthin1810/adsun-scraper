"""
Tinh lai Vung A/B cho TOAN BO du lieu DA CO SAN trong cache (KHONG dang
nhap/quet lai Adsun) — dung khi vua them/sua/xoa vung hoac cap vung tuy
chinh tren ban do va muon ap dung ngay cho toan bo lich su, khong can cho
lan quet tu dong tiep theo. Nhanh hon rat nhieu so voi backfill_month.py hay
refresh_range.py vi khong can mo trinh duyet — chi doc file cache + tinh lai
bang Python thuan.

Pham vi: tu ngay SOM NHAT co trong cache cho DEN HET HOM QUA (khong dung du
lieu "hom nay" vi cache khong bao gio luu ngay chua ket thuc — du lieu hom
nay van do cac lan quet dinh ky binh thuong phu trach, khong can dung script
nay dung tram len).

CHAY:
    python recompute_zones.py
"""

import sys
from datetime import date, datetime, timedelta

import pandas as pd

import raw_cache
import sheets_client
import zone_matching
from adsun_daily_report import FINAL_COLUMNS, process_vehicle_dataframe, upsert_split_by_month


def list_cached_plates():
    if not raw_cache.CACHE_DIR.exists():
        return []
    return sorted(p.name for p in raw_cache.CACHE_DIR.iterdir() if p.is_dir())


def main():
    today = date.today()
    plates = list_cached_plates()
    if not plates:
        print("Không có dữ liệu cache nào để tính lại.", file=sys.stderr)
        sys.exit(1)
    print(f"Tìm thấy {len(plates)} xe có dữ liệu cache: {', '.join(plates)}")

    earliest = None
    for plate in plates:
        dates = raw_cache.get_cached_dates(plate)
        if dates:
            d = min(dates)
            if earliest is None or d < earliest:
                earliest = d
    if earliest is None:
        print("Không có ngày nào trong cache.", file=sys.stderr)
        sys.exit(1)

    range_start = datetime.combine(earliest, datetime.min.time())
    range_end = datetime.combine(today, datetime.min.time())  # den het hom qua
    if range_end <= range_start:
        print("Chưa có ngày nào đã qua trong cache để tính lại.")
        return

    zone_polygons = zone_matching.load_zone_polygons()
    print(f"Đã tải {len(zone_polygons)} vùng từ Google Sheets.")
    pair_rules = sheets_client.list_pair_rules()
    print(f"Đã tải {len(pair_rules)} cặp vùng tùy chỉnh.")
    easy_pass_zones = zone_matching.load_easy_pass_zones()
    if easy_pass_zones:
        print(f"Đã tải {len(easy_pass_zones)} vùng \"dễ đi qua\".")
    print(f"Tính lại Vùng cho khoảng: {range_start} -> {range_end} (dùng cache, không quét Adsun)")

    all_rows = []
    used_plates = set()
    for plate in plates:
        cached_df, missing = raw_cache.load_cached_range(plate, earliest, today)
        if cached_df.empty:
            continue
        used_plates.add(plate)
        all_rows.extend(
            process_vehicle_dataframe(
                plate, cached_df, range_start, range_end, zone_polygons, pair_rules, easy_pass_zones
            )
        )

    if not all_rows:
        print("Không có dòng nào được tính lại.")
        try:
            sheets_client.set_refresh_status("done", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        except Exception:
            pass
        return

    merged = pd.DataFrame(all_rows)
    merged = merged.sort_values("Thời điểm A").reset_index(drop=True)
    merged.insert(0, "STT", range(1, len(merged) + 1))
    merged = merged[FINAL_COLUMNS]

    upsert_split_by_month(merged, range_start, range_end, only_replace_plates=used_plates)
    print(f"Hoàn tất — đã tính lại {len(merged)} dòng cho khoảng {range_start} -> {range_end}.")

    try:
        sheets_client.set_refresh_status("done", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    except Exception:
        pass


if __name__ == "__main__":
    main()

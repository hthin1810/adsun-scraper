"""
Xuat bao cao hanh trinh TONG HOP cho TOAN BO xe tren Adsun, gop vao Google
Sheets (1 tab / thang) + 1 ban sao Excel local. Muc dich: chay tu dong MOI
GIO (qua Windows Task Scheduler) de cap nhat gan-thoi-gian-thuc.

"Ngay lam viec" (business day) tinh tu 8h sang hom nay/hom qua den THOI DIEM
CHAY (khong co dinh la 8h-8h nhu truoc) — moi lan chay se GHI DE dung phan
du lieu cua ngay lam viec hien tai trong tab thang tren Sheets, GIU NGUYEN
cac ngay khac da co san trong tab do (xem sheets_client.upsert_range_report).
Nho vay moi lan chay chi ton thoi gian quet + ghi cho 1 ngay, khong lien
quan kich thuoc ca thang.

CACH HOAT DONG:
    1. Dang nhap, lay danh sach toan bo xe (dropdown "Thiết bị").
    2. Voi moi xe: chon xe, chon khoang "2 ngày" (du de phu 8h-8h, loc chinh
       xac hon o buoc sau), bam Xem, xuat Excel rieng cho xe do (luu tam vao
       thu muc raw/).
    3. Doc tat ca file rieng, LOC lai dung khung gio cua ngay lam viec hien tai.
    4. Voi tung xe: bo cac dong khong co Vung, gop cac dong LIEN TIEP cung
       Vung thanh 1 diem (giu thoi diem som nhat), roi ap dung uu tien vung
       neo (MỎ* > BÃI RẠCH CHIẾC/BÃI TẬP KẾT) de tao dong Vung A -> Vung B.
       Gop tat ca xe, sap xep theo Thoi diem A tang dan.
    5. Ghi ban sao Excel local + upsert vao tab thang tren Google Sheets.

CAI DAT (chay 1 lan):
    pip install playwright pandas openpyxl gspread google-auth shapely
    playwright install chromium

CHAY THU CONG:
    python adsun_daily_report.py

DAT LICH CHAY TU DONG MOI GIO (Windows Task Scheduler) — xem huong dan hoi
lai de duoc ho tro tao/sua task.
"""

import argparse
import shutil
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

import raw_cache
import sheets_client
import zone_matching
from adsun_common import (
    get_credentials,
    goto_report_page,
    has_no_data,
    list_all_plates,
    login,
    select_date_range,
    select_vehicle,
    view_report,
    wait_report_loaded,
    export_current_report,
)

RAW_RANGE_PRESET = "2"  # "2 ngày": du phu khung 8h hom qua - 8h hom nay
TIME_COLUMN = "Thời điểm"
ZONE_COLUMN = "Vùng"
HEADER_ROW_INDEX = 3  # dong 4 trong file excel (0-indexed) la dong ten cot

EMPTY_ZONE_VALUES = {"", "-", "–", "—"}

FINAL_COLUMNS = [
    "STT",
    "Thời điểm A",
    "Thời điểm B",
    "Biển số xe",
    "Vùng A",
    "Vùng B",
    "Tọa độ A",
    "Tọa độ B",
    "Địa chỉ A",
    "Địa chỉ B",
    "Nhiên liệu A",
    "Nhiên liệu B",
]

HIDDEN_COLUMNS = [
    "Tọa độ A",
    "Tọa độ B",
    "Địa chỉ A",
    "Địa chỉ B",
    "Nhiên liệu A",
    "Nhiên liệu B",
]

# Chi 6 cot nay duoc day len Google Sheets (spreadsheet bao cao rieng, xem
# sheets_client.get_report_spreadsheet) — Toa do/Dia chi/Nhien lieu van con
# day du trong ban Excel local (HIDDEN_COLUMNS o tren) de tham khao khi can.
REPORT_COLUMNS = [
    "STT",
    "Thời điểm A",
    "Thời điểm B",
    "Biển số xe",
    "Vùng A",
    "Vùng B",
]


ROLLING_WINDOW_HOURS = 2


def compute_window(now: datetime, rolling_hours=ROLLING_WINDOW_HOURS):
    """Tra ve (now - rolling_hours, now) — cua so "N gio gan nhat" (mac dinh
    2 gio), KHONG con la ca ngay lam viec tu 8h sang nhu truoc. Ly do doi:
    moi lan chay (moi 15 phut tren may local, moi gio qua GitHub Actions) se
    GHI DE lai dung khoang nay — vi cac lan chay lien tiep co khoang trung
    lap nhau lon (2 gio > chu ky chay), 1 lan bi loi/thieu du lieu (vd su co
    ghi Google Sheets ngay 27/7) se DUOC TU SUA boi lan chay ke tiep, khong
    can can thiep thu cong. Cung nhanh hon ro ret so voi xu ly ca ngay lam
    viec (vd luc 15h chieu se phai xu ly toi 7 tieng du lieu neu con dung
    kieu cu)."""
    return now - timedelta(hours=rolling_hours), now


def export_all_vehicles(page, raw_dir: Path, range_preset=RAW_RANGE_PRESET):
    plates = list_all_plates(page)
    if not plates:
        raise RuntimeError("Không lấy được danh sách xe từ dropdown 'Thiết bị'.")
    print(f"Tìm thấy {len(plates)} xe: {', '.join(plates)}")

    raw_dir.mkdir(parents=True, exist_ok=True)
    results = []  # list of (plate, path or None)

    for plate in plates:
        # Tai lai trang truoc moi xe: sau khi xuat Excel, toast thong bao
        # "Da xu ly du lieu xong..." co the con hien va CHE cac nut khac,
        # gay loi day chuyen cho xe ke tiep. Reload dam bao trang thai sach.
        goto_report_page(page)
        try:
            select_vehicle(page, plate)
            select_date_range(page, range_preset)
            view_report(page)
            wait_report_loaded(page, timeout_seconds=300)

            if has_no_data(page):
                print(f"  [{plate}] Không có dữ liệu, bỏ qua.")
                results.append((plate, None))
                continue

            download = export_current_report(page)
            raw_path = raw_dir / f"{plate}.xlsx"
            download.save_as(str(raw_path))
            print(f"  [{plate}] Đã xuất tạm: {raw_path.name}")
            results.append((plate, raw_path))

            # Boi cache cho cac ngay DA QUA (khong bao gio cache "hom nay" vi
            # con thay doi) — tan dung du lieu vua quet ma khong ton them chi
            # phi, giup lan "lam moi vi them vung" sau nay nhanh hon nhieu.
            try:
                day_df = pd.read_excel(raw_path, header=HEADER_ROW_INDEX)
                n = raw_cache.cache_complete_past_days(plate, day_df, date.today())
                if n:
                    print(f"  [{plate}] Đã bồi thêm {n} ngày vào cache.")
            except Exception as e:
                print(f"  [{plate}] Không bồi được cache: {e}", file=sys.stderr)

        except Exception as e:
            print(f"  [{plate}] LỖI khi xuất, bỏ qua xe này: {e}", file=sys.stderr)
            # Dung chuoi "ERROR" (khac None) de phan biet voi truong hop
            # "khong co du lieu that su" (has_no_data) — xe loi KHONG duoc
            # dua vao danh sach thay du lieu (xem only_replace_plates o main()),
            # tranh xoa mat du lieu cu con dang dung cua xe do.
            results.append((plate, "ERROR"))

    return results


def is_empty_zone(value):
    if pd.isna(value):
        return True
    return str(value).strip() in EMPTY_ZONE_VALUES


def extract_zone_points(df):
    """Loc bo cac dong khong co Vung, roi gop cac dong LIEN TIEP cung mot Vung
    thanh 1 diem duy nhat — nhung GIU LAI CA HAI moc thoi gian cua nhom do:
    thoi diem DEN (ping dau tien, luu o TIME_COLUMN nhu truoc) va thoi diem
    RỜI (ping CUOI CUNG con o vung do truoc khi doi vung, luu o "_leave").
    build_transition_rows se dung thoi diem RỜI cho Vung A (gio lay = luc xe
    THUC SU roi khoi noi lay hang, khong phai luc vua den do) va thoi diem
    DEN cho Vung B (gio giao = luc xe vua toi noi giao)."""
    points = []
    for _, row in df.iterrows():
        zone = row[ZONE_COLUMN]
        if is_empty_zone(zone):
            continue
        snapshot = {
            TIME_COLUMN: row[TIME_COLUMN],
            "Tọa độ": row.get("Tọa độ"),
            "Địa chỉ": row.get("Địa chỉ"),
            "Nhiên liệu": row.get("Nhiên liệu"),
        }
        if points and points[-1][ZONE_COLUMN] == zone:
            points[-1]["_leave"] = snapshot  # van con o vung nay -> cap nhat moc "roi" moi nhat
            continue
        points.append({**snapshot, ZONE_COLUMN: zone, "_leave": snapshot})
    return points


# Vung "neo" (bai/mo) mac dinh luon o vi tri Vung A: vung bat dau bang "MỎ",
# hoac dung "BÃI RẠCH CHIẾC" / "BÃI TẬP KẾT". Cac vung khac (Fico Mai Chi Tho,
# Fico Cat Lai...) mac dinh la Vung B.
ANCHOR_EXACT = {"BÃI RẠCH CHIẾC", "BÃI TẬP KẾT"}

# "VLXD BỜM" LUON duoc uu tien chon lam Vung B thuc su neu xuat hien trong 1
# nhom vai tro B lien tiep — vi duong vao "VLXD BỜM" phai di ngang qua "VLXD
# DƯƠNG" truoc (va thuong quay lai ngang qua DƯƠNG lan nua khi ra), nen neu
# chi lay diem CUOI CUNG trong nhom (mac dinh) se de nham thanh DƯƠNG la diem
# den thuc su thay vi BỜM (da gap thuc te: xe 51M68032 luc 3h32 sang 28/7).
PRIORITY_B_ZONES = {"VLXD BỜM"}


def _name_upper(p):
    return str(p[ZONE_COLUMN]).strip().upper()


def is_anchor_default(name):
    name_upper = str(name).strip().upper()
    return name_upper.startswith("MỎ") or name_upper in ANCHOR_EXACT


def _forced_close_as_b(name_upper, prev_upper):
    """Ngoai le rieng cho BÃI TẬP KẾT: neu diem NGAY TRUOC no la 1 vung "MỎ*"
    (xe tu mo cho hang ve bai), thi lan nay BÃI TẬP KẾT dong vai Vung B (khep
    chang tu mo) thay vi vai Vung A mac dinh cua no. Neu sau do xe di tiep
    sang vung khac, BÃI TẬP KẾT lai dong vai Vung A (mac dinh) de mo chang moi."""
    return name_upper == "BÃI TẬP KẾT" and prev_upper is not None and prev_upper.startswith("MỎ")


def _pair_touches(points, idx, pair_rules_upper):
    """True neu diem points[idx] la 1 dau (A hoac B) cua 1 cap tuy chinh nguoi
    dung da khai bao, xet voi diem NGAY TRUOC hoac NGAY SAU no."""
    name_upper = _name_upper(points[idx])
    if idx > 0 and (_name_upper(points[idx - 1]), name_upper) in pair_rules_upper:
        return True
    if idx + 1 < len(points) and (name_upper, _name_upper(points[idx + 1])) in pair_rules_upper:
        return True
    return False


def _pick_group_representative(points, i, j):
    """Chon diem dai dien cho 1 nhom vai tro lien tiep points[i..j]: uu tien
    diem nao thuoc PRIORITY_B_ZONES (vd "VLXD BỜM") neu co trong nhom, neu
    khong moi lay diem CUOI CUNG nhu mac dinh."""
    for k in range(i, j + 1):
        if _name_upper(points[k]) in PRIORITY_B_ZONES:
            return points[k]
    return points[j]


def _collapse_runs(points, pair_rules_upper):
    """Gop cac diem LIEN TIEP cung vai tro mac dinh (A voi A, B voi B) thanh 1
    diem duy nhat (giu diem CUOI CUNG trong nhom, TRU KHI nhom co diem thuoc
    PRIORITY_B_ZONES — xem _pick_group_representative) — vi du xe di ngang
    qua nhieu vung dich lien tiep truoc khi quay lai dung vung neo xuat phat,
    chi vung dich CUOI cung truoc khi quay lai moi la diem dung that su. Cac
    diem la "diem noi" (hinge — hoac bi ep dong vai B nhu BÃI TẬP KẾT sau MỎ,
    hoac la 1 dau cua cap tuy chinh) khong bao gio bi gop, luon dung rieng le."""
    reduced = []
    n = len(points)
    i = 0
    while i < n:
        name_upper = _name_upper(points[i])
        prev_upper = _name_upper(points[i - 1]) if i > 0 else None
        forced = _forced_close_as_b(name_upper, prev_upper)
        boundary = forced or _pair_touches(points, i, pair_rules_upper)

        if boundary:
            reduced.append(points[i])
            i += 1
            continue

        role = "A" if is_anchor_default(points[i][ZONE_COLUMN]) else "B"
        j = i
        while j + 1 < n:
            nxt_upper = _name_upper(points[j + 1])
            nxt_prev_upper = _name_upper(points[j])
            nxt_forced = _forced_close_as_b(nxt_upper, nxt_prev_upper)
            nxt_boundary = nxt_forced or _pair_touches(points, j + 1, pair_rules_upper)
            nxt_role = "A" if is_anchor_default(points[j + 1][ZONE_COLUMN]) else "B"
            if nxt_boundary or nxt_role != role:
                break
            j += 1
        reduced.append(_pick_group_representative(points, i, j))
        i = j + 1
    return reduced


HOANG_THINH_NAME = "HOÀNG THỊNH RẠCH CHIẾC"


def _is_anchor_neighbor(name_upper):
    return name_upper is not None and (name_upper.startswith("MỎ") or name_upper in ANCHOR_EXACT)


def _filter_hoang_thinh_noise(points):
    """HOÀNG THỊNH RẠCH CHIẾC nam rat gan BÃI RẠCH CHIẾC ve toa do, de bi
    nham la 1 diem dung that su trong khi thuc ra chi la nhieu GPS luc xe
    dang di ngang qua khu vuc do tren duong toi/roi 1 vung khac (da gap thuc
    te: an mat mot luot ghe FICO MAI CHÍ THỌ ngay 24/7 vi HOÀNG THỊNH xen vao
    ngay sau do). Chi cong nhan la diem dung THAT SU neu CA HAI diem lien ke
    (ngay truoc va ngay sau no trong day diem) deu la vung neo (MỎ*/BÃI TẬP
    KẾT/BÃI RẠCH CHIẾC) — neu khong, loai bo hoan toan khoi day diem (nhu
    chua tung xuat hien), tranh no "an" mat diem den that su ke ben."""
    n = len(points)
    result = []
    for i, p in enumerate(points):
        if _name_upper(p) == HOANG_THINH_NAME:
            prev_upper = _name_upper(points[i - 1]) if i > 0 else None
            next_upper = _name_upper(points[i + 1]) if i + 1 < n else None
            if not (_is_anchor_neighbor(prev_upper) and _is_anchor_neighbor(next_upper)):
                continue
        result.append(p)
    return result


def build_transition_rows(points, plate, pair_rules=None):
    """Tao cac dong chuyen vung A->B tu day diem vung da loc (xem
    _collapse_runs de biet buoc gop truoc). Duyet day da gop, giu 1 "diem dang
    mo" (ung vien Vung A):
      - Neu (diem dang mo -> diem hien tai) khop 1 cap tuy chinh nguoi dung
        khai bao -> LUON tao dong, bat ke vai tro mac dinh.
      - Neu diem hien tai la vung neo mac dinh (va khong bi ep dong vai B) ->
        cap nhat lam diem dang mo moi (gop chuoi neo lien tiep).
      - Nguoc lai (dong vai Vung B): tao dong (diem dang mo -> diem hien tai)
        neu da co diem dang mo; sau do diem nay tro thanh diem dang mo MOI chi
        khi no se "mo lai" duoc (bi ep dong vai B nhu BÃI TẬP KẾT, hoac la dau
        A cua 1 cap tuy chinh voi diem ke tiep) — neu khong thi dong lai (ve
        vung neo, khong ghi)."""
    points = _filter_hoang_thinh_noise(points)
    pair_rules_upper = {(a.strip().upper(), b.strip().upper()) for a, b in (pair_rules or [])}
    reduced = _collapse_runs(points, pair_rules_upper)

    rows = []
    open_point = None
    n = len(reduced)

    def make_row(a, b):
        a_leave = a["_leave"]  # gio lay = luc RỜI khoi Vung A (ping cuoi cung con o do)
        return {
            "Thời điểm A": a_leave[TIME_COLUMN],
            "Thời điểm B": b[TIME_COLUMN],
            "Biển số xe": plate,
            "Vùng A": a[ZONE_COLUMN],
            "Vùng B": b[ZONE_COLUMN],
            "Tọa độ A": a_leave["Tọa độ"],
            "Tọa độ B": b["Tọa độ"],
            "Địa chỉ A": a_leave["Địa chỉ"],
            "Địa chỉ B": b["Địa chỉ"],
            "Nhiên liệu A": a_leave["Nhiên liệu"],
            "Nhiên liệu B": b["Nhiên liệu"],
        }

    for idx, p in enumerate(reduced):
        name_upper = _name_upper(p)
        prev_upper = _name_upper(reduced[idx - 1]) if idx > 0 else None
        nxt_upper = _name_upper(reduced[idx + 1]) if idx + 1 < n else None
        forced_close_as_b = _forced_close_as_b(name_upper, prev_upper)

        pair_override = open_point is not None and (_name_upper(open_point), name_upper) in pair_rules_upper
        if pair_override:
            rows.append(make_row(open_point, p))
            will_reopen = (
                forced_close_as_b
                or (nxt_upper is not None and (name_upper, nxt_upper) in pair_rules_upper)
                or (is_anchor_default(p[ZONE_COLUMN]) and not forced_close_as_b)
            )
            open_point = p if will_reopen else None
            continue

        if is_anchor_default(p[ZONE_COLUMN]) and not forced_close_as_b:
            open_point = p
            continue

        if open_point is None:
            continue  # chua biet xe xuat phat tu vung neo nao, bo qua

        rows.append(make_row(open_point, p))
        will_reopen = forced_close_as_b or (nxt_upper is not None and (name_upper, nxt_upper) in pair_rules_upper)
        open_point = p if will_reopen else None

    return rows


def process_vehicle_dataframe(plate, df, window_start, window_end, zone_polygons=None, pair_rules=None):
    """Tinh Vung + tao cac dong chuyen vung cho 1 xe TU TOAN BO du lieu ping
    truyen vao (KHONG loc theo khung gio truoc khi tinh), roi CHI LOC CAC
    DONG KET QUA (transition) theo "Thời điểm A" nam trong [window_start,
    window_end). Dung chung cho ca luong doc tu file tam (adsun_daily_report)
    va luong doc tu cache/DataFrame trong bo nho (backfill_month, refresh_range).

    QUAN TRONG: truoc day ham nay loc RAW PING truoc (df = df[window_start <=
    t < window_end]) roi moi tinh diem/chuyen vung — dieu nay lam GAY mat 1
    chuyen di neu diem KET THUC (Thời điểm B) roi ra ngoai window nhung diem
    BAT DAU (Thời điểm A) van nam trong window (vi ping ket thuc bi cat mat,
    "diem dang mo" khong bao gio duoc dong lai nen khong sinh ra dong nao).
    Day chinh la nguyen nhan lam mat dong 51M65797 09:57:26 -> 10:28:01 khi
    chay refresh_range.py voi --to "...10:00" (dong bi cat vi 10:28:01 > 10:00).
    Sua bang cach: tinh transition tu FULL du lieu ping duoc truyen vao (caller
    chiu trach nhiem truyen du du lieu, vd ca ngay/nhieu ngay), chi loc O DAU
    RA theo Thời điểm A — vay du chuyen di bat dau trong window luon duoc tinh
    dung du no ket thuc sau window bao nhieu."""
    if TIME_COLUMN not in df.columns or ZONE_COLUMN not in df.columns:
        print(f"  [{plate}] Thiếu cột cần thiết, bỏ qua.", file=sys.stderr)
        return []

    if df.empty:
        return []
    df = df.sort_values(TIME_COLUMN)

    if zone_polygons is not None:
        # Ghi de cot Vung cua Adsun bang vung tu tinh tu toa do — giai
        # quyet van de Adsun khong hoi to vung moi tao cho du lieu qua khu.
        df = zone_matching.compute_zone_column(df, zone_polygons)

    points = extract_zone_points(df)
    rows = build_transition_rows(points, plate, pair_rules)
    return [r for r in rows if window_start <= r["Thời điểm A"] < window_end]


def merge_reports(results, window_start, window_end, zone_polygons=None, pair_rules=None):
    all_rows = []
    for plate, raw_path in results:
        if raw_path is None or raw_path == "ERROR":
            continue
        try:
            df = pd.read_excel(raw_path, header=HEADER_ROW_INDEX)
        except Exception as e:
            print(f"  [{plate}] LỖI khi đọc file tạm {raw_path.name}: {e}", file=sys.stderr)
            continue

        all_rows.extend(process_vehicle_dataframe(plate, df, window_start, window_end, zone_polygons, pair_rules))

    if not all_rows:
        return None

    merged = pd.DataFrame(all_rows)
    merged = merged.sort_values("Thời điểm A").reset_index(drop=True)
    merged.insert(0, "STT", range(1, len(merged) + 1))
    return merged[FINAL_COLUMNS]


def format_workbook(path, hidden_column_names):
    """An cac cot chi dinh, va chinh do rong tat ca cot khop voi do dai ky tu
    dai nhat trong cot do (ke ca header)."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    header = [c.value for c in ws[1]]

    for name in hidden_column_names:
        if name in header:
            col_letter = get_column_letter(header.index(name) + 1)
            ws.column_dimensions[col_letter].hidden = True

    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        ws.column_dimensions[col_letter].width = max_len + 2

    wb.save(path)


def upsert_split_by_month(merged, range_start, range_end, only_replace_plates=None):
    """upsert_range_report ghi theo 1 tab-thang duy nhat — neu khoang can cap
    nhat vo tinh vat qua ranh gioi thang (hiem, chi xay quanh nua dem dau
    thang do dung cua so lan can "N gio gan nhat"), chia nho ra ghi dung
    tung tab thang tuong ung, tranh ghi nham/mat du lieu thang ke tiep.

    only_replace_plates: xem docstring cua sheets_client.upsert_range_report
    — BAT BUOC truyen dung tap bien so DA QUET THANH CONG lan nay, neu khong
    nhung xe bi loi/bo qua se bi XOA MAT du lieu cu thay vi duoc giu nguyen."""
    boundaries = [range_start]
    cursor = range_start.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    while cursor <= range_end:
        next_month = cursor.month + 1 if cursor.month < 12 else 1
        next_year = cursor.year if cursor.month < 12 else cursor.year + 1
        cursor = cursor.replace(year=next_year, month=next_month)
        if range_start < cursor < range_end:
            boundaries.append(cursor)
    boundaries.append(range_end)
    boundaries = sorted(set(boundaries))

    for i in range(len(boundaries) - 1):
        seg_start, seg_end = boundaries[i], boundaries[i + 1]
        seg_rows = merged[(merged["Thời điểm A"] >= seg_start) & (merged["Thời điểm A"] < seg_end)]
        tab_name = f"{seg_start:%Y-%m}"
        sheets_client.upsert_range_report(
            tab_name, seg_start, seg_end, seg_rows[REPORT_COLUMNS], only_replace_plates=only_replace_plates
        )
        print(
            f"Đã cập nhật tab '{tab_name}' cho khoảng {seg_start:%d/%m/%Y %H:%M} -> "
            f"{seg_end:%d/%m/%Y %H:%M} ({len(seg_rows)} dòng)."
        )


def main():
    parser = argparse.ArgumentParser(
        description="Xuất báo cáo hành trình tổng hợp toàn bộ xe (8h hôm qua - 8h hôm nay)"
    )
    parser.add_argument(
        "--output-dir",
        default=str(Path(__file__).parent / "reports"),
        help="Thư mục lưu file Excel tổng hợp (mặc định: ./reports)",
    )
    parser.add_argument(
        "--show", action="store_true", help="Hiện trình duyệt (mặc định chạy ẩn, phù hợp chạy tự động)"
    )
    parser.add_argument(
        "--keep-raw", action="store_true", help="Giữ lại các file Excel riêng từng xe (để kiểm tra)"
    )
    args = parser.parse_args()

    now = datetime.now()
    window_start, window_end = compute_window(now)
    print(f"Khung thời gian báo cáo ({ROLLING_WINDOW_HOURS} giờ gần nhất): {window_start:%d/%m/%Y %H:%M} -> {window_end:%d/%m/%Y %H:%M}")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    raw_dir = output_dir / "_raw" / f"{window_start:%Y%m%d}"

    username, password = get_credentials()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.show)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            login(page, username, password)
            print("Đăng nhập thành công.")

            goto_report_page(page)
            results = export_all_vehicles(page, raw_dir)

        except Exception as e:
            page.screenshot(path="error.png")
            print(f"\nLỗi: {e}\nĐã lưu ảnh chụp màn hình error.png để kiểm tra.", file=sys.stderr)
            sys.exit(1)
        finally:
            browser.close()

    try:
        zone_polygons = zone_matching.load_zone_polygons()
        print(f"Đã tải {len(zone_polygons)} vùng từ Google Sheets để tự tính Vùng theo tọa độ.")
    except Exception as e:
        zone_polygons = None
        print(
            f"Không tải được danh sách vùng từ Google Sheets ({e}). "
            "Dùng tạm cột Vùng có sẵn của Adsun.",
            file=sys.stderr,
        )

    try:
        pair_rules = sheets_client.list_pair_rules()
        print(f"Đã tải {len(pair_rules)} cặp vùng tùy chỉnh.")
    except Exception as e:
        pair_rules = []
        print(f"Không tải được danh sách cặp vùng tùy chỉnh ({e}).", file=sys.stderr)

    # Chi nhung xe QUET THANH CONG lan nay (co du lieu hoac xac nhan khong co
    # du lieu that su) moi duoc phep thay du lieu cu — xe bi loi ("ERROR")
    # se duoc GIU NGUYEN du lieu cu, khong bi xoa mat (xem upsert_split_by_month).
    successful_plates = {plate for plate, status in results if status != "ERROR"}

    print("Đang gộp và lọc dữ liệu theo đúng khung giờ...")
    merged = merge_reports(results, window_start, window_end, zone_polygons, pair_rules)

    if merged is None or merged.empty:
        print("Không có dữ liệu hành trình nào trong khung thời gian này cho toàn bộ xe.")
        return

    output_path = output_dir / f"BaoCaoHanhTrinhTongHop_{window_start:%Y-%m-%d}.xlsx"
    try:
        merged.to_excel(output_path, index=False)
    except PermissionError:
        # File dang bi mo (vd. dang mo trong Excel) nen khong ghi de duoc,
        # luu sang ten khac de khong mat du lieu ngay hom nay.
        output_path = output_dir / f"BaoCaoHanhTrinhTongHop_{window_start:%Y-%m-%d}_{datetime.now():%H%M%S}.xlsx"
        merged.to_excel(output_path, index=False)
        print(
            f"File gốc đang bị mở (có thể đang mở trong Excel), đã lưu vào tên khác: {output_path.name}",
            file=sys.stderr,
        )
    format_workbook(output_path, HIDDEN_COLUMNS)
    print(f"Đã lưu báo cáo tổng hợp: {output_path.resolve()} ({len(merged)} dòng)")

    try:
        # Dung window_end (= "now" thuc su, tra ve tu compute_window), KHONG
        # phai window_start + 1 ngay — nham lan nay se lam upsert ghi de mot
        # khoang thoi gian tuong lai rat lon (loi da tung xay ra khi doi sang
        # cua so "N gio gan nhat" ma quen sua cho nay).
        upsert_split_by_month(merged[FINAL_COLUMNS], window_start, window_end, only_replace_plates=successful_plates)
    except Exception as e:
        print(f"Không đẩy được báo cáo lên Google Sheets: {e}", file=sys.stderr)

    try:
        # Danh dau "done" cho co che nut "Cap nhat ngay" tren web app (neu co
        # ai vua yeu cau) — ghi lai du la lan chay nay tu dong theo lich hay
        # duoc kich hoat thu cong, khong anh huong gi neu khong ai yeu cau.
        sheets_client.set_refresh_status("done", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    except Exception:
        pass

    if not args.keep_raw:
        shutil.rmtree(raw_dir, ignore_errors=True)


if __name__ == "__main__":
    main()

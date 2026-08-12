"""
Logic dung chung de dieu khien he thong Adsun (dinhvi.adsun.vn) bang Playwright.
Da kiem chung thuc te (khong phai doan mo) — xem ghi chu trong tung ham.
"""

import os
import re
import sys
import tempfile
import time
import uuid
from pathlib import Path

import pandas as pd

LOGIN_URL = "https://dinhvi.adsun.vn/"
REPORT_URL = "https://dinhvi.adsun.vn/bao-cao-doanh-nghiep/bao-cao-chi-tiet-hanh-trinh"

DEFAULT_USER = "Nghia0038"
DEFAULT_PASS = "12312344"

RANGE_PRESETS = {
    "hom-nay": "Hôm nay",
    "hom-qua": "Hôm qua",
    "thang-nay": "Tháng này",
    "2": "2 ngày",
    "3": "3 ngày",
    "7": "7 ngày",
    "14": "14 ngày",
    "31": "31 ngày",
}


def get_credentials():
    user = os.environ.get("ADSUN_USER", DEFAULT_USER)
    pw = os.environ.get("ADSUN_PASS", DEFAULT_PASS)
    return user, pw


def is_visible_safe(locator):
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def find_login_inputs(page):
    """Tim o mat khau (type=password) va o truoc no (username) theo thu tu DOM,
    khong phu thuoc ten/placeholder cu the."""
    page.wait_for_selector('input[type="password"]', timeout=20000)
    inputs = page.locator("input").all()
    pw_index = None
    for i, inp in enumerate(inputs):
        if (inp.get_attribute("type") or "").lower() == "password":
            pw_index = i
            break
    if pw_index is None:
        raise RuntimeError("Không tìm thấy ô mật khẩu trên trang đăng nhập.")

    username_input = None
    for i in range(pw_index - 1, -1, -1):
        t = (inputs[i].get_attribute("type") or "text").lower()
        if t in ("text", "tel", "email") and inputs[i].is_visible():
            username_input = inputs[i]
            break
    if username_input is None:
        raise RuntimeError("Không tìm thấy ô tài khoản trên trang đăng nhập.")

    return username_input, inputs[pw_index]


def login(page, username, password):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    user_input, pass_input = find_login_inputs(page)
    user_input.fill(username)
    pass_input.fill(password)

    try:
        page.get_by_role("button", name=re.compile("đăng nhập", re.I)).first.click(
            timeout=5000
        )
    except Exception:
        pass_input.press("Enter")

    # Dang nhap la AJAX (khong reload trang) nen phai doi o mat khau bien mat,
    # khong the dung wait_for_load_state("networkidle") ngay sau khi bam.
    pw_locator = page.locator('input[type="password"]').first
    for _ in range(40):
        page.wait_for_timeout(500)
        if pw_locator.count() == 0 or not pw_locator.is_visible():
            break
        if page.url != LOGIN_URL:
            break

    if pw_locator.count() > 0 and pw_locator.is_visible():
        page.screenshot(path="login_failed.png")
        raise RuntimeError(
            "Đăng nhập chưa thành công (còn ô mật khẩu trên màn hình). "
            "Xem login_failed.png (sai mật khẩu / có captcha?)."
        )
    page.wait_for_timeout(1500)


def goto_report_page(page):
    page.goto(REPORT_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)


def list_all_plates(page):
    """Doc toan bo danh sach xe tu dropdown 'Thiet bi'. Gia dinh bo loc 'Doi xe'
    dang la 'Tat ca' (mac dinh cua tai khoan nay) nen day la toan bo xe cua cong ty.

    QUAN TRONG: dropdown nay dung danh sach ao (virtualized) — chi render vao
    DOM cac option dang nam trong vung nhin cua khung cuon ben trong, KHONG
    phai toan bo danh sach cung luc. Neu chi doc 1 lan luc vua mo (khong cuon)
    se BI THIEU cac xe nam ngoai vung nhin ban dau (da phat hien thuc te: xe
    '61H15792' bi bo sot vi no nam ngay duoi cac xe hien san). Phai cuon dan
    qua toan bo khung ben trong dropdown, gop het cac ten xe xuat hien, dung
    lai khi cuon them vai lan lien tiep ma khong thay xe moi nao."""
    page.get_by_title("Thiết bị").click(timeout=5000)
    page.locator(".select__option").first.wait_for(state="visible", timeout=10000)

    scroll_container = page.locator(".select__menu > div").first
    plates = []
    seen = set()
    stable_rounds = 0
    for _ in range(60):  # gioi han so vong de tranh treo neu co gi bat thuong
        options = page.locator(".select__option").all()
        new_found = False
        for o in options:
            try:
                t = o.inner_text().strip()
            except Exception:
                continue
            if t and t not in seen:
                seen.add(t)
                plates.append(t)
                new_found = True
        if new_found:
            stable_rounds = 0
        else:
            stable_rounds += 1
        if stable_rounds >= 3:
            break
        try:
            scroll_container.evaluate("el => el.scrollBy(0, 150)")
        except Exception:
            break
        page.wait_for_timeout(150)

    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
    return plates


def select_vehicle(page, plate):
    """QUAN TRONG: tim option co text TRUNG KHOP TUYET DOI voi bien so roi bam
    TRUC TIEP vao do — KHONG chi go tim roi bam Enter "mu quang". Da gap thuc
    te (11-12/8): 50H04889 va 51M65797 lien tuc bi "hoan doi" du lieu cua
    nhau qua NHIEU LAN CHAY RIENG BIET (cach nhau hang gio, khong phai do
    chay song song — khoa quet toan cuc da loai truong hop do) — nhieu kha
    nang do bam Enter truoc khi danh sach da loc xong theo tu vua go, chon
    nham 1 option con sot lai tu lan chon truoc. Bam truc tiep vao dung dong
    chu trung khop tuyet doi loai bo hoan toan rui ro nay."""
    if not plate:
        return
    page.get_by_title("Thiết bị").click(timeout=5000)
    page.locator(".select__option").first.wait_for(state="visible", timeout=10000)
    page.keyboard.type(plate)
    page.wait_for_timeout(800)

    option = page.locator(".select__option", has_text=plate).first
    try:
        option.wait_for(state="visible", timeout=3000)
        if option.inner_text().strip() == plate:
            option.click(timeout=5000)
        else:
            page.keyboard.press("Enter")  # khong khop tuyet doi, danh du phong nhu cu
    except Exception:
        page.keyboard.press("Enter")  # khong tim thay option ro rang, danh du phong nhu cu
    page.wait_for_timeout(500)


def select_date_range(page, range_key):
    preset_text = RANGE_PRESETS[range_key]
    page.get_by_title("Thời gian").click(timeout=5000)
    page.wait_for_timeout(400)
    page.get_by_text(preset_text, exact=True).click(timeout=5000)
    page.wait_for_timeout(400)
    page.get_by_role("button", name="Đồng ý", exact=True).click(timeout=5000)
    page.wait_for_timeout(400)


def view_report(page):
    page.get_by_role("button", name="Xem", exact=True).click(timeout=5000)
    page.wait_for_timeout(500)


def wait_report_loaded(page, timeout_seconds=180):
    """Bao cao lon co the mat vai chuc giay de server xu ly (hien % tien do),
    doi den khi het loading/tien do."""
    for _ in range(timeout_seconds * 2):
        page.wait_for_timeout(500)
        loading = is_visible_safe(page.get_by_text("Đang tải", exact=True))
        progress = is_visible_safe(page.locator("text=/Tiến trình thực hiện/"))
        overlay = is_visible_safe(page.locator(".fixed.h-full.z-\\[100\\]"))
        if not loading and not progress and not overlay:
            return
    raise RuntimeError("Báo cáo tải quá lâu (vượt quá thời gian chờ).")


def has_no_data(page):
    return is_visible_safe(page.get_by_text("Không có dữ liệu"))


def export_current_report(page, timeout_toast_ms=90000, timeout_download_ms=120000):
    """QUAN TRONG: nut xuat Excel phai bam 2 LAN — lan 1 de server xu ly file,
    lan 2 (sau khi hien thong bao 'vui long nhan xuat file lan nua') moi thuc
    su tai file xuong. Da kiem chung thuc te tren he thong."""
    export_btn = page.get_by_title("Export all data to Excel")
    export_btn.wait_for(state="visible", timeout=15000)

    # Voi bao cao ca thang (nhieu dong hon han bao cao 1-2 ngay), trang con
    # hien 1 lop phu mo xam (khac voi overlay ma wait_report_loaded kiem tra)
    # trong luc render xong bang lon — cho no bien mat truoc khi bam, tranh
    # bam qua som roi bi che mat nut.
    try:
        page.locator(".fixed.overflow-y-auto.z-50").wait_for(state="hidden", timeout=60000)
    except Exception:
        pass

    with page.expect_download(timeout=timeout_download_ms) as dl_info:
        export_btn.click(timeout=60000)
        try:
            page.get_by_text("nhấn xuất file lần nữa", exact=False).wait_for(
                state="visible", timeout=timeout_toast_ms
            )
        except Exception:
            pass
        export_btn.click(timeout=60000)

    # Toast "Da xu ly du lieu xong..." con hien vai giay va co the CHE cac nut
    # khac (vd. preset ngay cua xe tiep theo), phai doi no bien mat truoc khi
    # tiep tuc thao tac ke tiep tren trang.
    try:
        page.get_by_text("Đã xử lý dữ liệu xong", exact=False).wait_for(
            state="hidden", timeout=15000
        )
    except Exception:
        pass

    return dl_info.value


def fingerprint_df(df, time_column="Thời điểm", coord_column="Tọa độ"):
    """"Dau van tay" don gian cua 1 DataFrame vua xuat: (so dong, moc thoi gian
    dau, moc thoi gian cuoi, toa do dau). Dung de phat hien truong hop xuat bao
    cao cho 1 xe MOI nhung file tai ve van la du lieu CUA XE NGAY TRUOC DO
    trong vong lap — da gap thuc te NHIEU LAN (vd ngay 3/8, 4-5/8: co lan toi
    7 xe khac nhau bi ghi CUNG 1 du lieu), nhieu kha nang do server Adsun tra
    ve file dang cho xu ly trong hang doi (cua xe truoc) thay vi file MOI vua
    yeu cau, du giao dien da hien thi dung xe/nut xuat da bam thanh cong. 2 xe
    that (dinh vi doc lap) khong the co dung dau van tay nay — trung tuyet doi
    la dau hieu chac chan cua loi tren."""
    if df is None or df.empty or time_column not in df.columns:
        return None
    coord = df[coord_column].iloc[0] if coord_column in df.columns and len(df) else None
    return (len(df), str(df[time_column].iloc[0]), str(df[time_column].iloc[-1]), str(coord))


def export_vehicle_report(page, plate, range_preset, header_row_index, seen_fingerprints=None, max_attempts=3):
    """Chon xe + khoang ngay + xem + xuat Excel + doc thanh DataFrame, voi 2
    lop bao ve chong loi thuc te da gap NHIEU LAN khi quet lien tuc nhieu xe
    trong 1 phien trinh duyet (dac biet cac xe cuoi danh sach, sau khi da
    chuyen qua ~10 xe — co ve trang/trinh duyet bi cham/qua tai dan):

    1. Loi ky thuat (vd dropdown "Thiết bị" khong mo kip, timeout cho o
       trang tai) — bat Exception BAT KY trong ca chuoi chon xe/xem/xuat, tai
       lai trang va thu lai (KHONG de loi lam hong ca lan quet).
    2. "Dinh du lieu xe truoc" — server Adsun tra ve file dang xu ly trong
       hang doi (CUA 1 XE BAT KY DA QUET TRUOC DO trong lan chay nay, KHONG
       nhat thiet la xe NGAY LIEN TRUOC) thay vi file MOI vua yeu cau, du
       giao dien da hien dung xe. Da gap thuc te (5/8): 6 xe khac nhau, KHONG
       lien tiep nhau trong vong lap, deu bi "dinh" dung 1 xe (04889) — kiem
       tra chi xe ngay truoc (1 gia tri) la KHONG DU, phai kiem tra voi TOAN
       BO cac xe da quet thanh cong tu dau lan chay nay (seen_fingerprints,
       1 set duoc TICH LUY va truyen chung xuyen suot vong lap cua ham goi).

    Ca 2 truong hop deu thu lai toi da max_attempts lan (moi lan cho them 1
    chut de trang on dinh) truoc khi chiu thua han (raise loi ro rang, KHONG
    tra ve du lieu sai/nghi ngo de tranh ghi de mat du lieu that cua xe do).

    seen_fingerprints: 1 set (nen dung CHUNG 1 object xuyen suot vong lap
    quet nhieu xe cua ham goi, vi ham nay se TU THEM fingerprint moi vao set
    do truoc khi tra ve thanh cong — khong can ham goi tu quan ly). Neu
    khong truyen (None), tao 1 set rieng cho lan goi nay (khong tich luy).

    Tra ve (DataFrame hoac None neu Adsun xac nhan khong co du lieu, fingerprint)."""
    if seen_fingerprints is None:
        seen_fingerprints = set()
    last_error = None
    for attempt in range(max_attempts):
        try:
            if attempt > 0:
                page.wait_for_timeout(2000 * attempt)  # cho trang "nghi" truoc khi thu lai
            goto_report_page(page)
            select_vehicle(page, plate)
            select_date_range(page, range_preset)
            view_report(page)
            wait_report_loaded(page, timeout_seconds=300)

            if has_no_data(page):
                return None, None

            download = export_current_report(page)
            # Luu vao THU MUC TAM CUA HE THONG (tempfile — vd %TEMP% tren
            # Windows), KHONG luu ngay trong thu muc du an. Da gap thuc te 2
            # lan: (10/8) ten file co dinh theo xe gay dung do 2 tien trinh
            # ghi cung 1 file; (12/8) DU DA doi ten file duy nhat theo pid+
            # uuid, 1 xe (51L97909) van lien tuc bi "PermissionError" o CA 6
            # LAN CHAY RIENG BIET — vi thu muc du an co qua nhieu tien trinh
            # (lich tu dong + refresh_range thu cong) cung ghi/doc file lien
            # tuc khi may dang chay don dap, de bi phan mem diet virus/Windows
            # quet va khoa tam thoi ngay sau khi tao. Thu muc tam he thong it
            # bi anh huong hon vi khong phai noi luu file du lieu du an.
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_file:
                tmp_path = Path(tmp_file.name)
            download.save_as(str(tmp_path))
            try:
                df = pd.read_excel(tmp_path, header=header_row_index)
            finally:
                tmp_path.unlink(missing_ok=True)
        except Exception as e:
            last_error = e
            print(
                f"  [{plate}] Lỗi khi xuất báo cáo ({e}) — thử lại lần "
                f"{attempt + 2}/{max_attempts}...",
                file=sys.stderr,
            )
            continue

        fp = fingerprint_df(df)
        if fp is None or fp not in seen_fingerprints:
            if fp is not None:
                seen_fingerprints.add(fp)
            return df, fp

        last_error = None
        print(
            f"  [{plate}] Nghi ngo bi \"dính\" dữ liệu 1 xe đã quét trước đó (dữ liệu "
            f"xuất ra giống hệt) — thử lại lần {attempt + 2}/{max_attempts}...",
            file=sys.stderr,
        )

    if last_error is not None:
        raise RuntimeError(
            f"Xuất báo cáo cho {plate} liên tục lỗi sau {max_attempts} lần thử: {last_error}"
        )
    raise RuntimeError(
        f"Xuất báo cáo cho {plate} liên tục trùng hệt dữ liệu 1 xe đã quét trước đó sau "
        f"{max_attempts} lần thử (có thể do server Adsun trả file cũ trong "
        "hàng đợi) — bỏ qua xe này lần này để tránh ghi nhầm dữ liệu."
    )


def make_scrape_lock_owner_id():
    """Tao 1 ten dinh danh cho tien trinh hien tai de ghi vao khoa quet Adsun
    toan cuc (xem sheets_client.acquire_scrape_lock) — gom nguon goc (GitHub
    Actions run id neu co, hoac ten may local) + pid + 1 chuoi ngau nhien
    ngan, du de phan biet trong log/khi debug."""
    source = os.environ.get("GITHUB_RUN_ID")
    if source:
        source = f"github-actions-{source}"
    else:
        source = os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or "local"
    return f"{source}-pid{os.getpid()}-{uuid.uuid4().hex[:6]}"


def wait_for_scrape_lock(owner_id, max_wait_seconds=300, poll_seconds=15):
    """Cho toi khi thu duoc khoa quet Adsun toan cuc (xem sheets_client.
    acquire_scrape_lock), toi da max_wait_seconds — vi 2 nguon co the cung
    kich hoat gan nhau (vd nguoi dung bam "Cập nhật ngay" dung luc lich dinh
    ky sap chay), cho 1 chut thay vi bo cuoc ngay se giup lan chay do khong
    bi bo lo hoan toan. Tra ve True neu thu duoc khoa, False neu het gio ma
    van bi khoa (nen bo qua lan chay nay, KHONG dang nhap Adsun)."""
    import sheets_client  # import tre de tranh vong import (sheets_client khong can adsun_common)

    waited = 0
    while True:
        try:
            if sheets_client.acquire_scrape_lock(owner_id):
                return True
        except Exception as e:
            print(f"[scrape_lock] Không kiểm tra được khóa ({e}) — cho chạy tiếp (fail-open).", file=sys.stderr)
            return True
        if waited >= max_wait_seconds:
            return False
        print(
            f"[scrape_lock] Đang có tiến trình khác quét Adsun — chờ {poll_seconds}s rồi thử lại "
            f"({waited}/{max_wait_seconds}s)...",
            file=sys.stderr,
        )
        time.sleep(poll_seconds)
        waited += poll_seconds

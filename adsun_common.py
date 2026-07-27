"""
Logic dung chung de dieu khien he thong Adsun (dinhvi.adsun.vn) bang Playwright.
Da kiem chung thuc te (khong phai doan mo) — xem ghi chu trong tung ham.

BAN NAY DUNG CHO GITHUB ACTIONS (repo cong khai) — KHONG co mat khau mac
dinh hardcode nhu ban chay local, bat buoc phai truyen qua bien moi truong
ADSUN_USER / ADSUN_PASS (xem .github/workflows/adsun_scrape.yml, lay tu
GitHub Secrets) de tranh lo mat khau that trong code cong khai.
"""

import os
import re

LOGIN_URL = "https://dinhvi.adsun.vn/"
REPORT_URL = "https://dinhvi.adsun.vn/bao-cao-doanh-nghiep/bao-cao-chi-tiet-hanh-trinh"

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
    user = os.environ.get("ADSUN_USER")
    pw = os.environ.get("ADSUN_PASS")
    if not user or not pw:
        raise RuntimeError(
            "Thiếu biến môi trường ADSUN_USER / ADSUN_PASS. "
            "Cần cấu hình trong GitHub Secrets của repo."
        )
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
    if not plate:
        return
    page.get_by_title("Thiết bị").click(timeout=5000)
    page.locator(".select__option").first.wait_for(state="visible", timeout=10000)
    page.keyboard.type(plate)
    page.wait_for_timeout(800)
    page.keyboard.press("Enter")
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

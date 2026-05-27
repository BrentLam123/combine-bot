import re
import calendar
from datetime import datetime, timedelta
import pandas as pd
import openpyxl  # <-- Đã thêm thư viện openpyxl
import math
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import subprocess
import os

# ── Timestamp print ──
import sys, io
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
_orig_print = print
def print(*args, **kwargs):
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    _orig_print(f"[{ts}]", *args, **kwargs)

# ── Flag bật/tắt timing chi tiết ──
ENABLE_TIMING = True  # Đổi thành False để tắt khi không cần nữa

# ==========================================
# TỰ ĐỘNG MỞ EDGE 9522 (nếu chưa mở)
# ==========================================
import socket

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def open_edge_9522():
    """Mở Edge với remote debugging port 9522 nếu chưa có"""
    if is_port_in_use(9522):
        print("[HỆ THỐNG] Edge ONE đã mở sẵn (port 9522). Bỏ qua lệnh khởi động trình duyệt.")
        return
    print("[HỆ THỐNG] Edge ONE chưa mở. Đang tự động khởi động...")
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    args = [
        edge_path,
        "--remote-debugging-port=9522",
        r"--user-data-dir=C:\edge_one",
        "--disable-background-timer-throttling",
        "--disable-renderer-backgrounding",
        "--disable-backgrounding-occluded-windows"
    ]
    try:
        subprocess.Popen(args)
        print("[OK] Đã mở Edge 9522 thành công!")
        time.sleep(3)
    except FileNotFoundError:
        print("[ERROR] Không tìm thấy Edge tại đường dẫn trên. Kiểm tra lại đường dẫn!")

open_edge_9522()

# ==========================================
# CẤU HÌNH TRÌNH DUYỆT & HẰNG SỐ
# ==========================================
edge_options = Options()
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9522")
driver = webdriver.Edge(options=edge_options)
wait   = WebDriverWait(driver, 15)

SPEED      = 0.1
ONE_URL    = "https://ecomm.one-line.com/one-ecom/prices/one-quote-booking"
FORM_XPATH = "(//input[@placeholder='Please search location'])[1]"

PORT_ALIASES = {
    "BUSAN":       "PUSAN",
    "HO CHI MINH": "HO CHI MINH",
    "ANTWERP":     "ANTWERP"
}
# Danh sách POD theo khu vực để xác định remark phụ
CHINA_PORTS    = ["CHINA", "CN", "SHANGHAI", "NINGBO", "GUANGZHOU", "SHENZHEN",
                  "TIANJIN", "QINGDAO", "XIAMEN", "DALIAN", "BEIJING"]
JAPAN_PORTS    = ["JAPAN", "JP", "TOKYO", "OSAKA", "NAGOYA", "YOKOHAMA",
                  "KOBE", "HAKATA", "MOJI"]
EUROPE_PORTS   = ["GERMANY", "DE", "FRANCE", "FR", "NETHERLANDS", "NL",
                  "BELGIUM", "BE", "SPAIN", "ES", "ITALY", "IT", "POLAND", "PL",
                  "SWEDEN", "SE", "DENMARK", "DK", "FINLAND", "FI", "NORWAY", "NO",
                  "PORTUGAL", "PT", "GREECE", "GR", "TURKEY", "TR", "UK", "GB",
                  "ROTTERDAM", "HAMBURG", "ANTWERP", "BARCELONA", "GDANSK",
                  "FELIXSTOWE", "SOUTHAMPTON", "LE HAVRE", "MARSEILLE",
                  "GENOA", "VALENCIA", "BREMERHAVEN", "PIRAEUS"]

# ==========================================
# CÁC HÀM HỖ TRỢ CƠ BẢN
# ==========================================

SPEED = 0.01  # Bạn có thể đổi thành 0.1, 0.2 tùy ý

def wait_and_speed(xpath, timeout=10):
    """Đợi phần tử sẵn sàng bằng WebDriverWait + nghỉ thêm SPEED"""
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    time.sleep(SPEED) # Nghỉ theo ý bạn sau khi web đã load xong
    return element

def click_icon_add(timeout=10):
    """
    Đợi nút icon-add xuất hiện rồi click. Tránh IndexError khi React
    chưa render xong sau add_equipment.
    """
    last_err = None
    for try_i in range(3):
        try:
            WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//img[@alt='icon-add']")
                )
            )
            icons = driver.find_elements(By.XPATH, "//img[@alt='icon-add']")
            if icons:
                # Scroll đến nút cuối cùng để chắc chắn nó visible
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", icons[-1]
                )
                time.sleep(0.2)
                js_click(icons[-1])
                return True
        except Exception as e:
            last_err = e
        # Chờ React render thêm
        time.sleep(0.6)
    raise Exception(f"Không tìm thấy nút icon-add sau {3*timeout}s (last err: {last_err})")

def get_valid_date(etd_dates):
    latest_etd = max(etd_dates)
    day = latest_etd.day
    if day <= 7:    valid_day = 7
    elif day <= 14: valid_day = 14
    elif day <= 21: valid_day = 21
    else:           valid_day = calendar.monthrange(latest_etd.year, latest_etd.month)[1]
    return datetime(latest_etd.year, latest_etd.month, valid_day).strftime("%d-%b")

def js_click(element):
    driver.execute_script("arguments[0].click();", element)

def select_port(input_index, port_name, country_name):
    actual_port = PORT_ALIASES.get(str(port_name).upper(), str(port_name))

    # Lọc country_str
    country_raw = str(country_name).strip()
    if (country_raw.upper() == "NAN"
            or not country_raw
            or re.match(r'\d{4}-\d{2}-\d{2}', country_raw)
            or re.match(r'\d{2}/\d{2}/\d{4}', country_raw)):
        country_str = ""
    else:
        country_str = country_raw.upper()

    port_input_xpath = f"(//input[@placeholder='Please search location'])[{input_index}]"
    port_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, port_input_xpath))
    )

    # Đọc giá trị hiện tại của ô input
    current_val = port_input.get_attribute("value") or ""
    if actual_port.upper() in current_val.upper():
        print(f"  ✅ Cảng {actual_port} đã có sẵn, bỏ qua bước nhập.")
        return # Thoát hàm luôn, không cần tìm kiếm

    print(f"  👉 Nhập cảng: {actual_port} ({country_str if country_str else 'no country filter'})")

    port_input.click()
    driver.execute_script("""
        var el = arguments[0];
        var nativeSetter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        nativeSetter.call(el, '');
        el.dispatchEvent(new Event('input', {bubbles:true}));
    """, port_input)
    port_input.send_keys(actual_port)


    # Đợi dropdown xuất hiện
    try:
        WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.XPATH, "//li[@role='option']"))
        )
    except:
        time.sleep(0.5)
    all_options = driver.find_elements(By.XPATH, "//li[@role='option']")
    # ── RETRY: nếu không có option khớp tên cảng → xóa 1 ký tự cuối rồi thử lại ──
    def has_matching_option(opts, name):
        return any(name.upper() in opt.text.strip().upper() for opt in opts)

    if not has_matching_option(all_options, actual_port):
        print(f"  🔄 Không thấy option '{actual_port}', thử xóa 1 ký tự cuối...")
        trimmed = actual_port[:-1]
        port_input.send_keys(Keys.BACKSPACE)
        try:
            WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.XPATH, "//li[@role='option']"))
            )
        except:
            time.sleep(0.5)
        all_options = driver.find_elements(By.XPATH, "//li[@role='option']")
        if not has_matching_option(all_options, trimmed):
            print(f"  ❌ Vẫn không thấy option sau khi xóa ký tự → bỏ qua")
        else:
            print(f"  ✅ Tìm thấy option sau khi trim thành '{trimmed}' (dùng làm key tìm kiếm tiếp)")
            # FIX: dùng trimmed làm key cho VÒNG 1/2/3 phía dưới, nếu không các vòng đó
            # vẫn tìm theo 'PARADIP'/'VISAKHAPATNAM' nguyên gốc và không match được.
            actual_port = trimmed

    # ── HÀM NỘI BỘ: Đọc badge CY/DOOR trực tiếp từ element con ──
    def get_badge_text(opt):
        """
        Đọc text của badge (CY / DOOR / CFS...) bên trong option.
        Badge là element con nhỏ nhất, không phải toàn bộ text của li.
        """
        # Tìm tất cả element con có background/color khác biệt (badge)
        candidates = opt.find_elements(By.XPATH,
            ".//*[string-length(normalize-space(text())) <= 6 "   # badge ngắn: CY, DOOR, CFS
            "and normalize-space(text()) != '']"
        )
        for el in candidates:
            t = el.text.strip().upper()
            if t in ("CY", "DOOR", "CFS", "RAMP"):
                return t
        return ""

    def get_port_name_text(opt):
        """Lấy tên cảng từ option, bỏ badge ra"""
        full = opt.text.strip().upper()
        for badge in ("CY", "DOOR", "CFS", "RAMP"):
            full = full.replace(badge, "").strip()
        return full

    selected = False
    # ── VÒNG 1: khớp tên cảng (+ country nếu có) + badge = CY ──
    for opt in all_options:
        try:
            port_text = get_port_name_text(opt)
            if actual_port.upper() not in port_text:
                continue
            if country_str and country_str not in port_text:
                continue
            badge = get_badge_text(opt)
            if badge == "CY":
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", opt)
                opt.click()
                selected = True
                print(f"  ✅ CY (với country): {opt.text.strip()}")
                break
        except:
            continue

    # ── VÒNG 2: bỏ lọc country, vẫn cần badge = CY ──
    if not selected:
        for opt in all_options:
            try:
                port_text = get_port_name_text(opt)
                if actual_port.upper() not in port_text:
                    continue
                badge = get_badge_text(opt)
                if badge == "CY":
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", opt)
                    opt.click()
                    selected = True
                    print(f"  ✅ CY (no country filter): {opt.text.strip()}")
                    break
            except:
                continue

    # ── VÒNG 3: fallback - chọn đầu tiên có tên cảng (dù badge gì) ──
    if not selected:
        for opt in all_options:
            try:
                if actual_port.upper() in opt.text.strip().upper():
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", opt)
                    opt.click()
                    selected = True
                    print(f"  ⚠️ Không có CY, chọn đầu tiên: {opt.text.strip()}")
                    break
            except:
                continue

    if not selected:
        print(f"  ❌ Không tìm thấy option nào cho {actual_port}. Dùng phím mũi tên...")
        port_input.send_keys(Keys.ARROW_DOWN)
        port_input.send_keys(Keys.ENTER)

    time.sleep(0.2)

def wait_for_loading_popup():
    """Học từ Bot April: Đợi vòng xoay biến mất hoàn toàn"""
    loader_xpath = "//div[contains(@class, 'CarouselLoadingPopup_progress-container') or contains(@class, 'ajax-progress')]"
    try:
        # Đợi nó xuất hiện (trong tối đa 1.5s)
        WebDriverWait(driver, 1.5).until(EC.presence_of_element_located((By.XPATH, loader_xpath)))
        # Đợi nó biến mất (trong tối đa 30s)
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.XPATH, loader_xpath)))
        time.sleep(SPEED) # Web xong rồi, nghỉ SPEED theo ý bạn
    except TimeoutException:
        pass # Nếu không có popup thì thôi chạy tiếp

def add_equipment(index, equip_type, weight, qty_clicks=0):
    equip_input_xpath = f"(//input[@placeholder='Select an Equipment Type'])[{index}]"
    equip_input = wait.until(EC.presence_of_element_located((By.XPATH, equip_input_xpath)))
    
    # Check Loại Container
    current_type = equip_input.get_attribute("value") or ""
    if current_type.strip() != equip_type:
        print(f"  📦 Container {index}: Đang chọn {equip_type}...")
        driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", equip_input)
        time.sleep(SPEED)
        js_click(equip_input)
        time.sleep(0.3)
        equip_options = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//li[@role='option']")))
        for opt in equip_options:
            if opt.text.strip() == equip_type:
                js_click(opt)
                break
        time.sleep(SPEED)
    
    # Check Số lượng (Qty)
    inc_btn_xpath = f"(//button[@data-action='increment'])[{index}]"
    inc_btn = wait.until(EC.presence_of_element_located((By.XPATH, inc_btn_xpath)))
    qty_input = inc_btn.find_element(By.XPATH, "./parent::div//input")
    if qty_input.get_attribute("value") != "1":
        js_click(qty_input)
        time.sleep(SPEED)
        qty_input.send_keys(Keys.CONTROL + "a")
        qty_input.send_keys(Keys.BACKSPACE)
        qty_input.send_keys("1")
        time.sleep(SPEED)

    # Check Khối lượng (Weight)
    weight_input_xpath = f"(//input[@placeholder='0' and @inputmode='numeric'])[{index}]"
    weight_input = wait.until(EC.presence_of_element_located((By.XPATH, weight_input_xpath)))
    current_weight = weight_input.get_attribute("value") or ""
    if current_weight.replace(",", "") != weight:
        print(f"  ⚖️ Cập nhật khối lượng thành {weight}kg")
        js_click(weight_input)
        time.sleep(SPEED)
        weight_input.send_keys(Keys.CONTROL + "a")
        weight_input.send_keys(Keys.BACKSPACE)
        weight_input.send_keys(weight)
        time.sleep(SPEED)
    else:
        print(f"  ✅ Container {index} ({equip_type} - {weight}kg) đã cấu hình chuẩn.")

# ==========================================
# HÀM ĐĂNG NHẬP (dùng lại ở nhiều nơi)
# ==========================================
def do_login():
    """Đăng nhập nhanh bằng JavaScript inject giá trị trực tiếp"""
    print("  🔐 Đang đăng nhập (nhanh)...")
    
    user_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='text' or @name='username' or @id='username']"))
    )
    pass_input = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
    )

    # Nhập username bằng JS (nhanh nhất)
    driver.execute_script("arguments[0].value = arguments[1];", user_input, "PIOLOG")
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('input', {bubbles:true})); "
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
        user_input
    )

    # Nhập password bằng JS
    driver.execute_script("arguments[0].value = arguments[1];", pass_input, "Vankiep@21")
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('input', {bubbles:true})); "
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
        pass_input
    )

    time.sleep(0.3)  # Đợi React cập nhật state

    login_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@id='btn-login']"))
    )
    js_click(login_btn)
    print("  ⏳ Đã bấm Login, đợi form nhập liệu...")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, FORM_XPATH)))
    print("  ✅ Đăng nhập thành công!")

# ==========================================
# MỞ TAB MỚI + TỰ ĐỘNG ĐĂNG NHẬP NẾU CẦN
# ==========================================
ONE_URL = "https://ecomm.one-line.com/one-ecom/prices/one-quote-booking"
LOGIN_URL_PREFIX = "https://auth.one-line.com/login"

def open_tab_and_ensure_ready():
    """
    Mở tab mới → vào ONE_URL.
    Check URL hiện tại:
      - URL = ONE_URL  → session còn, vào luôn
      - URL chứa auth.one-line.com/login → đăng nhập rồi vào form
    """
    driver.switch_to.new_window('tab')
    handle = driver.current_window_handle
    driver.get(ONE_URL)

    # Đợi trang load xong (tối đa 10s)
    time.sleep(2)

    # CHECK URL để quyết định hành động
    current_url = driver.current_url

    if ONE_URL in current_url:
        # Đang ở trang check giá → vào luôn
        try:
            WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, FORM_XPATH))
            )
            print("  ✅ Session còn hiệu lực! URL đúng trang check giá.")
        except:
            print("  ⚠️ URL đúng nhưng form chưa hiện, đợi thêm...")
            time.sleep(3)
        return handle

    elif LOGIN_URL_PREFIX in current_url:
        # Đang ở trang đăng nhập
        print("  ⚠️ Bị redirect trang login. Tiến hành đăng nhập...")
        try:
            do_login()
        except Exception as e:
            print(f"  ❌ Lỗi đăng nhập: {e}")
        return handle

    else:
        # URL lạ khác → thử đăng nhập phòng hờ
        print(f"  ⚠️ URL không xác định: {current_url}")
        print("  🔄 Thử navigate lại trang check giá...")
        driver.get(ONE_URL)
        time.sleep(2)
        current_url = driver.current_url
        if LOGIN_URL_PREFIX in current_url:
            try:
                do_login()
            except Exception as e:
                print(f"  ❌ Lỗi đăng nhập: {e}")
        return handle

# ==========================================
# NHẬP POL/POD SAU RELOAD (có check login)
# ==========================================
def fill_tab_ports(row):
    """
    Đợi form trống sau reload rồi nhập POL/POD.
    Nếu thay vào đó thấy trang login (session hết) → tự đăng nhập trước.
    """
    country = str(row[0]).strip()
    pol     = str(row[2]).strip()
    pod     = str(row[3]).strip()

    # Đợi form nhập liệu
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, FORM_XPATH)))
    except TimeoutException:
        # Session có thể đã hết sau reload
        print("  ⚠️ Sau reload không thấy form → kiểm tra đăng nhập lại...")
        try:
            do_login()
        except Exception as e:
            print(f"  ❌ Không thể đăng nhập lại: {e}")
            return

    select_port(1, pol, country)
    select_port(2, pod, country)


# ==========================================
# SWITCH TAB TRICK → ĐỢI HẾT POPUP LOADING
# ==========================================
def switch_tab_trick_until_clear(tabs, rounds=5, pause=0.2):
    print("🔄 Switch tab ép React render + đợi hết popup loading...")
    loader_xpath = "//div[contains(@class, 'CarouselLoadingPopup_progress-container') or contains(@class, 'ajax-progress')]"

    # Phase 1: đảo nhanh để React kịp văng popup lên màn hình
    for _ in range(rounds):
        for tab in tabs:
            driver.switch_to.window(tab)
            time.sleep(pause)

    # Phase 2: từng tab, đợi popup biến mất hẳn
    print("⏳ Đang chờ popup tắt trên từng tab...")
    for i, tab in enumerate(tabs):
        driver.switch_to.window(tab)
        try:
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, loader_xpath)))
            WebDriverWait(driver, 45).until(EC.invisibility_of_element_located((By.XPATH, loader_xpath)))
            print(f"  ✅ Tab {i+1}: popup đã tắt")
        except TimeoutException:
            pass  # Không có popup hoặc đã tắt rồi → OK

    # FIX Phase 3: Xác nhận từng tab đã load xong bằng cách đợi form equipment HOẶC thông báo lỗi
    # → tránh tình trạng popup dùng class khác mà Phase 2 không phát hiện được
    print("🔎 Xác nhận trang đã load xong (chờ form equipment hoặc error msg)...")
    PAGE_READY_XPATH = (
        "//input[@placeholder='Select an Equipment Type']"
        " | //div[contains(@class,'RouteInput_wrap-error-loading')]"
        " | //p[contains(text(),'No port pair')]"
    )
    for i, tab in enumerate(tabs):
        driver.switch_to.window(tab)
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, PAGE_READY_XPATH))
            )
            print(f"  ✅ Tab {i+1}: form sẵn sàng")
        except TimeoutException:
            # Thử thêm 1 lần switch tab cuối để kick React
            for other_tab in tabs:
                if other_tab != tab:
                    driver.switch_to.window(other_tab)
                    time.sleep(0.3)
                    break
            driver.switch_to.window(tab)
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, PAGE_READY_XPATH))
                )
                print(f"  ✅ Tab {i+1}: form sẵn sàng (sau retry)")
            except TimeoutException:
                print(f"  ⚠️ Tab {i+1}: timeout sau 45s — tiếp tục (sẽ handle trong scrape_tab)")

    print("✅ Tất cả tab sạch popup!")


# ==========================================
# RELOAD ĐỒNG LOẠT (KHÔNG CHỜ NHAU)
# ==========================================
def reload_all_tabs_simultaneously(tabs):
    print(f"🔄 Gửi lệnh Reload đồng loạt {len(tabs)} tab...")
    for i, tab in enumerate(tabs):
        driver.switch_to.window(tab)
        driver.refresh()  # Không đợi, sang tab kế ngay → tất cả reload song song
        print(f"  ↩️  Tab {i+1}: đã gửi lệnh reload")
    print("✅ Đã gửi reload cho toàn bộ tabs!")


# ==========================================
# BÓC GIÁ MỘT TAB (driver đã switch sẵn)
# ==========================================
def scrape_tab(row_data):
    pol_name = str(row_data[2]).strip().upper()
    
    # ── KIỂM TRA "No port pair" BẰNG ĐÚNG ELEMENT CỦA WEB ──
    # Đợi web phản hồi sau khi switch tab trick load xong
    NO_PAIR_XPATH = "//div[contains(@class,'RouteInput_wrap-error-loading')]"
    
    try:
        # Đợi tối đa 2s xem có hiện thông báo lỗi không
        WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, NO_PAIR_XPATH))
        )
        # Nếu thấy → không có giá
        print(f"  ⚠️ Web báo 'No port pair available' → trả về '-'")
        return {
            "POL": row_data[2], "POD": row_data[3], "Status": "No port pair",
            "20 DRY": "-", "40 DRY": "-", "40 HC": "-",
            "ETD": "-", "Transit Time": "-", "Valid": "-",
            "Remark": "-", "Free Time": "-",
            "Vessel": "-", "Transshipment": "-"
        }
    except TimeoutException:
        pass  # Không có thông báo lỗi → có thể có giá, chạy tiếp bình thường

    # ── Kiểm tra thêm: ô chọn container có hiện không ──
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@placeholder='Select an Equipment Type']")
            )
        )
    except TimeoutException:
        print(f"  ⚠️ Không thấy ô chọn container → trả về '-'")
        return {
            "POL": row_data[2], "POD": row_data[3], "Status": "No container form",
            "20 DRY": "-", "40 DRY": "-", "40 HC": "-",
            "ETD": "-", "Transit Time": "-", "Valid": "-",
            "Remark": "-", "Free Time": "-",
            "Vessel": "-", "Transshipment": "-"
        }

    # ── Nhập container bình thường ──
    add_equipment(1, "DRY 20", "22222")
    
    # Chỉ bấm Add nếu ô thứ 2 chưa tồn tại (giải quyết triệt để lỗi khi tab được Retry)
    if len(driver.find_elements(By.XPATH, "(//input[@placeholder='Select an Equipment Type'])")) < 2:
        try: click_icon_add(); time.sleep(SPEED)
        except Exception: pass
    add_equipment(2, "DRY 40", "22222")
    
    # Chỉ bấm Add nếu ô thứ 3 chưa tồn tại
    if len(driver.find_elements(By.XPATH, "(//input[@placeholder='Select an Equipment Type'])")) < 3:
        try: click_icon_add(); time.sleep(SPEED)
        except Exception: pass
    add_equipment(3, "DRY 40H", "22222")    

    # ── CHỐT CHẶN: ÉP CHẮC CHẮN 3 Ô SỐ LƯỢNG LÀ 1 TRƯỚC KHI ĐI TIẾP ──
    driver.execute_script("""
        let qty_inputs = document.querySelectorAll('button[data-action="increment"]');
        qty_inputs.forEach(btn => {
            let input = btn.parentElement.querySelector('input');
            if(input) {
                // Ép form React nhận giá trị 1 một cách triệt để
                let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
                nativeInputValueSetter.call(input, '1');
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
            }
        });
    """)
    time.sleep(0.5)

# ── Nhập Commodity (Check trước khi nhập) ──
    fast_wait = WebDriverWait(driver, 1.5)
    try:
        commodity_input = fast_wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='searchCommodityByName']")))
        current_commodity = commodity_input.get_attribute("value") or ""
        
        # Nếu chưa có mã 6211 hoặc chữ TRACK SUITS thì mới nhập
        if "6211" not in current_commodity and "TRACK SUITS" not in current_commodity.upper():
            driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", commodity_input)
            time.sleep(SPEED)
            fast_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='searchCommodityByName']"))).click()
            time.sleep(SPEED)
            commodity_input.send_keys(Keys.CONTROL + "a"); commodity_input.send_keys(Keys.BACKSPACE)
            commodity_input.send_keys("6211")
            js_click(fast_wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and contains(., 'TRACK SUITS')]"))))
            time.sleep(SPEED)
        else:
            print("  ✅ Commodity đã có sẵn, bỏ qua bước nhập.")
    except TimeoutException:
        raise Exception("WEB_LAG_RETRY: Commodity load quá chậm (> 1.5s)")

    # Chọn ngày tím đầu tiên còn chỗ (bỏ qua Sold out)
    calendar_opened = False

    # Danh sách XPath ưu tiên để tìm ô lịch tàu (tách riêng, không dùng | trong WebDriverWait)
    CALENDAR_XPATHS = [
        "/html/body/div[1]/main/div[2]/div[2]/main/div[2]/div[2]/div/div[2]/div/div[2]/div[4]/div[3]/div/div/div/div/div/input",
        "//input[@placeholder='Please select vessel departure date at origin']",
        "//div[contains(@class,'date-picker')]//input",
    ]

    for attempt in range(1): # Ép chạy 1 lần, lỗi thì ném ra ngoài để retry tab sau
        try:
            date_input = None
            for xp in CALENDAR_XPATHS:
                try:
                    date_input = WebDriverWait(driver, 1.5).until( # Rút xuống 1.5s
                        EC.presence_of_element_located((By.XPATH, xp))
                    )
                    if date_input:
                        print(f"  📅 Tìm thấy ô lịch tàu bằng XPath: {xp[:60]}...")
                        break
                except TimeoutException:
                    continue

            if date_input is None:
                raise Exception("WEB_LAG_RETRY: Không tìm được ô input lịch tàu (> 1.5s)")

            driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", date_input)
            time.sleep(0.5)

            try:
                ActionChains(driver).move_to_element(date_input).click().perform()
            except Exception:
                driver.execute_script("arguments[0].click();", date_input)

            CALENDAR_READY_XPATH = (
                "//div[contains(@class,'date-picker-date-highlight') and @role='option' and @aria-disabled='false']"
            )
            try:
                WebDriverWait(driver, 1.5).until( # Rút xuống 1.5s
                    EC.presence_of_element_located((By.XPATH, CALENDAR_READY_XPATH))
                )
                print(f"  ✅ Calendar đã render xong ngày có giá!")
            except TimeoutException:
                # Nếu lag, văng lỗi ngay lập tức để chuyển tab
                raise Exception("WEB_LAG_RETRY: Lịch tàu chưa hiện ngày tím (> 1.5s)")

            # ── DEBUG: dump TẤT CẢ date cells để xem class thực tế ──
            all_date_cells = driver.find_elements(By.XPATH,
                "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month'))]"
            )
            print(f"  [DEBUG] Tổng date cells (enabled, in-month): {len(all_date_cells)}")
            for i, dc in enumerate(all_date_cells[:15]):
                try:
                    cls = dc.get_attribute("class") or ""
                    lbl = (dc.get_attribute("aria-label") or "")[:60]
                    txt = dc.text.strip().replace('\n', ' ')[:40]
                    print(f"  [DEBUG]  [{i}] class='{cls}' | label='{lbl}' | text='{txt}'")
                except Exception:
                    pass

            # Tìm ngày có giá: ưu tiên highlight, fallback sang tất cả enabled date cells có giá
            HIGHLIGHT_XPATH = "//div[contains(@class,'date-picker-date-highlight') and @role='option' and @aria-disabled='false' and not(contains(@class,'outside-month'))]"
            highlight_dates = driver.find_elements(By.XPATH, HIGHLIGHT_XPATH)
            print(f"  🔍 Số ngày 'highlight': {len(highlight_dates)}")

            # Nếu ít highlight → thử tìm ngày có giá bằng cách rộng hơn
            if len(highlight_dates) <= 1:
                # Thử tìm ngày có text giá (chứa "K" như 8.6K, 9.5K) hoặc có class khác
                broader_xpaths = [
                    "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month')) and (contains(@class,'highlight') or contains(@class,'in-range') or contains(@class,'available') or contains(@class,'selectable'))]",
                    "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month')) and not(contains(@class,'disabled'))]",
                ]
                for bxp in broader_xpaths:
                    broader_dates = driver.find_elements(By.XPATH, bxp)
                    # Lọc chỉ những ngày có text chứa giá (K = nghìn)
                    dates_with_price = [d for d in broader_dates if re.search(r'\d+\.?\d*K', d.text or "")]
                    if dates_with_price:
                        print(f"  🔍 Tìm thêm {len(dates_with_price)} ngày có giá (broader search)")
                        highlight_dates = dates_with_price
                        break

            print(f"  🔍 Tổng ngày có giá: {len(highlight_dates)}")

            if not highlight_dates:
                # FIX: cho thêm 2s rồi đọc lại 1 lần nữa, nhiều khi ngày chưa kịp populate
                print(f"  ⏳ Chưa thấy ngày có giá, chờ thêm 2s rồi đọc lại...")
                time.sleep(2)
                highlight_dates = driver.find_elements(By.XPATH, HIGHLIGHT_XPATH)
                if len(highlight_dates) <= 1:
                    all_enabled = driver.find_elements(By.XPATH,
                        "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month')) and not(contains(@class,'disabled'))]"
                    )
                    dates_with_price = [d for d in all_enabled if re.search(r'\d+\.?\d*K', d.text or "")]
                    if dates_with_price:
                        highlight_dates = dates_with_price
                print(f"  🔍 Lần đọc lại: {len(highlight_dates)} ngày có giá")

            if not highlight_dates:
                print(f"  ℹ️ Calendar đã mở nhưng không có ngày có giá → thử attempt khác")
                try:
                    date_input.send_keys(Keys.ESCAPE)
                except Exception:
                    pass
                time.sleep(1)
                calendar_opened = False
                continue  # FIX: retry thay vì break luôn

            # ── Lưu aria-label thay vì element để tránh stale ──
            chosen_aria_label = None
            chosen_date_str = ""
            min_etd_date = datetime.today().date() + timedelta(days=6)

            for hd in highlight_dates:
                try:
                    aria_label = hd.get_attribute("aria-label") or ""
                    if "sold out" in aria_label.lower():
                        continue
                    m = re.search(r'Choose \w+,\s+(\w+)\s+(\d+)\w*,\s+(\d+)', aria_label)
                    if not m:
                        continue
                    month_str = m.group(1)
                    day_num   = int(m.group(2))
                    year_num  = int(m.group(3))
                    etd_date  = datetime.strptime(
                        f"{day_num} {month_str} {year_num}", "%d %B %Y"
                    ).date()
                    if etd_date < min_etd_date:
                        print(f"  ⏭️ Bỏ qua {etd_date} (< today+6={min_etd_date})")
                        continue
                    # ── Chỉ lưu aria-label, KHÔNG lưu element ──
                    chosen_aria_label = aria_label
                    chosen_date_str = etd_date.strftime('%d-%b')
                    print(f"  ✅ Chọn ngày còn chỗ: {chosen_date_str}")
                    break
                except Exception:
                    continue

            if chosen_aria_label is None:
                # Thử chờ thêm 3s cho tháng sau render giá (async loading)
                print(f"  ⏳ Chưa có ngày hợp lệ, chờ thêm 3s cho tháng sau load giá...")
                time.sleep(3)

                # Tìm lại với broader search
                highlight_dates_2 = driver.find_elements(By.XPATH, HIGHLIGHT_XPATH)
                if len(highlight_dates_2) <= 1:
                    all_en = driver.find_elements(By.XPATH,
                        "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month')) and not(contains(@class,'disabled'))]"
                    )
                    dp = [d for d in all_en if re.search(r'\d+\.?\d*K', d.text or "")]
                    if dp:
                        highlight_dates_2 = dp
                        print(f"  🔍 Sau 3s: {len(highlight_dates_2)} ngày có giá (broader)")
                    else:
                        print(f"  🔍 Sau 3s: {len(highlight_dates_2)} ngày highlight")
                else:
                    print(f"  🔍 Sau 3s: {len(highlight_dates_2)} ngày highlight")

                for hd in highlight_dates_2:
                    try:
                        aria_label = hd.get_attribute("aria-label") or ""
                        if "sold out" in aria_label.lower():
                            continue
                        m = re.search(r'Choose \w+,\s+(\w+)\s+(\d+)\w*,\s+(\d+)', aria_label)
                        if not m: continue
                        etd_date = datetime.strptime(
                            f"{int(m.group(2))} {m.group(1)} {int(m.group(3))}", "%d %B %Y"
                        ).date()
                        if etd_date < min_etd_date: continue
                        chosen_aria_label = aria_label
                        chosen_date_str = etd_date.strftime('%d-%b')
                        print(f"  ✅ Chọn ngày còn chỗ (sau wait): {chosen_date_str}")
                        break
                    except Exception: continue

            # Nếu vẫn chưa có → thử navigate sang tháng sau
            if chosen_aria_label is None:
                print(f"  ⏳ Thử chuyển sang tháng sau trên calendar...")
                next_month_clicked = False
                for next_xp in [
                    "//button[contains(@class,'next') and contains(@class,'month')]",
                    "//button[@aria-label='Next month']",
                    "//span[contains(@class,'next-icon')]/ancestor::button",
                    "//div[contains(@class,'date-picker')]//button[last()]",
                ]:
                    try:
                        next_btn = driver.find_element(By.XPATH, next_xp)
                        next_btn.click()
                        next_month_clicked = True
                        print(f"  ✅ Đã click next month")
                        break
                    except Exception: continue

                if next_month_clicked:
                    time.sleep(3)
                    # DEBUG: dump cells sau next month
                    all_cells_nm = driver.find_elements(By.XPATH,
                        "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month'))]"
                    )
                    print(f"  [DEBUG] Sau next month: {len(all_cells_nm)} date cells")
                    for i, dc in enumerate(all_cells_nm[:10]):
                        try:
                            cls = dc.get_attribute("class") or ""
                            txt = dc.text.strip().replace('\n', ' ')[:40]
                            print(f"  [DEBUG]  [{i}] class='{cls}' | text='{txt}'")
                        except Exception: pass

                    highlight_dates_3 = driver.find_elements(By.XPATH, HIGHLIGHT_XPATH)
                    if len(highlight_dates_3) <= 1:
                        all_en = driver.find_elements(By.XPATH,
                            "//div[@role='option' and @aria-disabled='false' and not(contains(@class,'outside-month')) and not(contains(@class,'disabled'))]"
                        )
                        dp = [d for d in all_en if re.search(r'\d+\.?\d*K', d.text or "")]
                        if dp:
                            highlight_dates_3 = dp
                            print(f"  🔍 Tháng sau: {len(highlight_dates_3)} ngày có giá (broader)")
                        else:
                            print(f"  🔍 Tháng sau: {len(highlight_dates_3)} ngày highlight")
                    else:
                        print(f"  🔍 Tháng sau: {len(highlight_dates_3)} ngày highlight")

                    for hd in highlight_dates_3:
                        try:
                            aria_label = hd.get_attribute("aria-label") or ""
                            if "sold out" in aria_label.lower(): continue
                            m = re.search(r'Choose \w+,\s+(\w+)\s+(\d+)\w*,\s+(\d+)', aria_label)
                            if not m: continue
                            etd_date = datetime.strptime(
                                f"{int(m.group(2))} {m.group(1)} {int(m.group(3))}", "%d %B %Y"
                            ).date()
                            if etd_date < min_etd_date: continue
                            chosen_aria_label = aria_label
                            chosen_date_str = etd_date.strftime('%d-%b')
                            print(f"  ✅ Chọn ngày còn chỗ (tháng sau): {chosen_date_str}")
                            break
                        except Exception: continue

            if chosen_aria_label is None:
                print(f"  ℹ️ Tất cả ngày tím đều Sold out hoặc quá sớm → không có lịch khả dụng")
                date_input.send_keys(Keys.ESCAPE)
                calendar_opened = False
                break

            # ── Tìm lại element TƯƠI bằng aria-label rồi mới click ──
            escaped_label = chosen_aria_label.replace('"', '\\"')
            fresh_day = driver.find_elements(By.XPATH,
                f'//div[@role="option" and @aria-label="{escaped_label}" and not(contains(@class,"outside-month"))]'
            )
            if not fresh_day:
                print(f"  ⚠️ Không tìm lại được ngày {chosen_date_str}")
                date_input.send_keys(Keys.ESCAPE)
                calendar_opened = False
                break

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", fresh_day[0]
            )
            time.sleep(0.3)
            
            # Dòng Debug: Báo cáo ngày chọn gần nhất trước khi click
            print(f"  🐛 [DEBUG] Chuẩn bị click chọn ngày: {chosen_date_str} (aria-label: {chosen_aria_label})")
            
            # Ưu tiên dùng Selenium click chuẩn để React nhận sự kiện, nếu lỗi mới ép JS
            try:
                fresh_day[0].click()
            except:
                driver.execute_script("arguments[0].click();", fresh_day[0])
            
            # BẮT BUỘC CHỜ 1 giây để React xử lý sự kiện và điền ngày vào ô input
            time.sleep(1)

            # Đợi nút View Quote xuất hiện (nếu có) rồi click
            try:
                view_quote_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH,
                        "//button[contains(text(),'View Quote') or contains(text(),'ViewQuote')]"
                    ))
                )
                driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", view_quote_btn)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", view_quote_btn)
                print(f"  ✅ Đã click View Quote cho ngày {chosen_date_str}!")
            except TimeoutException:
                print(f"  ℹ️ Không thấy nút View Quote → tiếp tục bình thường")            

            time.sleep(SPEED)
            calendar_opened = True
            break  # Thành công thì thoát vòng lặp ngay

        except TimeoutException:
            print(f"  ⚠️ Web lag, chưa mở được lịch (thử lại lần {attempt + 1}), chờ 1s...")
            time.sleep(1)

    # Nếu sau 3 lần ráng sức mà vẫn không có -> Trả về No Schedule
    if not calendar_opened:
        print("  ⚠️ Không có lịch tàu khả dụng (không có ngày màu tím)!")
        return {
            "POL": row_data[2], "POD": row_data[3], "Status": "No Schedule",
            "20 DRY": "-", "40 DRY": "-", "40 HC": "-",
            "ETD": "-", "Transit Time": "-", "Valid": "-",
            "Remark": "-", "Free Time": "-",
            "Vessel": "-", "Transshipment": "-"
        }

    # Get Quote
    get_quote_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'GetQuote') or contains(text(), 'Get Quote')]")))
    driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", get_quote_btn)
    time.sleep(SPEED); js_click(get_quote_btn)
    print("  🎉 Đã bấm Get Quote. Đang đợi card giá...")
    time.sleep(3)

    # Parse cards - CHỈ LƯU DATA THUẦN, KHÔNG LƯU ELEMENT
    cards = driver.find_elements(By.XPATH, "//div[contains(@class, 'NewQuoteSummary_summary-card__')]")
    parsed_cards = []
    for idx, card in enumerate(cards):
        try:
            accept_btn = card.find_element(By.XPATH, ".//button[contains(@class, 'NewQuoteSummary_accept-btn__')]")
            if "Accept" not in accept_btn.text:
                continue
            price_text  = card.find_element(By.XPATH, ".//p[contains(@class, 'FormatMoneyV2_total-price__')]").text
            price_clean = re.sub(r'[^\d.]', '', price_text.replace('\n', '').strip())
            if '.' in price_clean:
                price_clean = price_clean.split('.')[0]
            clean_price = float(price_clean) if price_clean else 0.0

            etd_text     = card.find_element(By.XPATH, ".//div[contains(@class, 'NewQuoteSummary_route-info__')]/div[1]/p[2]").text
            etd_date     = datetime.strptime(etd_text, "%Y-%m-%d").date()
            transit_text = card.find_element(By.XPATH, ".//div[contains(@class, 'NewQuoteSummary_estimate-item__')]/p[1]").text
            transit_time = int(re.search(r'\d+', transit_text).group())

            # Chỉ lưu INDEX của card, không lưu element
            parsed_cards.append({
                "card_index": idx,
                "price": clean_price,
                "etd": etd_date,
                "transit": transit_time
            })
        except:
            continue

    if not parsed_cards:
        print("  ⚠️ Không có thẻ giá Accept nào!")
        return {"POL": row_data[2], "POD": row_data[3], "Status": "No Accept Card"}

    # Lọc ETD
    min_price  = min(c["price"] for c in parsed_cards)
    price_filt = [c for c in parsed_cards if c["price"] <= min_price + 40]
    unique_etd = {}
    for c in price_filt:
        d = c["etd"]
        if d not in unique_etd or c["transit"] < unique_etd[d]["transit"]:
            unique_etd[d] = c
    sorted_cands = sorted(unique_etd.values(), key=lambda x: x["etd"])
    min_etd      = datetime.today().date() + timedelta(days=6)
    final_sels   = []
    for c in sorted_cands:
        if c["etd"] < min_etd:
            continue
        if not final_sels:
            final_sels.append(c)
        else:
            if (c["etd"] - final_sels[-1]["etd"]).days >= 2 and \
               (c["etd"] - final_sels[0]["etd"]).days <= 9:
                final_sels.append(c)
        if len(final_sels) == 3:
            break

    if not final_sels:
        print("  ⚠️ Không có lịch tàu nào thoả mãn ETD!")
        return {"POL": row_data[2], "POD": row_data[3], "Status": "No Valid ETD"}

    # Format Excel
    etd_strs = [c["etd"].strftime("%d-%b").lstrip("0") for c in final_sels]
    if len(etd_strs) == 1:
        etd_excel = etd_strs[0]
    elif len(etd_strs) == 2:
        etd_excel = f"{etd_strs[0]} & {etd_strs[1]}"
    else:
        month = final_sels[0]["etd"].strftime("%b")
        days  = [str(c["etd"].day) for c in final_sels]
        etd_excel = f"{', '.join(days[:-1])}, {days[-1]}-{month}"

    transits      = [c["transit"] for c in final_sels]
    transit_excel = str(transits[0]) if len(set(transits)) == 1 else f"{transits[0]}-{transits[-1]}"
    valid_excel   = get_valid_date([c["etd"] for c in final_sels])

    # ── TÌM LẠI ELEMENT CỦA target_card BẰNG INDEX (tránh stale) ──
    target_data = final_sels[0]
    target_idx  = target_data["card_index"]

    def get_fresh_cards():
        """Lấy lại toàn bộ card element tươi từ DOM"""
        return driver.find_elements(
            By.XPATH, "//div[contains(@class, 'NewQuoteSummary_summary-card__')]"
        )

    def get_details_btn(card_index):
        """Lấy lại nút Details của card theo index"""
        fresh = get_fresh_cards()
        if card_index < len(fresh):
            return fresh[card_index].find_element(
                By.XPATH, ".//button[contains(@class, 'NewQuoteSummary_breakdown-button__')]"
            )
        return None
    
    # Format Excel
    etd_strs = [c["etd"].strftime("%d-%b").lstrip("0") for c in final_sels]
    if len(etd_strs) == 1:
        etd_excel = etd_strs[0]
    elif len(etd_strs) == 2:
        etd_excel = f"{etd_strs[0]} & {etd_strs[1]}"
    else:
        month = final_sels[0]["etd"].strftime("%b")
        days  = [str(c["etd"].day) for c in final_sels]
        etd_excel = f"{', '.join(days[:-1])}, {days[-1]}-{month}"

    transits      = [c["transit"] for c in final_sels]
    transit_excel = str(transits[0]) if len(set(transits)) == 1 else f"{transits[0]}-{transits[-1]}"
    valid_excel   = get_valid_date([c["etd"] for c in final_sels])

    # Bóc phí Details
    # Tìm lại details_btn tươi theo index (tránh stale)
    details_btn = get_details_btn(target_idx)
    if not details_btn:
        print("  ⚠️ Không tìm lại được nút Details!")
        return {"POL": row_data[2], "POD": row_data[3], "Status": "Details btn not found"}

    driver.execute_script(
        "arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", details_btn
    )
    time.sleep(SPEED)
    js_click(details_btn)
    time.sleep(1.5)

    EXCLUDED_FEES = [
        "doc fee (origin)", "entry summary declaration surcharge", "seal fee",
        "terminal handling charge (l)", "vietnam tax (general)", "doc fee (dest)",
        "container management fee for discharge", "terminal handling charge (d)",
        "terminal security charge (d)", "heavy surcharge"
    ]
    final_prices = {"DRY 20": 0.0, "DRY 40": 0.0, "DRY 40H": 0.0}
    has_thc, has_ens_ams, has_ows = False, False, False

    charge_items = driver.find_elements(By.XPATH, "//ul[contains(@class, 'ChargeBreakdownItem_p-sub-detail__') or contains(@class, 'ChargeBreakdownItem_p-sub-detail__LYkLW')]")
    for ul in charge_items:
        try:
            fee_name = ul.find_element(By.XPATH, "./preceding-sibling::div[1]//span[contains(@class, 'ChargeBreakdownItem_p-sub-title__')]").text.strip().lower()
        except: continue
        if "terminal handling charge (l)" in fee_name: has_thc = True
        if "entry summary declaration surcharge" in fee_name: has_ens_ams = True
        if "heavy surcharge" in fee_name or "heavy lift" in fee_name: has_ows = True
        if fee_name in EXCLUDED_FEES: continue
        is_discount = "special promotion service" in fee_name
        for line in ul.find_elements(By.XPATH, "./li"):
            text_line = line.text.strip()
            if " x 1" not in text_line: continue
            m = re.search(r'(DRY 40H|DRY 40|DRY 20).*?\(USD\s*([0-9,.]+)\)', text_line)
            if m:
                final_prices[m.group(1)] += (-float(m.group(2).replace(',','')) if is_discount else float(m.group(2).replace(',','')))

    # Xác định remark phụ dựa theo POD + country
    pod_upper     = str(row_data[3]).strip().upper()
    country_upper = str(row_data[0]).strip().upper()
    check_str     = pod_upper + " " + country_upper  # gộp để tìm kiếm

    def is_region(check, keyword_list):
        # Dùng \b để đảm bảo chỉ khớp nguyên 1 từ độc lập, không khớp chuỗi con
        return any(re.search(r'\b' + re.escape(kw) + r'\b', check) for kw in keyword_list)

    remark_str = "SUBJECT TO THC, BILL, SEAL" if has_thc else "INCL THC, SUBJECT TO BILL, SEAL"

    if is_region(check_str, CHINA_PORTS):
        remark_str += ", AMS"
    elif is_region(check_str, JAPAN_PORTS):
        remark_str += ", AFS"
    elif is_region(check_str, EUROPE_PORTS):
        remark_str += ", ENS"
    elif has_ens_ams:
        # Trường hợp ngoài các nước trên nhưng web có fee ENS/AMS → thêm AMS
        remark_str += ", AMS"

    if has_ows:
        remark_str += ", OWS"

    # ── FREE TIME: tìm lại element tươi theo card_index ──
    freetime_result = "N/A"
    try:
        # Lấy lại card tươi theo index
        fresh_cards = get_fresh_cards()
        if target_idx < len(fresh_cards):
            fresh_card_el = fresh_cards[target_idx]
        else:
            fresh_card_el = None
            print("  ⚠️ Không tìm lại được card cho free time")

        if fresh_card_el:
            # Cuộn đến card
            driver.execute_script(
                "arguments[0].scrollIntoView({behavior:'smooth', block:'center'});",
                fresh_card_el
            )
            time.sleep(0.2)

            # Tìm nút free time BÊN TRONG card tươi
            free_time_btn = driver.execute_script("""
                var card = arguments[0];
                return card.querySelector('[data-aoq-v2-free-time-tag]')
                    || card.querySelector('[class*="free-time"]')
                    || card.querySelector('[class*="FreeTime"]')
                    || card.querySelector('[class*="ChipsPopover"]');
            """, fresh_card_el)

            if free_time_btn:
                time.sleep(0.2)
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", free_time_btn
                )
                time.sleep(0.2)
                # Click MỘT LẦN duy nhất bằng JS
                driver.execute_script("arguments[0].click();", free_time_btn)

                # Đợi panel FreeTimeInfor xuất hiện
                try:
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH,
                            "//div[contains(@class,'FreeTimeInfor_body')]"
                        ))
                    )
                    time.sleep(0.4)

                    # Đọc đúng section DESTINATION
                    col_elements = driver.find_elements(By.XPATH,
                        "//div[contains(@class,'FreeTimeInfor_col-element')]"
                    )
                    dest_section = None
                    for col in col_elements:
                        try:
                            title = col.find_element(By.XPATH,
                                ".//p[contains(@class,'FreeTimeInfor_title')]"
                            ).text.strip().upper()
                            if title == "DESTINATION":
                                dest_section = col
                                break
                        except:
                            continue

                    if dest_section:
                        dest_text = dest_section.text.upper()
                        print(f"  🔍 Destination free time text: {dest_text}")

                        com_match = re.search(r'COMBINED[^\d]*(\d+)', dest_text)
                        dem_match = re.search(r'DEMURRAGE[^\d]*(\d+)', dest_text)
                        det_match = re.search(r'DETENTION[^\d]*(\d+)', dest_text)

                        if com_match:
                            freetime_result = f"{com_match.group(1)} COMBINED"
                        elif dem_match and det_match:
                            freetime_result = f"{dem_match.group(1)} DEM + {det_match.group(1)} DET"
                        elif dem_match:
                            freetime_result = f"{dem_match.group(1)} DEM"
                        elif det_match:
                            freetime_result = f"{det_match.group(1)} DET"
                        else:
                            freetime_result = "Xem thủ công"
                    else:
                        print("  ⚠️ Không tìm thấy section Destination")

                    print(f"  ⏱️ Free time (Destination): {freetime_result}")

                except TimeoutException:
                    print("  ⚠️ Panel free time không xuất hiện sau khi click")

                # Đóng panel
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.4)

            else:
                print("  ⚠️ Không tìm thấy nút free time trong card")

    except Exception as e:
        print(f"  ⚠️ Lỗi khi lấy free time: {e}")
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except:
            pass
        time.sleep(0.3)

        # Tìm nút free time nằm trong CÙNG card với target_card["btn"]
        # Đi lên ancestor summary-card rồi tìm xuống
        free_time_btn = driver.execute_script("""
            var card = arguments[0];
            return card.querySelector('[data-aoq-v2-free-time-tag]')
                || card.querySelector('[class*="free-time"]')
                || card.querySelector('[class*="FreeTime"]')
                || card.querySelector('[class*="ChipsPopover"]');
        """, fresh_card_el)

        if free_time_btn:
            time.sleep(0.2)
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", free_time_btn
            )
            time.sleep(0.2)
            # Click MỘT LẦN duy nhất
            driver.execute_script("arguments[0].click();", free_time_btn)

            # Đợi panel FreeTimeInfor xuất hiện
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH,
                        "//div[contains(@class,'FreeTimeInfor_body')]"
                    ))
                )
                time.sleep(0.4)

                # ── ĐỌC ĐÚNG SECTION DESTINATION ──
                # Dựa theo HTML: FreeTimeInfor_col-element chứa title "Destination"
                dest_section = None
                col_elements = driver.find_elements(By.XPATH,
                    "//div[contains(@class,'FreeTimeInfor_col-element')]"
                )
                for col in col_elements:
                    try:
                        title = col.find_element(By.XPATH,
                            ".//p[contains(@class,'FreeTimeInfor_title')]"
                        ).text.strip().upper()
                        if title == "DESTINATION":
                            dest_section = col
                            break
                    except:
                        continue

                if dest_section:
                    dest_text = dest_section.text.upper()
                    print(f"  🔍 Destination free time text: {dest_text}")

                    # Tìm Combined
                    com_match = re.search(r'COMBINED[^\d]*(\d+)', dest_text)
                    # Tìm Demurrage + Detention riêng
                    dem_match = re.search(r'DEMURRAGE[^\d]*(\d+)', dest_text)
                    det_match = re.search(r'DETENTION[^\d]*(\d+)', dest_text)

                    if com_match:
                        freetime_result = f"{com_match.group(1)} COMBINED"
                    elif dem_match and det_match:
                        freetime_result = f"{dem_match.group(1)} DEM + {det_match.group(1)} DET"
                    elif dem_match:
                        freetime_result = f"{dem_match.group(1)} DEM"
                    elif det_match:
                        freetime_result = f"{det_match.group(1)} DET"
                    else:
                        freetime_result = "Xem thủ công"
                else:
                    print("  ⚠️ Không tìm thấy section Destination trong free time panel")

                print(f"  ⏱️ Free time (Destination): {freetime_result}")

            except TimeoutException:
                print("  ⚠️ Panel free time không xuất hiện sau khi click")

            # Đóng panel bằng ESC
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.4)

        else:
            print("  ⚠️ Không tìm thấy nút free time trong card mục tiêu")

    except Exception as e:
        print(f"  ⚠️ Lỗi khi lấy free time: {e}")
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except:
            pass
        time.sleep(0.3)        

    # Tàu & Transshipment
    schedule_items = driver.find_elements(By.XPATH, "//li[contains(@class, 'ScheduleDepartureItem_li-boat__') or contains(@class, 'ScheduleArrival_li-location__')]")
    ports_in_route, vessels_dict = [], {}
    for item in schedule_items:
        try:
            pname = item.find_element(By.XPATH, ".//div[contains(@class, 'ScheduleItemDetails_title__')]").text.split('(')[0].strip()
            ports_in_route.append(pname)
            try: vessels_dict[pname] = item.find_element(By.XPATH, ".//span[contains(@class, 'ScheduleDepartureItem_transport-name-text__')]").text.strip()
            except: pass
        except: continue

    vessel_final, ts_ports = "", []
    if "HO CHI MINH" in pol_name or "CAI MEP" in pol_name:
        vessel_final = vessels_dict.get("CAI MEP", vessels_dict.get("HO CHI MINH", "TBA"))
        ts_ports = [p for p in ports_in_route[1:-1] if p != "CAI MEP"]
    else:
        vessel_final = vessels_dict.get(pol_name, "TBA")
        ts_ports = list(ports_in_route[1:-1])

    vessel_excel = f"{vessel_final} / ETD: {target_data['etd'].strftime('%d-%b').lstrip('0')} / Transit time: {target_data['transit']} Days"
    ts_excel     = " + ".join(ts_ports) if ts_ports else "DIRECT"

    print(f"  ✅ Xong: {row_data[2]} → {row_data[3]}")
    print(f"  💰 DRY20:{final_prices['DRY 20']:,.0f} | DRY40:{final_prices['DRY 40']:,.0f} | DRY40H:{final_prices['DRY 40H']:,.0f} USD")
    print(f"  📝 {remark_str} | Free: {freetime_result}")
    print(f"  🚢 {vessel_excel} | ⚓ {ts_excel}")

    return {
        "POL": row_data[2], "POD": row_data[3],
        "20 DRY": final_prices["DRY 20"], "40 DRY": final_prices["DRY 40"], "40 HC": final_prices["DRY 40H"],
        "ETD": etd_excel, "Transit Time": transit_excel, "Valid": valid_excel,
        "Remark": remark_str, "Free Time": freetime_result,
        "Vessel": vessel_excel, "Transshipment": ts_excel
    }


# ==========================================
# HÀM GHI KẾT QUẢ VÀO EXCEL (dùng để gọi sau MỖI batch → không mất dữ liệu)
# ==========================================
def save_results_to_excel(results, file_name=None):
    if file_name is None:
        file_name = os.environ.get("EXCEL_PATH", "input_gia.xlsx")
    """
    Ghi toàn bộ results vào file Excel. Gọi sau mỗi batch để bảo vệ
    dữ liệu khỏi mất nếu Edge/Selenium crash.
    """
    if not results:
        print("  ℹ️ Chưa có kết quả nào để ghi.")
        return
    try:
        wb = openpyxl.load_workbook(file_name)
        ws = wb.worksheets[0]

        for res in results:
            idx = res.get("orig_index")
            if idx is None:
                continue
            excel_row = idx + 2

            status   = res.get("Status", "")
            c_20     = res.get("20 DRY", "-")
            c_40     = res.get("40 DRY", "-")
            c_40hc   = res.get("40 HC", "-")
            etd      = res.get("ETD", "-")
            tt       = res.get("Transit Time", "-")
            valid    = res.get("Valid", "-")
            remark   = res.get("Remark", "-")
            free_tm  = res.get("Free Time", "-")
            vessel   = res.get("Vessel", "-")
            trans    = res.get("Transshipment", "-")

            if status and status not in ["OK", "-", ""]:
                remark = status if remark == "-" else f"{status} | {remark}"

            ws.cell(row=excel_row, column=6).value  = c_20
            ws.cell(row=excel_row, column=7).value  = c_40
            ws.cell(row=excel_row, column=8).value  = c_40hc
            ws.cell(row=excel_row, column=9).value  = etd
            ws.cell(row=excel_row, column=10).value = tt
            ws.cell(row=excel_row, column=11).value = valid
            ws.cell(row=excel_row, column=13).value = remark
            ws.cell(row=excel_row, column=14).value = free_tm
            ws.cell(row=excel_row, column=15).value = vessel
            ws.cell(row=excel_row, column=16).value = trans

        try:
            wb.save(file_name)
            print(f"  💾 Đã lưu {len(results)} dòng vào {file_name}")
        except PermissionError:
            backup_name = "input_gia_KETQUA.xlsx"
            wb.save(backup_name)
            print(f"  ⚠️ File '{file_name}' đang mở trong Excel → lưu vào '{backup_name}'")
    except Exception as e:
        print(f"  ❌ Lỗi ghi Excel: {e}")
        # Dự phòng cuối cùng: cố gắng dump CSV
        try:
            import csv
            csv_name = "input_gia_BACKUP.csv"
            with open(csv_name, "w", newline="", encoding="utf-8-sig") as f:
                if results:
                    w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
                    w.writeheader()
                    w.writerows(results)
            print(f"  💾 Backup CSV: {csv_name}")
        except Exception as e2:
            print(f"  ❌ Lỗi dự phòng CSV: {e2}")


# ==========================================
# MAIN: ĐỌC EXCEL → CHẠY TRICK 10 TAB + RELOAD ĐỒNG LOẠT
# ==========================================
results_list    = []
original_window = driver.current_window_handle

print("📂 Đang đọc file Excel...")
_ONE_EXCEL = os.environ.get("EXCEL_PATH", "input_gia.xlsx")
FILTER_POL = os.environ.get("FILTER_POL", "").strip().upper()
FILTER_POD = os.environ.get("FILTER_POD", "").strip().upper()
df = pd.read_excel(_ONE_EXCEL)

# 🟢 Tạo bản sao và lưu lại index gốc để lát map vào đúng dòng trong Excel
df_temp = df.copy()
df_temp['orig_index'] = df_temp.index

# Lọc chỉ lấy dòng có cột E (index 4) chứa chữ "ONE"
df_filtered = df_temp[
    df_temp.iloc[:, 4].astype(str).str.upper().str.contains("ONE", na=False)
]
# Lọc theo FILTER_POL/FILTER_POD nếu có (mode_route)
if FILTER_POL:
    df_filtered = df_filtered[df_filtered.iloc[:, 2].astype(str).str.upper().str.strip() == FILTER_POL]
if FILTER_POD:
    df_filtered = df_filtered[df_filtered.iloc[:, 3].astype(str).str.upper().str.strip() == FILTER_POD]
raw_data = df_filtered.dropna(
    subset=[df_filtered.columns[2], df_filtered.columns[3]]
).values.tolist()

print(f"📋 Sau khi lọc hãng ONE: {len(raw_data)} tuyến (bỏ qua các hãng khác)")
BATCH_SIZE    = 28 # Số tab tối đa mỗi batch (có thể điều chỉnh tuỳ theo hiệu năng máy và web)
total_batches = math.ceil(len(raw_data) / BATCH_SIZE)
print(f"Tổng cộng {len(raw_data)} tuyến → {total_batches} lượt (tối đa 10 tab/lượt)")

tabs = []  # Giữ xuyên suốt tất cả batch, chỉ đóng sau batch cuối

for batch_idx in range(total_batches):
    start_idx  = batch_idx * BATCH_SIZE
    end_idx    = start_idx + BATCH_SIZE
    batch_data = raw_data[start_idx:end_idx]
    is_last    = (batch_idx == total_batches - 1)

    print(f"\n{'='*54}")
    print(f"🚀 LƯỢT CHẠY {batch_idx+1}/{total_batches}  ({len(batch_data)} tuyến)")
    print(f"{'='*54}")

    # -------------------------------------------------------
    # BƯỚC 1: KIỂM TRA TAB HIỆN CÓ, MATCHING ROUTE VÀ NHẬP POL/POD
    # -------------------------------------------------------
    if batch_idx == 0:
        print("📂 Batch đầu: Đọc các tab hiện có để matching tuyến...")
        existing_handles = driver.window_handles
        
        # Bước 1.1: Đọc POL/POD đang có sẵn trên từng tab
        tab_routes = {} # Lưu {handle: {"POL": pol_name, "POD": pod_name}}
        print(f"  🔍 Đang quét {len(existing_handles)} tab đang mở...")
        
        for handle in existing_handles:
            driver.switch_to.window(handle)
            # Chỉ xử lý các tab đang ở trang ONE
            if ONE_URL in driver.current_url:
                try:
                    # Lấy text trong ô POL
                    pol_val = driver.find_element(By.XPATH, "(//input[@placeholder='Please search location'])[1]").get_attribute("value") or ""
                    # Lấy text trong ô POD
                    pod_val = driver.find_element(By.XPATH, "(//input[@placeholder='Please search location'])[2]").get_attribute("value") or ""
                    
                    # Cắt chuỗi lấy phần tên cảng trước dấu phẩy hoặc khoảng trắng (nếu có)
                    pol_clean = pol_val.split(',')[0].strip().upper() if pol_val else ""
                    pod_clean = pod_val.split(',')[0].strip().upper() if pod_val else ""
                    
                    if pol_clean and pod_clean:
                        tab_routes[handle] = {"POL": pol_clean, "POD": pod_clean}
                        print(f"    - Tab {handle[:8]}... có sẵn: {pol_clean} → {pod_clean}")
                except Exception:
                    pass # Tab này có thể chưa load xong form, bỏ qua

        # Bước 1.2: Gán tab cho các tuyến trong Excel
        unused_handles = list(existing_handles) # Các tab chưa được dùng
        for row in batch_data:
            target_pol = str(row[2]).strip().upper()
            target_pod = str(row[3]).strip().upper()
            target_country = str(row[0]).strip()
            
            matched_handle = None
            
            # Tìm xem có tab nào khớp POL và POD không
            for handle, route in tab_routes.items():
                if handle in unused_handles and target_pol in route["POL"] and target_pod in route["POD"]:
                    matched_handle = handle
                    break
            
            if matched_handle:
                print(f"\n✅ Đã match tab cho: {target_pol} → {target_pod}")
                driver.switch_to.window(matched_handle)
                tabs.append(matched_handle)
                unused_handles.remove(matched_handle)
            else:
                # Nếu không match được, lấy một tab trống/thừa hoặc mở tab mới để nhập
                print(f"\n⚠️ Chưa có tab cho: {target_pol} → {target_pod}. Chuẩn bị nhập mới...")
                if unused_handles:
                    handle = unused_handles.pop(0)
                    driver.switch_to.window(handle)
                    if ONE_URL not in driver.current_url:
                        driver.get(ONE_URL)
                        time.sleep(2)
                    tabs.append(handle)
                else:
                    handle = open_tab_and_ensure_ready()
                    tabs.append(handle)
                
                # Gọi hàm select_port cho tab này
                # (đã có try/except bên trong select_port, nhưng ta bọc thêm để tránh văng lỗi cả hệ thống)
                try:
                    select_port(1, target_pol, target_country)
                    select_port(2, target_pod, target_country)
                except Exception as e:
                    print(f"  ❌ Lỗi khi nhập cảng {target_pol} → {target_pod}: {e}")

    # -------------------------------------------------------
    # BƯỚC 2: SWITCH TAB LIÊN TỤC → ĐỢI HẾT POPUP LOADING
    # -------------------------------------------------------
    switch_tab_trick_until_clear(tabs[:len(batch_data)], rounds=5, pause=0.2)

    # -------------------------------------------------------
    # BƯỚC 3: TỪNG TAB → NHẬP CONT + GET QUOTE + BÓC GIÁ (VỚI HÀNG ĐỢI RETRY)
    # -------------------------------------------------------
    print("\n▶️  Bắt đầu nhập cont và bóc giá từng tab...")
    
    pending_tabs = list(range(len(batch_data))) # Hàng đợi các tab cần xử lý
    retries_count = {i: 0 for i in pending_tabs}
    MAX_RETRIES = 3 # Số lần cho phép quay lại tối đa nếu web lag

    while pending_tabs:
        i = pending_tabs.pop(0) # Lấy tab đầu tiên trong hàng đợi ra làm
        driver.switch_to.window(tabs[i])
        row_data = batch_data[i]
        orig_index = row_data[-1] 
        
        print(f"\n[Tab {i+1}/{len(batch_data)}] {row_data[2]} → {row_data[3]} (Lần thử: {retries_count[i] + 1})")
        try:
            result = scrape_tab(row_data)
            result["orig_index"] = orig_index 
            results_list.append(result)
        except Exception as e:
            error_msg = str(e)
            if "WEB_LAG_RETRY" in error_msg:
                if retries_count[i] < MAX_RETRIES:
                    print(f"  ⏳ Phát hiện lag! Bỏ qua tab {i+1} lúc này, đẩy xuống cuối hàng đợi để quay lại sau.")
                    retries_count[i] += 1
                    pending_tabs.append(i) # Đẩy tab này xuống cuối hàng đợi
                else:
                    print(f"  ❌ Đã quay lại {MAX_RETRIES} lần nhưng vẫn lag. Bỏ qua hoàn toàn tab {i+1}.")
                    results_list.append({
                        "POL": row_data[2], "POD": row_data[3], 
                        "orig_index": orig_index, "Status": "Lỗi: Lịch/Commodity load quá chậm"
                    })
            else:
                import traceback
                print(f"  ❌ Lỗi cứng tab {i+1}: {e}")
                print(traceback.format_exc())
                results_list.append({
                    "POL": row_data[2], "POD": row_data[3], 
                    "orig_index": orig_index, "Status": f"Lỗi: {error_msg}"
                })

    # -------------------------------------------------------
    # BƯỚC 3.5: GHI KẾT QUẢ VÀO EXCEL NGAY SAU MỖI BATCH (FIX: tránh mất dữ liệu)
    # -------------------------------------------------------
    print(f"\n💾 Lưu Excel tạm sau batch {batch_idx+1}/{total_batches}...")
    save_results_to_excel(results_list)

    # -------------------------------------------------------
    # BƯỚC 4: SAU KHI BÓC XONG BATCH
    # -------------------------------------------------------
    if not is_last:
        next_batch = raw_data[end_idx: end_idx + BATCH_SIZE]

        reload_all_tabs_simultaneously(tabs[:len(batch_data)])

        print(f"\n📝 Nhập liệu cho batch {batch_idx+2} ({len(next_batch)} tuyến)...")
        for i, row in enumerate(next_batch):
            driver.switch_to.window(tabs[i])
            print(f"  Tab {i+1}: {str(row[2]).strip()} → {str(row[3]).strip()}")
            try:
                fill_tab_ports(row)
            except Exception as e:
                print(f"  ⚠️ Tab {i+1} nhập lỗi: {e}")

        print("\n🔄 Switch tab đợi hết popup loading cho batch tiếp...")
        switch_tab_trick_until_clear(tabs[:len(next_batch)], rounds=5, pause=0.2)
        print(f"✅ Batch {batch_idx+2} sẵn sàng!")

    else:
        print("\n🧹 Batch cuối xong. Đóng tất cả tabs...")
        for tab in tabs:
            try:
                driver.switch_to.window(tab)
                driver.close()
            except:
                pass
        tabs = []
        time.sleep(1)

        try:
            remaining = driver.window_handles
            if remaining:
                driver.switch_to.window(remaining[0])
                print(f"  ✅ Đã switch sang tab còn lại: {remaining[0]}")
            else:
                print("  ℹ️ Không còn tab nào, trình duyệt đã đóng hết.")
        except Exception as e:
            print(f"  ⚠️ Không thể switch tab: {e}")


# ==========================================
# XUẤT EXCEL CUỐI CÙNG (lưu tổng hợp sau khi xong tất cả batch)
# ==========================================
print(f"\n🎉 ĐÃ CHẠY XONG! Tổng {len(results_list)} kết quả. Ghi file Excel lần cuối...")
save_results_to_excel(results_list)

# Lệnh này giúp cửa sổ console không bị tự tắt ngay lập tức
if not os.environ.get("EXCEL_PATH"):
    input("\nBấm nút Enter để thoát chương trình...")
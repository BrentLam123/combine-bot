import requests
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import calendar
import openpyxl
import os
import time
import traceback
import random  
import subprocess
import socket
import sys
import io



# ===================================================================================
# --- 1. SETUP HỆ THỐNG & API & BÙA TÀNG HÌNH ---
# ===================================================================================
current_folder = os.getcwd()
driver_path = os.path.join(current_folder, "msedgedriver.exe")

import subprocess
import socket

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

# Chỉ gọi lệnh mở Edge nếu Port 9523 chưa có ai xài
if not is_port_in_use(9523):
    print("[HỆ THỐNG] Edge COSCO chưa mở. Đang tự động khởi động...")
    try:
        subprocess.Popen([
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            "--remote-debugging-port=9523",
            r"--user-data-dir=C:\edge_cosco",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-backgrounding-occluded-windows",
            "--start-maximized"
        ])
        time.sleep(3)
    except: pass
else:
    print("[HỆ THỐNG] Edge COSCO đã mở sẵn. Bỏ qua lệnh khởi động trình duyệt.")

edge_options = Options()
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9523")
service = Service(executable_path=driver_path)
driver = webdriver.Edge(service=service, options=edge_options)

# Bùa tàng hình trị WAF
stealth_script = """
Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
window.navigator.chrome = { runtime: {} };
Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
"""
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": stealth_script
})

def login_cosco(driver):
    import time
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    print("\n--- BẮT ĐẦU ĐĂNG NHẬP COSCO ---")
    
    # TẬN DỤNG CÁC TAB CÓ SẴN (DO EDGE TỰ MỞ) THAY VÌ ĐÓNG CHÚNG
    handles = driver.window_handles
    elines_tab = None
    synconhub_tab = None

    # Quét xem tab nào đang chứa web nào
    for h in handles:
        driver.switch_to.window(h)
        url = (driver.current_url or "").lower()
        if "elines" in url: 
            elines_tab = h
        elif "synconhub" in url: 
            synconhub_tab = h

    # Lỡ mạng lag, URL chưa kịp hiện ra thì tự động lấy 2 tab đầu tiên chia cho 2 bên
    unassigned = [h for h in handles if h not in [elines_tab, synconhub_tab]]
    
    if not elines_tab:
        if unassigned: elines_tab = unassigned.pop(0)
        else:
            driver.switch_to.new_window('tab')
            elines_tab = driver.current_window_handle
            
    if not synconhub_tab:
        if unassigned: synconhub_tab = unassigned.pop(0)
        else:
            driver.switch_to.new_window('tab')
            synconhub_tab = driver.current_window_handle


    # ==========================================
    # PHASE 1: ELINES
    # ==========================================
    print("1. Đang xử lý Elines...")
    driver.switch_to.window(elines_tab)
    if "elines.coscoshipping.com/ebusiness" not in (driver.current_url or ""):
        driver.get("https://elines.coscoshipping.com/ebusiness/")
    time.sleep(3)

    # Bỏ qua Cookie pop-up (nếu có)
    try:
        cookie_btn = driver.find_element(By.XPATH, "/html/body/div[7]/div[2]/div/div/div[2]/div[2]/div[2]/button[1]")
        cookie_btn.click()
        time.sleep(2)
    except:
        pass

    # Logic đăng nhập Elines
    try:
        login_btn = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[1]/div/div[1]/div/div[1]/div/div/div/div[3]/div/div/div/div/div[1]/a")
        login_btn.click()
        time.sleep(2)
        
        email_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[9]/div[2]/div/div/div[2]/div/div/form/div[1]/div/div/div/div[1]/input")))
        email_input.send_keys("celine@pio-logistics.vn")
        
        driver.find_element(By.XPATH, "/html/body/div[9]/div[2]/div/div/div[2]/div/div/form/div[2]/div/button").click()
        time.sleep(4)
        
        submit_login = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div/div[2]/div[1]/div/form/div[4]/input[4]")))
        submit_login.click()
        
        # Dừng luồng chờ người dùng giải Captcha
        print("\n[!] CẦN GIẢI CAPTCHA ELINES! Đang chờ tự động 60s...")
        for _wait in range(60):
            time.sleep(1)
            try:
                # Nếu captcha đã được giải (trang chuyển sang trang chính), thoát vòng lặp
                if "aczoneSpotBooking" in driver.current_url or "ebusiness" in driver.current_url:
                    print("   ✅ Captcha đã được giải!")
                    break
            except:
                pass
        else:
            print("   ⚠️ Hết thời gian chờ captcha, tiếp tục...")
    except:
        pass # Đã có session sẵn

    # Trỏ thẳng tới trang check giá của Elines
    driver.get("https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")
    time.sleep(3)

    # ==========================================
    # PHASE 2: SYNCONHUB
    # ==========================================
    print("2. Đang xử lý Synconhub...")
    driver.switch_to.window(synconhub_tab)
    if "synconhub.coscoshipping.com" not in (driver.current_url or ""):
        driver.get("https://synconhub.coscoshipping.com/")
    time.sleep(3)

    # Bỏ qua Cookie pop-up (nếu có)
    try:
        cookie_btn_syn = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[3]/div[2]/div/div/div[3]/div/div/div/div/div[2]/button[1]")
        cookie_btn_syn.click()
        time.sleep(2)
    except:
        pass

    # Kiểm tra đã login chưa (có menu user / trang spot không)
    already_logged_in = False
    try:
        # Nếu trang spot đã load hoặc có avatar user → đã login
        if "synconhub.coscoshipping.com/spot" in driver.current_url:
            already_logged_in = True
        else:
            # Tìm element báo hiệu đã login (avatar / user menu)
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    ".user-info, .avatar, [class*='user-avatar'], [class*='userInfo'], .el-dropdown"))
            )
            already_logged_in = True
    except:
        pass

    if already_logged_in:
        print("   ✅ Synconhub: session còn hiệu lực, bỏ qua login.")
    else:
        # Logic đăng nhập Synconhub — dùng CSS selector thay XPath cứng
        try:
            # Tìm nút Sign In (nhiều cách)
            sign_in_syn = None
            for sel in [
                "button.sign-in-btn", "[class*='sign-in']", "[class*='login']",
                "//div[contains(text(),'Sign In') or contains(text(),'Log In') or contains(text(),'Login')]",
            ]:
                try:
                    if sel.startswith("//"):
                        sign_in_syn = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, sel))
                        )
                    else:
                        sign_in_syn = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                        )
                    break
                except:
                    continue

            if not sign_in_syn:
                # Fallback: XPath gốc
                sign_in_syn = driver.find_element(By.XPATH, "/html/body/div/div/section/div[1]/div/div[2]/div[2]")

            sign_in_syn.click()
            time.sleep(2)

            # Nhập email
            email_input_syn = None
            for sel in ["input[type='email']", "input[placeholder*='email' i]", "input[placeholder*='Email']",
                        "input[name='email']", "input[name='username']"]:
                try:
                    email_input_syn = WebDriverWait(driver, 4).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                    )
                    break
                except:
                    continue
            if not email_input_syn:
                email_input_syn = WebDriverWait(driver, 4).until(
                    EC.presence_of_element_located((By.XPATH,
                        "/html/body/div[1]/div/section/div[2]/div/div/div[2]/form/div/div/div/input"))
                )

            email_input_syn.clear()
            email_input_syn.send_keys("celine@pio-logistics.vn")
            time.sleep(0.5)

            # Bấm Next / Continue
            next_btn = None
            for sel in ["button[type='submit']", "button.next-btn", "[class*='next']", "[class*='continue']",
                        "/html/body/div[1]/div/section/div[2]/div/div/div[2]/div/button"]:
                try:
                    if sel.startswith("/"):
                        next_btn = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, sel))
                        )
                    else:
                        next_btn = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                        )
                    break
                except:
                    continue
            if next_btn:
                next_btn.click()
                time.sleep(4)
                print("   ⏳ Đang đợi Synconhub xử lý login (tối đa 30s)...")
                # Đợi redirect sang trang chính
                try:
                    WebDriverWait(driver, 30).until(
                        lambda d: "synconhub.coscoshipping.com" in d.current_url
                        and "/login" not in d.current_url
                        and "/signin" not in d.current_url
                    )
                    print("   ✅ Synconhub login thành công!")
                except:
                    print("   ⚠️ Chưa xác nhận login xong — tiếp tục...")
            else:
                print("   ⚠️ Không tìm được nút Next/Submit trên Synconhub")
        except Exception as _syn_err:
            print(f"   ⚠️ Synconhub login lỗi: {_syn_err}. Thử navigate thẳng tới spot...")

    # Trỏ thẳng tới trang check giá của Synconhub
    driver.get("https://synconhub.coscoshipping.com/spot")
    time.sleep(3)
    print("--- HOÀN TẤT ĐĂNG NHẬP COSCO ---\n")

EXCHANGE_RATE_CACHE = {} # Biến nhớ tỷ giá toàn cục
def get_live_exchange_rate(base_currency, target_currency="USD"):
    base = base_currency.upper()
    if base == target_currency.upper(): return 1.0
    
    # Rút từ Cache (Tốc độ 0.00001s)
    if base in EXCHANGE_RATE_CACHE: return EXCHANGE_RATE_CACHE[base]
    
    try:
        res = requests.get(f"https://api.frankfurter.app/latest?from={base}&to=USD", timeout=2)
        rate = float(res.json()['rates']['USD'])
        EXCHANGE_RATE_CACHE[base] = rate
        return rate
    except:
        fallbacks = {"EUR": 1.16, "AUD": 0.65, "CNY": 0.14, "VND": 0.00004, "THB": 0.028}
        EXCHANGE_RATE_CACHE[base] = fallbacks.get(base, 1.0)
        return EXCHANGE_RATE_CACHE[base]
    
# ===================================================================================
# TẢI TRƯỚC TỶ GIÁ VÀO RAM NGAY KHI KHỞI ĐỘNG
# ===================================================================================
print("[HỆ THỐNG] Đang nạp trước tỷ giá ngoại tệ vào bộ nhớ đệm...")
get_live_exchange_rate("EUR", "USD")
get_live_exchange_rate("AUD", "USD")
get_live_exchange_rate("VND", "USD")
print("[HỆ THỐNG] Nạp tỷ giá hoàn tất!")

# ===================================================================================
# DANH SÁCH CHẶN ĐỨNG (THC, SEAL, DOC VÀ CÁC PHÍ THEO BILL NHƯ AMS, ENS...)
# ===================================================================================
BLOCKLIST_CHARGES = [
    'THC', 'TERMINAL', 'THD', 'DOC', 'SLF', 'SEAL', 'BILL', 'PBF', 'TLX', 'TELEX', 
    'AMS', 'AFS', 'AFR', 'ENS', 'ISPS', 'ISP', 'PSF', 'PSU', 'DCI', 'CLE', 'EMP'
]
# ===================================================================================
# Quy tắc phí Elines (FIX):
# - Local charge đầu LOADING → luôn bỏ (dù Prepaid hay Prepaid/Collect)
# - Payment term = COLLECT thuần → bỏ (local destination)
# - OWS/Overweight → bỏ, flag has_ows
# - Tất cả phí per-container còn lại → CỘNG VÀO giá
# ===================================================================================
ELINES_LOADING_LOCAL = [
    'THC', 'TERMINAL HANDLING',   # THC đầu load
    'SLF', 'SEAL',                # Seal Fee
    'DOC', 'BILL', 'DOCUMENTATION',  # B/L, doc fee
]
ELINES_OVERWEIGHT = ['OWS', 'OVERWEIGHT', 'HCS', 'HES', 'HEAVY']

def focus_tab_by_url(domain_keyword, fallback_url):
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if domain_keyword in (driver.current_url or ""):
            return True
    print(f"   ⚠️ Tab {domain_keyword} bị tắt, đang gọi hồn lại...")
    driver.execute_script(f"window.open('{fallback_url}', '_blank');")
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)
    return True

# ===================================================================================
# ── Tự động thêm Timestamp (Thời gian thực) vào lệnh Print ──
# ===================================================================================
import sys
import io

# Xử lý lỗi Unicode trên Windows Terminal
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Lưu lại hàm print gốc
_orig_print = print

# Định nghĩa hàm print mới "độ" thêm thời gian
def print(*args, **kwargs):
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3] # Định dạng: HH:MM:SS.mili
    _orig_print(f"[{ts}]", *args, **kwargs)

# ===================================================================================
# --- 2. LOGIC LỌC 9 QUY TẮC ETD ---
# ===================================================================================
def calculate_validity(last_etd):
    import calendar
    year, month = last_etd.year, last_etd.month
    last_day = calendar.monthrange(year, month)[1]
    
    milestones = [7, 14, 21, last_day]
    for m in milestones:
        if m >= last_etd.day:
            dt = datetime(year, month, m)
            return f"{dt.day}-{dt.strftime('%b')}"
    return f"{last_day}-{last_etd.strftime('%b')}"

def apply_9_golden_rules(danh_sach_chuyen):
    danh_sach_chuyen.sort(key=lambda x: (x["etd_dt"], x["tt_days"]))
    list_loc_trung = []
    seen_dates = set()
    for c in danh_sach_chuyen:
        if c["etd_dt"] not in seen_dates:
            list_loc_trung.append(c)
            seen_dates.add(c["etd_dt"])

    etd_dat_chuan = []
    if list_loc_trung:
        ngan_nhat_global = min(c["tt_days"] for c in list_loc_trung)
        first_date = list_loc_trung[0]["etd_dt"]
        for c in list_loc_trung:
            if len(etd_dat_chuan) >= 3: break 
            if len(etd_dat_chuan) > 0 and (c["etd_dt"] - etd_dat_chuan[-1]["etd_dt"]).days < 2: continue
            if (c["etd_dt"] - first_date).days <= 9 and (c["tt_days"] <= ngan_nhat_global + 10):
                etd_dat_chuan.append(c)

    format_str = ""
    num = len(etd_dat_chuan)
    
    if num == 1:
        dt0 = etd_dat_chuan[0]["etd_dt"]
        format_str = f"{dt0.day}-{dt0.strftime('%b')}"
    elif num == 2:
        dt0 = etd_dat_chuan[0]["etd_dt"]
        dt1 = etd_dat_chuan[1]["etd_dt"]
        format_str = f"{dt0.day}-{dt0.strftime('%b')} & {dt1.day}-{dt1.strftime('%b')}"
    elif num >= 3:
        dt0 = etd_dat_chuan[0]["etd_dt"]
        dt1 = etd_dat_chuan[1]["etd_dt"]
        dt2 = etd_dat_chuan[2]["etd_dt"]
        format_str = f"{dt0.day}, {dt1.day}, {dt2.day}-{dt2.strftime('%b')}"

    all_tt = [c["tt_days"] for c in etd_dat_chuan]
    tt_min, tt_max = min(all_tt), max(all_tt)
    str_tt = f"{tt_min}" if tt_min == tt_max else f"{tt_min}-{tt_max}"
    
    return etd_dat_chuan, format_str, str_tt

# ===================================================================================
# --- 3. BỘ NÃO NHẬP CẢNG (BẢN TỐI ƯU: KHÔNG XÓA MÙ QUÁNG - ZERO DELAY) ---
# ===================================================================================
def select_port_smart(xpath, port_name, country, sys_name, is_elines=False):
    print(f"      + [{sys_name}] Nạp cảng: {port_name}, {country}")
    
    clean_port_name = port_name.split(',')[0].strip().upper()
    country_upper = country.upper()

    for attempt in range(3): 
        try:
            # 1. Tìm ô nhập liệu (Dùng WebDriverWait để không bị rớt do load chậm)
            inp = WebDriverWait(driver, 5).until(lambda d: next((e for e in d.find_elements(By.XPATH, xpath) if e.is_displayed()), None))
            if not inp: raise Exception("Không tìm thấy ô!")
            
            # 2. Đọc giá trị hiện tại
            curr_val = driver.execute_script("return arguments[0].value;", inp).strip().upper()
            
            # KIỂM TRA CHUẨN: 
            if country_upper in curr_val and (curr_val.startswith(f"{clean_port_name},") or curr_val.startswith(f"{clean_port_name} ,") or curr_val == clean_port_name):
                print(f"        -> Đã có sẵn CHUẨN: {curr_val} -> ĐI TIẾP!")
                return

            # 3. NẠP CHỮ (SMART FILL): 
            # Nếu ô đang trắng hoặc chứa chữ tào lao -> Mới xóa gõ lại.
            # Nếu đã có sẵn "ALGECIRAS" -> Giữ nguyên, chỉ kích hoạt sự kiện.
            if clean_port_name not in curr_val:
                if sys_name == "Synconhub": driver.execute_script("arguments[0].removeAttribute('readonly');", inp)
                driver.execute_script("""
                    let input = arguments[0];
                    let val = arguments[1];
                    let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    nativeInputValueSetter.call(input, val);
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                """, inp, clean_port_name)
            
            # Kích hoạt sự kiện để Dropdown hiện ra (Không xóa chữ cũ)
            driver.execute_script("arguments[0].dispatchEvent(new Event('focus', { bubbles: true }));", inp)
            driver.execute_script("arguments[0].dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: ' ' }));", inp)
                
            # 4. ĐỢI DROPDOWN VÀ CLICK
            css_sel = "div.el-autocomplete-suggestion:not([style*='display: none']) li" if is_elines else "div.el-select-dropdown:not([style*='display: none']) li"
            
            target_strict = f"{clean_port_name}," 
            target_semi = f"{clean_port_name} ,"
            
            timeout = time.time() + 5
            match = None
            while time.time() < timeout:
                opts = driver.find_elements(By.CSS_SELECTOR, css_sel)
                for o in opts:
                    txt = o.text.strip().upper()
                    if country_upper not in txt: 
                        continue
                    
                    # LUẬT LỌC KHẮT KHE: Bắt buộc phần đầu phải khớp hoàn toàn tới dấu phẩy
                    if txt.startswith(target_strict) or txt.startswith(target_semi) or txt == clean_port_name:
                        match = o
                        break
                if match: break
                time.sleep(0.05)
            
            if match:
                final_text = match.text.strip()
                driver.execute_script("arguments[0].click();", match)
                print(f"        -> Đã chốt từ list: {final_text}")
                return 
            
            raise Exception("Không tìm thấy option phù hợp chuẩn xác trong list")
            
        except Exception as e:
            time.sleep(0.5)
            
    raise Exception(f"Thất bại nạp cảng {port_name} sau 3 lần thử.")


# ===================================================================================
# --- 4. SYNCONHUB ---
# ===================================================================================
def run_synconhub(pol, pod, pod_country):
    step_tracker = "Khởi động hàm Synconhub"
    try:
        driver.switch_to.default_content()
        
        step_tracker = "Nhập Origin"
        select_port_smart("//div[contains(@class, 'ect-label') and text()='Origin']/following-sibling::div//input", pol, "VIETNAM", "Synconhub")
        
        step_tracker = "Nhập Destination"
        select_port_smart("//div[contains(@class, 'ect-label') and text()='Destination']/following-sibling::div//input", pod, pod_country, "Synconhub")

        step_tracker = "Nhập ETD + 10 Ngày"
        print("      -> [Synconhub] Đang nạp ngày ETD + 10...")
        try:
            etd_inp = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@name='sailing_product_date_picker_start']")))
            dt_str = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", etd_inp)
            driver.execute_script("arguments[0].removeAttribute('readonly');", etd_inp)
            driver.execute_script("arguments[0].value = '';", etd_inp)
            driver.execute_script(f"arguments[0].value = '{dt_str}';", etd_inp)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", etd_inp)
            time.sleep(0.5)
            
            try: etd_inp.send_keys(Keys.ESCAPE)
            except: pass
            driver.execute_script("document.body.click();")
            print(f"      -> [Synconhub] Đã chốt ngày ETD: {dt_str}")
        except Exception as e:
            print(f"      ⚠️ [Synconhub] Kẹt lúc nạp ETD: {e}")

        step_tracker = "Tắt Cont Lạnh (RF/NOR)"
        for lbl in ["RF", "NOR"]:
            try:
                elem = driver.find_element(By.XPATH, f"//label[contains(@class, 'el-checkbox') and .//span[contains(text(), '{lbl}')]]")
                if "is-checked" in elem.get_attribute("class"): driver.execute_script("arguments[0].click();", elem)
            except: pass

        step_tracker = "Bấm nút Search"
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(@class, 'ect-search-btn')]"))
        print("      -> [Synconhub] Đã bấm Search...")
        
        step_tracker = "Chờ Loading Mask"
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".el-loading-mask")))
            WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".el-loading-mask")))
        except: pass

        step_tracker = "Kiểm tra kết quả / No Result"
        time.sleep(1) 
        try:
            WebDriverWait(driver, 15).until(
                lambda d: any(e.is_displayed() for e in d.find_elements(By.CSS_SELECTOR, ".ect-search-result-body")) or 
                          any(e.is_displayed() for e in d.find_elements(By.CSS_SELECTOR, ".ect-search-no-result"))
            )
        except: 
            print("      ❌ [Synconhub] 15s trôi qua không thấy thẻ giá!")
            return None
            
        no_res_elements = driver.find_elements(By.CSS_SELECTOR, ".ect-search-no-result")
        if any(e.is_displayed() for e in no_res_elements):
            print("      -> [Synconhub] NO SERVICE / SOLD OUT!")
            return None
            
        step_tracker = "Đọc thẻ giá"
        print("      -> [Synconhub] Thẻ giá hiện, đang đọc...")
        time.sleep(1)
        cards = driver.find_elements(By.CSS_SELECTOR, ".ect-search-result-body")
        list_chuyen = []
        for card in cards:
            if not card.is_displayed(): continue
            try:
                etd_dt = datetime.strptime(card.find_element(By.XPATH, ".//div[text()='ETD']/following-sibling::div").text.strip(), "%Y-%m-%d")
                tt_text = card.find_element(By.CSS_SELECTOR, ".transit-time span").text.replace("days", "").strip()
                prices = [float(i.find_element(By.CSS_SELECTOR, ".ect-price").text.replace('$', '').replace(',', '').strip()) for i in card.find_elements(By.CSS_SELECTOR, ".container-type-item") if i.find_element(By.CSS_SELECTOR, ".ect-price").text.strip()]
                if prices: list_chuyen.append({"element": card, "price": min(prices), "etd_dt": etd_dt, "tt_days": int(tt_text) if tt_text.isdigit() else 999})
            except: continue

        if not list_chuyen: 
            print("      ❌ [Synconhub] Không đọc được thẻ giá!")
            return None

        step_tracker = "Lọc 9 quy tắc vàng"
        gia_re_nhat = min(c["price"] for c in list_chuyen)
        list_chuyen = [c for c in list_chuyen if c["price"] == gia_re_nhat]
        etd_chuan, str_etd, str_tt = apply_9_golden_rules(list_chuyen)
        dai_dien = etd_chuan[0]["element"]
        valid_date_str = calculate_validity(etd_chuan[-1]["etd_dt"]) # <-- THÊM DÒNG NÀY
        print(f"      -> [Synconhub] Chọn ETD {str_etd}, TT {str_tt} days")

# --- BẮT ĐẦU THÊM: RÚT THÔNG TIN TÀU VÀ CHUYỂN TẢI (SYNCONHUB) ---
        step_tracker = "Lấy Thông Tin Tàu (Synconhub)"
        print("      -> [Synconhub] Đang soi tên tàu và cảng chuyển tải...")
        vessel_texts = []
        ts_combos = []
        
        for c in etd_chuan:
            card = c["element"]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
            time.sleep(0.4)
            
            # 1. Lấy Tên Tàu (Cắt trước dấu | và dấu /) BẰNG JAVASCRIPT
            vessel_name = "TBA"
            try:
                # Dùng JS leo ra ngoài thẻ wrapper rồi mới móc xuống footer
                v_raw = driver.execute_script("""
                    let wrapper = arguments[0].closest('.ect-search-result-wrapper');
                    if (wrapper) {
                        let footer = wrapper.querySelector('.ect-search-result-footer');
                        if (footer) {
                            let span = footer.querySelector('span');
                            if (span) return span.textContent.trim();
                        }
                    }
                    return '';
                """, card)
                
                if v_raw:
                    # Tách phần trước dấu | (bỏ phần tàu phụ)
                    v_part = v_raw.split('|')[0].strip()
                    # Tách phần trước dấu / (bỏ mã service)
                    vessel_name = v_part.split('/')[0].strip()
            except Exception as e:
                print(f"      [Debug] Lỗi rút tên tàu: {e}")
                
            # 2. Lấy Cảng Chuyển Tải
            ts_ports = []
            try:
                # Kiểm tra xem có icon T/S không
                ts_el = card.find_elements(By.XPATH, ".//div[contains(@class, 'schedule-tip')]//div[contains(text(), 'T/S')]")
                if ts_el:
                    # Bơm JS kích hoạt hover chuột để ép popup bung ra
                    driver.execute_script("""
                        arguments[0].scrollIntoView({block: 'center', inline: 'center'});
                        arguments[0].dispatchEvent(new MouseEvent('mouseenter', {bubbles: true, cancelable: true, view: window}));
                    """, ts_el[0])
                    time.sleep(1.5) # Đợi popup bung
                    
                    # Rút data từ popup
                    st_names = driver.execute_script("""
                        let pops = document.querySelectorAll('.el-tooltip__popper[x-placement]');
                        for (let i = pops.length - 1; i >= 0; i--) {
                            let p = pops[i];
                            if (window.getComputedStyle(p).display === 'none') continue;
                            let points = p.querySelectorAll('.point');
                            if (points.length > 0) {
                                let names = [];
                                points.forEach(pt => {
                                    let divs = pt.querySelectorAll('div');
                                    // Cột div số 3 (index 2) chứa tên tiếng Anh
                                    if (divs.length >= 3) {
                                        let txt = divs[2].textContent.trim();
                                        if(txt) names.push(txt.toUpperCase());
                                    }
                                });
                                return names;
                            }
                        }
                        return [];
                    """)
                    
                    if st_names and len(st_names) >= 2:
                        origin = st_names[0]
                        dest = st_names[-1]
                        for p in st_names:
                            # Lọc bỏ cảng xếp, cảng dỡ và các trạm trung chuyển bị lặp tên
                            if p != origin and p != dest and p not in ts_ports:
                                ts_ports.append(p)
                                
                    # Đóng popup bằng cách click ra ngoài
                    driver.execute_script("document.body.click();")
                    driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('mouseleave', {bubbles: true, cancelable: true, view: window}));", ts_el[0])
                    time.sleep(0.3)
            except Exception as e:
                print(f"      [Debug] Lỗi rút T/S Synconhub: {e}")
                
            # Tổng hợp chuỗi
            str_ts = " + ".join(ts_ports) if ts_ports else "DIRECT"
            str_etd_fmt = f"{c['etd_dt'].day}-{c['etd_dt'].strftime('%b')}"
            str_tt_fmt = str(c["tt_days"])
            
            info_str = f"{vessel_name} / ETD: {str_etd_fmt} / Transit time: {str_tt_fmt} Days / Transshipment Port: {str_ts}"
            vessel_texts.append(info_str)
            ts_combos.append(str_ts)
            
        unique_ts = []
        for t in ts_combos:
            if t not in unique_ts: unique_ts.append(t)
            
        final_ts_str = " or \n".join(unique_ts)
        final_vessel_str = "\n".join(vessel_texts)
        # --- KẾT THÚC THÊM ---

        step_tracker = "Đọc Base Rate"
        rates = {"20GP": None, "40GP": None, "40HQ": None}
        for item in dai_dien.find_elements(By.CSS_SELECTOR, ".container-type-item"):
            ctype = item.find_element(By.CSS_SELECTOR, ".cntr-title").text.strip()
            if ctype in rates:
                try: rates[ctype] = float(item.find_element(By.CSS_SELECTOR, ".ect-price").text.replace('$', '').replace(',', '').strip())
                except: pass
        print(f"      -> [Synconhub] Base rates: {rates}")

        step_tracker = "Mở dấu chấm hỏi Surcharge"
        has_ows = False
        thc_inc = False
        manifest_fee_found = False
        try:
            print("      -> [Synconhub] Mở bảng Surcharge...")
            s_btn = dai_dien.find_element(By.CSS_SELECTOR, ".extra-charge-btn")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", s_btn)
            time.sleep(0.5); driver.execute_script("arguments[0].click();", s_btn)
            
            table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'el-popover') and not(contains(@style, 'display: none'))]//table[contains(@class, 'el-table__body')]")))
            time.sleep(1)
            
            step_tracker = "Xử lý bảng Surcharge Synconhub"
            curr_cat = ""
            for row in table.find_elements(By.TAG_NAME, "tr"):
                tds = row.find_elements(By.TAG_NAME, "td")
                if len(tds) < 6: continue
                if len(tds) == 7: curr_cat = tds[0].text.strip().upper()
                if curr_cat == "DESTINATION": continue
                
                c_name = tds[-6].text.strip().upper()
                u_text = tds[-5].text.strip().upper()
                if "B/L" in u_text or "BL" in u_text: continue
                
                if any(x in c_name for x in ["ENS", "AMS", "AFS", "ADVANCED MANIFEST", "ENTRY SUMMARY"]):
                    manifest_fee_found = True
                    print(f"        [FLAG] Đã tóm được {c_name[:25]} -> Ghi chú vào Remark!")

                if any(x in c_name for x in ELINES_OVERWEIGHT):
                    has_ows = True; continue
                if any("INCLUDED" in tds[-3+i].text.strip().upper() for i in range(3)) and ("THC" in c_name or "TERMINAL" in c_name):
                    thc_inc = True; continue
                    
                if any(b in c_name for b in BLOCKLIST_CHARGES): 
                    continue
                
                for idx, ctype in enumerate(["20GP", "40GP", "40HQ"]):
                    val = tds[-3+idx].text.strip().replace(',', '')
                    if "INCLUDED" in val.upper() or not val: continue
                    parts = val.split()
                    if len(parts) == 2:
                        amt = float(parts[1])
                        # Lấy tỷ giá trực tiếp tại đây, nhờ có Cache nên tốc độ là 0.0001s
                        if parts[0].upper() == 'EUR': 
                            amt *= get_live_exchange_rate("EUR", "USD")
                        
                        if rates[ctype] is not None:
                            rates[ctype] += amt
                            print(f"        [+] {amt:.2f} USD ({c_name[:15]}) → {ctype}")
                            
            driver.execute_script("document.body.click();")
            print("      -> [Synconhub] Xong bảng Surcharge!")
        except Exception as e: 
            print(f"      ⚠️ [Synconhub] Lỗi Surcharge: {e}")

        if thc_inc:
            print("      -> [Synconhub] THC đã included hoặc không xuất hiện -> ghi INCLUDED O.THC vào remark.")

        print(f"      -> [Synconhub] XONG! Cước: {rates}")
        # THÊM BIẾN VALID VÀO LÚC RETURN
        return {
            "rates": rates, "etd": str_etd, "tt": str_tt, "ows": has_ows, 
            "thc_inc": thc_inc, "valid": valid_date_str, "manifest_fee": manifest_fee_found,
            "surcharge_error": False, "vessel_info": final_vessel_str, "transshipment": final_ts_str
        }
    
    except Exception as e:
        print("\n" + "!"*60)
        print(f"      🚨 SYNCONHUB CRASH TẠI: [{step_tracker}]")
        print(f"      👉 {e}")
        print(traceback.format_exc())
        print("!"*60 + "\n")
        return None

# ===================================================================================
# --- 5a. HÀM PHỤ: NHẬP CÂN NHANH ELINES (CÓ RADAR BÁO CÁO) ---
# ===================================================================================
def _elines_fill_weight_fast(driver):
    deadline = time.time() + 20
    injected = False
    
    # Biến để theo dõi thời gian in log
    start_time = time.time()
    last_print_time = time.time()
    
    while time.time() < deadline:
        all_inputs = driver.find_elements(By.CSS_SELECTOR,
            "input[type='number']:not([readonly]):not([disabled])")
        visible = [i for i in all_inputs if i.is_displayed()]
        
        if visible:
            # FIX: Chờ 0.5s để UI web bung nốt các ô nhập cân của 20GP và 40GP
            time.sleep(0.5)
            
            # Lấy lại danh sách các ô lần nữa sau khi đã bung đủ
            all_inputs = driver.find_elements(By.CSS_SELECTOR,
                "input[type='number']:not([readonly]):not([disabled])")
            visible = [i for i in all_inputs if i.is_displayed()]
            
            for inp in visible:
                try:
                    driver.execute_script(
                        "arguments[0].value='22222';"
                        "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                        inp
                    )
                except:
                    pass
            for inp in driver.find_elements(By.CSS_SELECTOR, ".el-input-number input"):
                try:
                    if inp.is_displayed():
                        driver.execute_script(
                            "arguments[0].value='1';"
                            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                            inp
                        )
                except:
                    pass
            injected = True
            print(f"      -> [Elines] ✅ Web đã vẽ xong form! Đã nhập 15000kg siêu tốc cho {len(visible)} ô!")
            time.sleep(0.5) # Nghỉ thêm nửa giây cho web nhận đủ số trước khi bấm Calculate
            break
            
        # --- RADAR BÁO CÁO MỖI 1 GIÂY ---
        current_time = time.time()
        if current_time - last_print_time >= 1.0:
            elapsed = int(current_time - start_time)
            print(f"        ⏳ Vẫn đang chờ mạng load form Detail... ({elapsed}s)")
            last_print_time = current_time
            
        time.sleep(0.1)
        
    if not injected:
        print("      ⚠️ [Elines] Chờ 20s mà web COSCO rớt mạng không load nổi form!")


# ===================================================================================
# --- 5b. HÀM PHỤ: ĐỌC GIÁ TỪ UNIT PRICE CELL (FIX) ---
# Dùng find_elements thay vì .//span[1] và .//span[2] để tránh lỗi im lặng
# ===================================================================================
def _parse_price_cell(price_td, rate_eur, rate_aud=None, rate_vnd=None):
    """
    Đọc (currency, amount_usd) từ Unit Price cell của Elines.
    Cell HTML: <div><span class='text-xs'>USD</span><span class='text-sm'>94</span></div>
    Trả về float (USD) hoặc None nếu lỗi.
    """
    try:
        spans = price_td.find_elements(By.TAG_NAME, "span")
        # Lọc span có text thực (bỏ span rỗng)
        valid = [s for s in spans if s.text.strip()]
        if len(valid) < 2:
            return None
        c_str = valid[0].text.strip().upper()
        amt_raw = valid[1].text.strip().replace(',', '')
        if not amt_raw:
            return None
        amt = float(amt_raw)
        if c_str == 'EUR':
            amt *= rate_eur
        elif c_str == 'AUD':
            amt *= (rate_aud or get_live_exchange_rate("AUD", "USD"))
        elif c_str == 'VND':
            amt *= (rate_vnd or get_live_exchange_rate("VND", "USD"))
        # USD: giữ nguyên
        return amt
    except Exception as e:
        return None


# ===================================================================================
# --- 5c. ELINES (ĐÃ FIX SURCHARGE LOGIC) ---
# ===================================================================================
def run_elines(pol, pod, pod_country):
    step_tracker = "Khởi động"
    try:
        driver.switch_to.default_content()

    # --- BẮT ĐẦU SỬA: VÒNG LẶP ANTI WEB TRẮNG ---
        search_success = False
        for attempt in range(2):
            if attempt == 1:
                print("      -> [Elines] 🔄 Phát hiện web trắng hoặc load lỗi! Đang ép tải lại trang (Reload)...")
                driver.refresh()
                time.sleep(5)
                driver.switch_to.default_content()

            try:
                step_tracker = "Chuyển iframe"
                if attempt == 0:
                    print("      -> [Elines] BƯỚC 1: Chui vào iframe...")
                else:
                    print("      -> [Elines] BƯỚC 1: Chui vào iframe (Thử lại lần 2)...")
                    
                try:
                    iframe = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "aczoneIframe")))
                    driver.switch_to.frame(iframe)
                except: pass

                step_tracker = "Nhập Origin"
                select_port_smart("//input[@placeholder='Please input Origin City']", pol, "VIETNAM", "Elines", True)

                step_tracker = "Nhập Destination"
                select_port_smart("//input[@placeholder='Please input Destination City']", pod, pod_country, "Elines", True)

                step_tracker = "Nhập ETD"
                try:
                    date_inp = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'el-date-editor')]//input")))
                    dt_str = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
                    driver.execute_script("arguments[0].value='';", date_inp)
                    driver.execute_script(f"arguments[0].value='{dt_str}';", date_inp)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", date_inp)            
                except: pass

                step_tracker = "Tick Container"
                for c in ["20GP", "40GP", "40HQ"]:
                    try:
                        lbl = driver.find_element(By.XPATH, f"//label[contains(@class,'el-checkbox') and contains(.,'{c}')]")
                        if "is-checked" not in lbl.get_attribute("class"):
                            driver.execute_script("arguments[0].click();", lbl)
                    except: pass

                step_tracker = "Tắt Operating Reefer"
                try:
                    reefer_lbl = driver.find_element(By.XPATH, "//label[contains(@class,'el-checkbox') and contains(.,'Operating Reefer')]")
                    if "is-checked" in reefer_lbl.get_attribute("class"):
                        driver.execute_script("arguments[0].click();", reefer_lbl)
                        print("      -> [Elines] Đã TẮT tùy chọn Operating Reefer!")
                except: pass

                step_tracker = "Bấm Search Service"
                driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(.,'Search Service')]"))
                print("      -> [Elines] BƯỚC 2: Đã bấm Search...")

                time.sleep(2)
                try:
                    WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".el-loading-mask")))
                    WebDriverWait(driver, 45).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".el-loading-mask")))
                except: pass
                time.sleep(1)

                step_tracker = "Kiểm tra kết quả"
                # Rút ngắn thời gian chờ xuống 15s để nếu nó lỗi trắng thì nhanh chóng reload
                WebDriverWait(driver, 15).until(lambda d:
                    d.find_elements(By.XPATH, "//div[contains(@class,'box-border') and contains(@class,'hover:shadow-lg')]") or
                    d.find_elements(By.CSS_SELECTOR, ".el-empty") or
                    d.find_elements(By.XPATH, "//*[contains(text(),'No products matching')]")
                )
                
                search_success = True
                break # Thành công -> Thoát vòng lặp reload an toàn
                
            except Exception as e:
                if attempt == 0:
                    continue # Bị timeout do web trắng -> Vòng lặp sẽ chạy attempt 1 (reload)
                else:
                    print("      ❌ [Elines] Đã Reload nhưng web vẫn trắng/treo. Chịu thua!")
                    return None
                    
        if not search_success:
            return None
            # --- KẾT THÚC SỬA ---

        if (driver.find_elements(By.CSS_SELECTOR, ".el-empty") or
            driver.find_elements(By.XPATH,
                "//*[contains(text(),'No products matching')]")):
            print("      -> [Elines] NO SERVICE / SOLD OUT!")
            return None

        step_tracker = "Kiểm tra Flash Sale filter"
        print("      -> [Elines] BƯỚC 3: Kiểm tra Flash Sale...")

        has_any_flash_sale = bool(driver.find_elements(By.XPATH,
            "//div[contains(@class,'blue-service') or "
            "contains(@class,'blue-service-selected')]"
            "[contains(text(),'Flash Sale')]"))

        is_flash_sale_mode = False

        if has_any_flash_sale:
            print("      -> [Elines] 🔥 Phát hiện Flash Sale!")
            try:
                select_wrapper = driver.find_element(By.XPATH,
                    "//div[contains(@class,'el-select__tags')]"
                    "[.//input[contains(@class,'el-select__input')]]")
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_wrapper)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", select_wrapper)

                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "li.el-select-dropdown__item")))
                time.sleep(0.3)

                flash_option = driver.find_element(By.XPATH,
                    "//li[contains(@class,'el-select-dropdown__item')]"
                    "[.//span[text()='Flash Sale']]")
                driver.execute_script("arguments[0].click();", flash_option)
                print("      -> [Elines] Đã chọn Flash Sale!")

                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, ".el-loading-mask")))
                    WebDriverWait(driver, 15).until(EC.invisibility_of_element_located(
                        (By.CSS_SELECTOR, ".el-loading-mask")))
                except:
                    pass
                time.sleep(1)
                is_flash_sale_mode = True
            except Exception as e:
                print(f"      ⚠️ [Elines] Không filter được Flash Sale: {e}")
        else:
            print("      -> [Elines] Không có Flash Sale, dùng giá thường.")

        step_tracker = "Đọc dữ liệu cards (DEBUG MODE)"
        print("      -> [Elines] Bật chế độ DEBUG soi thẻ card...")
        
        cards = driver.find_elements(By.XPATH,
            "//div[contains(@class,'box-border') and "
            "contains(@class,'hover:shadow-lg')]")
        list_c = []

        if not cards:
            print("      ❌ [DEBUG] Màn hình hiện tại không có thẻ card nào (XPath box-border không khớp)!")
            return None

        for idx, card in enumerate(cards):
            if not card.is_displayed():
                continue
            
            # CHỤP X-QUANG: Lưu ngay cấu trúc HTML của thẻ đầu tiên ra file
            if idx == 0:
                html_content = card.get_attribute('outerHTML')
                with open("debug_card_Venezia.html", "w", encoding="utf-8") as f:
                    f.write(html_content)
                print(f"      ✅ [DEBUG] Đã lưu mã HTML của Thẻ 1 vào file 'debug_card_Venezia.html' nằm chung thư mục với bot.")

            try:
                # 1. Đọc Ngày đi (ETD)
                try:
                    etd_raw = card.find_element(By.XPATH,
                        ".//div[contains(@class,'bg-[#1890FF]')]"
                        "//p[contains(@class,'bottom-2.5')]//span[1]"
                    ).text.split('(')[0].strip()
                    etd_dt = datetime.strptime(f"2026{etd_raw}", "%Y%b%d")
                except Exception as e_etd:
                    raise Exception(f"Kẹt ở lúc đọc Ngày ETD: {e_etd}")

                # 2. Đọc Transit Time (TT)
                tt_days = 999
                try:
                    tt_el = card.find_element(By.XPATH,
                        ".//p[contains(@class,'space-x-0.5')]//span[1]")
                    tt_text = tt_el.text.strip()
                    if tt_text.isdigit():
                        tt_days = int(tt_text)
                except:
                    print(f"        [DEBUG] Thẻ {idx+1}: Không tìm thấy Transit time bằng chữ, thử trừ từ ETA...")
                    try:
                        eta_raw = card.find_element(By.XPATH,
                            ".//div[contains(@class,'bg-[#CCCCCC]')]"
                            "//p[contains(@class,'bottom-2.5')]//span[1]"
                        ).text.split('(')[0].strip()
                        eta_dt = datetime.strptime(f"2026{eta_raw}", "%Y%b%d")
                        tt_days = (eta_dt - etd_dt).days
                    except Exception as e_eta:
                        print(f"        [DEBUG] Thẻ {idx+1}: ETA cũng móm nốt! Lỗi: {e_eta}")

                # 3. Đọc Giá tiền
                price_spans = card.find_elements(By.XPATH,
                    ".//p[contains(@class,'text-[#F1A104]')]//span[last()]")
                
                if not price_spans:
                    raise Exception("Mảng price_spans bị trống (Không tìm thấy XPath chứa màu chữ vàng #F1A104)")
                    
                prices = []
                for p in price_spans:
                    txt = p.text.replace(',', '').strip()
                    try:
                        val = float(txt)
                        if val > 50:
                            prices.append(val)
                    except:
                        pass

                if not prices:
                    raise Exception(f"Có element giá nhưng không lôi được số ra. Giá trị thô đang lấy được là: {[p.text for p in price_spans]}")

                # 4. Đọc tình trạng chỗ (Space)
                try:
                    space_txt = card.find_element(By.XPATH,
                        ".//p[contains(@class,'text-warning')]//span").text.strip().upper()
                except:
                    space_txt = "TBC"

                list_c.append({
                    "element": card,
                    "price":   min(prices),
                    "etd_dt":  etd_dt,
                    "tt_days": tt_days,
                    "space":   space_txt,
                })
                print(f"        Card {idx+1}: {etd_raw} TT={tt_days}d ${min(prices)} [{space_txt}]")
                
            except Exception as e:
                # IN RA LỖI RÕ RÀNG ĐỂ BẮT BỆNH
                print(f"        ⚠️ [DEBUG] Thẻ số {idx+1} bị rớt đài do: {e}")
                continue

        if not list_c:
            print("      ❌ [Elines] Không đọc được card nào hoàn chỉnh!")
            return None

        step_tracker = "Lọc 9 quy tắc"
        
        space_priority = {
            "TIGHT": 1,
            "TBC": 2
        }
        
        # --- BẮT ĐẦU SỬA: LỌC TEU VÀ CHỌN GIÁ RẺ NHẤT CÓ ĐỦ CHỖ ---
        import re
        valid_list_re = []
        
        # Lấy danh sách các mức giá từ rẻ đến mắc (Ví dụ: $425 test trước, $525 test sau)
        cac_muc_gia = sorted(list(set(c["price"] for c in list_c)))
        
        for muc_gia in cac_muc_gia:
            # Lấy các card có mức giá này
            list_re_temp = [c for c in list_c if c["price"] == muc_gia]
            
            # Sắp xếp để test card AVAILABLE trước, card TBC sau trong cùng 1 mức giá
            list_re_temp.sort(key=lambda x: space_priority.get(x["space"].upper(), 3), reverse=True)
            
            if is_flash_sale_mode:
                print(f"      -> [Elines] Xét mức giá ${muc_gia}: Kiểm tra Remaining Stock (yêu cầu >= 5 TEU)...")
                valid_for_this_price = []
                
                for c in list_re_temp:
                    card = c["element"]
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
                    time.sleep(0.4)
                    
                    # 1. Bơm JS ép click mở panel Flash Sale (nếu có)
                    driver.execute_script("""
                        let card = arguments[0];
                        let tags = card.querySelectorAll('.blue-service, .blue-service-selected');
                        for(let t of tags) {
                            if(t.textContent.includes('Flash Sale')) {
                                t.click();
                            }
                        }
                    """, card)
                    time.sleep(1) # Nghỉ 1s chờ panel bung ra trọn vẹn
                    
                    # 2. Bơm JS đọc xuyên DOM lấy cụm text "Remaining Stock... TEU"
                    teu_text = driver.execute_script("""
                        let card = arguments[0];
                        let h4s = card.querySelectorAll("h4");
                        for (let h of h4s) {
                            if (h.textContent.includes("Remaining Stock")) {
                                return h.parentElement.textContent; 
                            }
                        }
                        return null;
                    """, card)
                    
                    if teu_text:
                        # Dùng regex trích xuất đúng con số nằm trước chữ TEU
                        match = re.search(r'(\d+)\s*TEU', teu_text, re.IGNORECASE)
                        if match:
                            teu_val = int(match.group(1))
                            if teu_val < 5:
                                print(f"        ⚠️ Bỏ qua card ETD {c['etd_dt'].strftime('%d-%b')} [{c['space']}] vì chỉ còn {teu_val} TEU (< 5)")
                                # Đóng bảng Flash Sale của Card bị loại để dọn DOM
                                driver.execute_script("""
                                    let tags = arguments[0].querySelectorAll('.blue-service, .blue-service-selected');
                                    for(let t of tags) {
                                        if(!t.textContent.includes('Flash Sale')) {
                                            t.click();
                                            break;
                                        }
                                    }
                                """, card)
                                time.sleep(0.5)
                                continue
                            else:
                                print(f"        ✅ Card ETD {c['etd_dt'].strftime('%d-%b')} [{c['space']}] còn {teu_val} TEU -> Đạt chuẩn!")
                        else:
                            print(f"        ⚠️ Có chữ Remaining Stock nhưng không trích xuất được số từ: {teu_text}")
                            continue # Lỗi format của hãng tàu -> Không an toàn -> Vứt!
                    else:
                        print(f"        ⚠️ Không tìm thấy dòng Remaining Stock của card ETD {c['etd_dt'].strftime('%d-%b')}. Sẽ bỏ qua thẻ này.")
                        continue # Kỷ luật sắt: Không đọc được cũng vứt luôn!
                    
                    # Nếu chạy đến được dòng này tức là card đã xuất sắc vượt qua test
                    valid_for_this_price.append(c)
                
                if valid_for_this_price:
                    valid_list_re = valid_for_this_price
                    break # Đã tìm thấy các card hợp lệ ở mức giá này, chốt và thoát vòng lặp giá
            else:
                # Nếu giá THƯỜNG (không check được TEU) thì vẫn phải lọc ưu tiên SPACE (Bỏ TIGHT, lấy AVAILABLE)
                max_prio_temp = max(space_priority.get(c["space"].upper(), 3) for c in list_re_temp)
                valid_list_re = [c for c in list_re_temp if space_priority.get(c["space"].upper(), 3) == max_prio_temp]
                break
                
        if not valid_list_re:
            print("      ❌ [Elines] Tất cả các card Flash Sale ở mọi mức giá đều không đủ điều kiện (Dưới 5 TEU hoặc lỗi)!")
            return None
            
        list_re = valid_list_re
        e_chuan, s_etd, s_tt = apply_9_golden_rules(list_re)
        # --- KẾT THÚC SỬA ---       
       
        valid_date_str = calculate_validity(e_chuan[-1]["etd_dt"]) # <-- THÊM DÒNG NÀY
        print(f"      -> [Elines] Chốt ETD: {s_etd} | T/T: {s_tt} days")
        
        # --- BẮT ĐẦU THÊM: LẤY THÔNG TIN TÀU & TRANSSHIPMENT PORT ---
        step_tracker = "Lấy Thông Tin Tàu"
        print("      -> [Elines] Đang soi tên tàu và cảng chuyển tải...")
        vessel_texts = []
        ts_combos = []

        # Import bộ giả lập chuột người thật
        from selenium.webdriver.common.action_chains import ActionChains

        for c in e_chuan:
            card = c["element"]
            # Cuộn trang cho thẻ Card nằm giữa màn hình để dễ rê chuột
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
            time.sleep(0.4)
            
            # 1. Lấy Tên Tàu (Cắt bỏ mã Service ở đầu)
            vessel_name = "TBA"
            try:
                v_els = card.find_elements(By.XPATH, ".//span[contains(@class, 'text-dark-black') and contains(@class, 'font-bold')]")
                for v in v_els:
                    t = v.text.strip().replace('\n', ' ')
                    if t and not t.startswith("202"):
                        # Chặt chuỗi ở khoảng trắng đầu tiên (Bỏ HPX2, VTS...)
                        parts = t.split(" ", 1)
                        vessel_name = parts[1] if len(parts) > 1 else t
                        break
            except: pass
                
           
            # 2. KÍCH HOẠT VÀ HÚT DATA POPUP HOÀN TOÀN BẰNG JAVASCRIPT (TRỊ LỖI TEXT TÀNG HÌNH)
            ts_ports = []
            try:
                days_el = card.find_element(By.XPATH, ".//*[contains(translate(text(), 'DAY', 'day'), 'day')]")
                
                # Bơm JS kích hoạt TẤT CẢ các thẻ có khả năng là nút bấm trong khu vực đó
                trigger_ok = driver.execute_script("""
                    let box = arguments[0].closest('.flex-grow') || arguments[0].parentElement.parentElement;
                    // Bắt trọn ổ: class trigger của tooltip, hoặc class cursor-pointer
                    let triggers = box.querySelectorAll('.el-tooltip__trigger, svg.cursor-pointer');
                    if (triggers.length > 0) {
                        triggers[0].scrollIntoView({block: 'center', inline: 'center'});
                        triggers.forEach(t => {
                            // Bắn liên thanh cả hover lẫn click để ép popup phải lòi ra
                            t.dispatchEvent(new MouseEvent('mouseenter', {bubbles: true, cancelable: true, view: window}));
                            t.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, view: window}));
                        });
                        return true;
                    }
                    return false;
                """, days_el)
                
                if trigger_ok:
                    time.sleep(1.5) # Đợi 1.5s cho popup vẽ xong
                    
                    # CẬP NHẬT: Dò bằng textContent và quét bao quát hơn
                    st_names = driver.execute_script("""
                        let pops = document.querySelectorAll('.el-popper');
                        let names = [];
                        
                        // Quét ngược từ dưới lên (ưu tiên lấy popup vừa mới nặn ra)
                        for (let i = pops.length - 1; i >= 0; i--) {
                            let p = pops[i];
                            
                            // Bỏ qua nếu popup bị ẩn
                            let isHidden = window.getComputedStyle(p).display === 'none' || p.getAttribute('aria-hidden') === 'true';
                            if (isHidden) continue;
                            
                            // Tìm tất cả các trạm
                            let spans = p.querySelectorAll("span.font-bold.text-black");
                            if (spans.length > 0) {
                                spans.forEach(s => {
                                    // Dùng textContent để lấy chữ thô tuyệt đối (Không lo CSS che khuất)
                                    let txt = s.textContent || "";
                                    if(txt.trim()) names.push(txt.trim().toUpperCase());
                                });
                                return names; // Hút xong của popup này là té luôn
                            }
                        }
                        return names;
                    """)
                    
                    if st_names:
                        if len(st_names) > 2:
                            ts_ports = st_names[1:-1] # Cắt bỏ Origin và Dest, lấy phần ruột chuyển tải
                    else:
                        print("      [Debug] JS đã kích hoạt nhưng lúc móc túi thì popup trống rỗng!")
                else:
                    print("      [Debug] JS không tìm thấy cái nút trigger nào!")
                        
                # Dọn dẹp: Bấm ra ngoài và bắn sự kiện leave để đóng popup cũ lại
                driver.execute_script("document.body.click();")
                driver.execute_script("""
                    let box = arguments[0].closest('.flex-grow') || arguments[0].parentElement.parentElement;
                    let triggers = box.querySelectorAll('.el-tooltip__trigger, svg.cursor-pointer');
                    triggers.forEach(t => {
                        t.dispatchEvent(new MouseEvent('mouseleave', {bubbles: true, cancelable: true, view: window}));
                    });
                """, days_el)
                time.sleep(0.3)
                
            except Exception as e:
                print(f"      [Debug] Lỗi rình popup Cảng: {e}")

            str_ts = " + ".join(ts_ports) if ts_ports else "DIRECT"
            str_etd_fmt = f"{c['etd_dt'].day}-{c['etd_dt'].strftime('%b')}"
            str_tt_fmt = str(c["tt_days"])
            
            info_str = f"{vessel_name} / ETD: {str_etd_fmt} / Transit time: {str_tt_fmt} Days / Transshipment Port: {str_ts}"
            vessel_texts.append(info_str)
            ts_combos.append(str_ts)

        unique_ts = []
        for t in ts_combos:
            if t not in unique_ts: unique_ts.append(t)
            
        final_ts_str = " or \n".join(unique_ts)
        final_vessel_str = "\n".join(vessel_texts)
        # --- KẾT THÚC THÊM ---

        target_card = e_chuan[0]["element"]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_card)
        time.sleep(0.5)

        step_tracker = "Đọc giá + Booking"
        rates = {"20GP": None, "40GP": None, "40HQ": None}

        if is_flash_sale_mode:
            try:
                flash_panel = target_card.find_element(By.ID, "spot-booking-premium-service")
                if not flash_panel.is_displayed():
                    raise Exception("panel ẩn")
            except:
                try:
                    fb = target_card.find_element(By.XPATH,
                        ".//div[contains(@class,'blue-service')][contains(text(),'Flash Sale')]")
                    driver.execute_script("arguments[0].click();", fb)
                    WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.ID, "spot-booking-premium-service")))
                    time.sleep(1)
                except Exception as e:
                    print(f"      ⚠️ Không mở Flash Sale panel: {e}")

            try:
                flash_panel = target_card.find_element(By.ID, "spot-booking-premium-service")
                cont_blocks = flash_panel.find_elements(By.XPATH,
                    ".//div[contains(@class,'flex-col') and contains(@class,'items-start')]"
                    "[.//span[contains(@class,'text-[#666666]')]]")
                for block in cont_blocks:
                    try:
                        ctype = block.find_element(By.XPATH,
                            ".//span[contains(@class,'text-[#666666]')]").text.strip()
                        if ctype not in rates: continue
                        val = block.find_element(By.XPATH,
                            ".//div[contains(@class,'text-[#F1A104]')]//span[last()]"
                        ).text.replace(',', '').strip()
                        try:
                            rates[ctype] = float(val)
                            print(f"        [🔥FLASH] {ctype} = ${rates[ctype]}")
                        except:
                            pass
                            print(f"        [🔥FLASH] {ctype} = ${rates[ctype]}")
                    except:
                        continue
            except Exception as e:
                print(f"      ⚠️ Không đọc giá Flash Sale: {e}")

            print("      -> [Elines] Bấm Booking Flash Sale...")
            b_btn = WebDriverWait(target_card, 5).until(EC.presence_of_element_located((
                By.XPATH, ".//button[starts-with(@id,'bkg2-premium-service-booking-btn-')]")))

        else:
            cont_blocks = target_card.find_elements(By.XPATH,
                ".//div[contains(@class,'flex-col') and contains(@class,'flex-1')]"
                "[.//span[contains(@class,'text-[#666666]')]]")
            for block in cont_blocks:
                try:
                    ctype = block.find_element(By.XPATH,
                        ".//span[contains(@class,'text-[#666666]')]").text.strip()
                    if ctype not in rates: continue
                    val = block.find_element(By.XPATH,
                        ".//p[contains(@class,'text-[#F1A104]')]//span[last()]"
                    ).text.replace(',', '').strip()
                    try:
                        rates[ctype] = float(val)
                        print(f"        [THƯỜNG] {ctype} = ${rates[ctype]}")
                    except:
                        pass
                        print(f"        [THƯỜNG] {ctype} = ${rates[ctype]}")
                except:
                    continue

            print("      -> [Elines] Bấm Booking thường...")
            b_btn = target_card.find_element(By.XPATH, ".//button[@id='bkg2-spot-booking-btn']")

        driver.execute_script("arguments[0].click();", b_btn)
        print("      -> [Elines] ĐÃ BẤM BOOKING! Chờ form load...")

        # ================================================================
        # BƯỚC 7: TRANG DETAIL - NHẬP CÂN
        # ================================================================
        step_tracker = "Trang Detail - Điền cân"
        driver.switch_to.default_content()

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "aczoneIframe")))
        except:
            pass

        try:
            iframe_detail = driver.find_element(By.ID, "aczoneIframe")
            driver.switch_to.frame(iframe_detail)
            print("      -> [Elines] Vào iframe Detail!")
        except:
            print("      -> [Elines] Không thấy iframe Detail, tiếp tục...")

        print("      -> [Elines] Bắt đầu dò tìm ô nhập cân (Tốc độ 0.1s)...")

        _elines_fill_weight_fast(driver)

        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'Container Info') or contains(text(),'Cargo Weight')]")))
        except:
            pass

        step_tracker = "Trang Detail - Calculate"
        try:
            calc_btn = driver.find_element(By.XPATH,
                "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'calculate')]")
            driver.execute_script("arguments[0].click();", calc_btn)
            print("      -> [Elines] Đã bấm Calculate!")
        except:
            pass

        # ================================================================
        # BƯỚC 8: HÚT SURCHARGE - LOGIC CHỜ DỮ LIỆU ỔN ĐỊNH 
        # ================================================================
        step_tracker = "Hút Surcharge"
        
        # 1. Chờ vòng xoay (max 2s)
        try:
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".el-loading-mask")))
        except: pass

        # 2. Chờ vòng xoay biến mất (max 10s)
        try:
            WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".el-loading-mask")))
        except: pass

        # 3. POLLING CHỜ ỔN ĐỊNH (Bất chấp web trả về bao nhiêu phí)
        print("      -> [Elines] Đang rình bảng Surcharge (Chờ dữ liệu ngừng nhảy dòng)...")
        timeout_content = time.time() + 15
        
        last_row_count = 0
        stable_time = 0
        
        while time.time() < timeout_content:
            status = driver.execute_script("""
                let rows = document.querySelectorAll('table.el-table__body tr.el-table__row');
                let tableText = document.querySelector('table.el-table__body') ? document.querySelector('table.el-table__body').innerText.toUpperCase() : '';
                return {
                    count: rows.length,
                    has_surcharge: tableText.includes('SURCHARGE'),
                    has_ocean: tableText.includes('OCEAN')
                };
            """)
            
            curr_count = status['count']
            
            # Phải có đủ cả Ocean và Surcharge mới bắt đầu kiểm tra
            if curr_count > 0 and status['has_ocean'] and status['has_surcharge']:
                if curr_count == last_row_count:
                    # Nếu số lượng dòng không đổi trong 1 giây (5 nhịp x 0.2s) -> Web đã bung xong data
                    stable_time += 0.2
                    if stable_time >= 1.0:
                        break
                else:
                    # Đang đếm mà web đẻ thêm dòng mới -> Reset, đếm lại 1 giây từ đầu!
                    last_row_count = curr_count
                    stable_time = 0
            else:
                stable_time = 0
                
            time.sleep(0.2)

        print(f"      -> [Elines] Bảng đã ổn định ({last_row_count} dòng)! Bắt đầu hút...")

        rates_detail = {"20GP": None, "40GP": None, "40HQ": None}
        has_ows = False
        manifest_fee_found = False
        
        js_code = """
        let data = [];
        let rows = document.querySelectorAll("table.el-table__body tr.el-table__row");
        for (let r of rows) {
            let tds = r.querySelectorAll("td");
            if (tds.length < 8) continue;
            let c_name_idx = tds.length >= 11 ? 1 : 0;
            let c_name = tds[c_name_idx].innerText.trim().toUpperCase();
            let unit = tds[tds.length - 6].innerText.trim().toUpperCase();
            let term = tds[tds.length - 5].innerText.trim().toUpperCase();
            let currency = "";
            let amount = "";
            let price_td = tds[tds.length - 3];
            let spans = price_td.querySelectorAll("span");
            let valid_spans = [];
            for (let s of spans) {
                if (s.innerText.trim() !== "") valid_spans.push(s.innerText.trim());
            }
            if (valid_spans.length >= 2) {
                currency = valid_spans[0].toUpperCase();
                amount = valid_spans[1].replace(/,/g, '');
            }
            data.push({c_name: c_name, unit: unit, term: term, currency: currency, amount: amount});
        }
        return data;
        """
        table_data = driver.execute_script(js_code)
      

        # PYTHON XỬ LÝ DỮ LIỆU TỪ JS TRÊN RAM
        for row in table_data:
            c_name = row['c_name']
            unit = row['unit']
            term = row['term']
            curr = row['currency']
            amt_raw = row['amount']

            if term == "COLLECT":
                print(f"        [-COLLECT] Bỏ qua {c_name[:25]} (Do term là COLLECT)")
                continue
            
            if any(x in c_name for x in ["ENS", "AMS", "AFS", "ADVANCED MANIFEST", "ENTRY SUMMARY"]):
                manifest_fee_found = True
                print(f"        [FLAG] Đã tóm được {c_name[:25]} -> Ghi chú vào Remark!")
            
            if any(b in c_name for b in BLOCKLIST_CHARGES):
                print(f"        [-BLOCKLIST] Bỏ qua {c_name[:25]}")
                continue
           
            if "BL" in unit or "B/L" in unit or unit not in ["20GP", "40GP", "40HQ"]:
                print(f"        [-BL] Bỏ qua {c_name[:25]} (Do tính theo B/L hoặc sai đơn vị)")
                continue

            if not amt_raw: 
                print(f"        ⚠️ [BỎ QUA] Phí {c_name[:15]} không có giá tiền (Rỗng).")
                continue
            try:
                amt = float(amt_raw)
            except:
                continue

            # ĐỔI TIỀN LAZY (Chỉ khi nào thấy EUR/AUD/VND mới móc hàm tỷ giá ra tính)
            if curr == 'EUR': 
                amt *= get_live_exchange_rate("EUR", "USD")
            elif curr == 'AUD': 
                amt *= get_live_exchange_rate("AUD", "USD")
            elif curr == 'VND': 
                amt *= get_live_exchange_rate("VND", "USD")

            if "OCEAN RATE" in c_name or "OCEAN FREIGHT" in c_name:
                rates_detail[unit] = amt
                print(f"        [Ocean] {unit} = ${amt:.2f}")
                continue

            if any(x in c_name for x in ELINES_OVERWEIGHT):
                has_ows = True

            if rates_detail[unit] is not None:
                rates_detail[unit] += amt
                print(f"        [+SURCHARGE] +${amt:.2f} ({c_name[:25]}) → {unit}")
            else:
                rates_detail[unit] = amt
                print(f"        [+SURCHARGE] +${amt:.2f} ({c_name[:25]}) → {unit}")

        # ─── Merge: ưu tiên Detail, fallback về giá card ───
        final_rates = {}
        for ct in ["20GP", "40GP", "40HQ"]:
            final_rates[ct] = (rates_detail[ct]
                               if rates_detail[ct] is not None
                               else rates.get(ct))
        # Kiểm tra xem có lấy được Surcharge nào ngoài Ocean không
        surcharge_error_flag = True
        for row in table_data:
            name = row['c_name'].upper()
            if name and not any(x in name for x in ["OCEAN RATE", "OCEAN FREIGHT"]):
                surcharge_error_flag = False
                break

        src = "🔥Flash Sale" if is_flash_sale_mode else "Thường"
        print(f"      -> [Elines] ✅ XONG [{src}]! Cước cuối: {final_rates}")
        
         # ─── Back ───
        step_tracker = "Back"
        try:
            b_back = driver.find_element(By.XPATH, "//button[contains(.,'Back')]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", b_back)
            time.sleep(0.4)
            driver.execute_script("arguments[0].click();", b_back)
            time.sleep(1)
        except:
            driver.get("https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")

        return {
            "rates": final_rates,
            "etd":   s_etd,
            "tt":    s_tt,
            "ows":   has_ows,
            "thc_inc": False,
            "valid": valid_date_str, 
            "manifest_fee": manifest_fee_found,
            "surcharge_error": surcharge_error_flag, # <-- CỜ BÁO LỖI
            "vessel_info": final_vessel_str,         # <-- TÊN TÀU
            "transshipment": final_ts_str            # <-- CẢNG CHUYỂN TẢI
        }

    except Exception as e:
        print("\n" + "!"*60)
        print(f"      🚨 ELINES CRASH: [{step_tracker}]")
        print(f"      👉 {e}")
        print(traceback.format_exc())
        print("!"*60 + "\n")
        driver.get("https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")
        return None
    

# ===================================================================================
# --- 6. HÀM TẠO REMARK CHUẨN (CÓ BÁO LỖI SURCHARGE) ---
# ===================================================================================
def generate_remark(thc_inc, ows, manifest_fee, pod, source, is_surcharge_error=False):
    lst = ["BILL", "SEAL", "TELEX"]
    if not thc_inc: lst.insert(0, "THC")
    
    if manifest_fee:
        pod_upper = pod.upper()
        if any(k in pod_upper for k in ["SHANGHAI", "NINGBO", "QINGDAO", "XINGANG", "TIANJIN", "YANTIAN", "SHEKOU", "NANSHA", "XIAMEN", "DALIAN", "CHINA"]):
            lst.append("AMS")
        elif any(k in pod_upper for k in ["TOKYO", "YOKOHAMA", "NAGOYA", "OSAKA", "KOBE", "MOJI", "HAKATA", "JAPAN"]):
            lst.append("AFS")
        else:
            lst.append("ENS")
            
    if ows: lst.append("OWS")
    rem = "SUBJECT TO " + ", ".join(lst)
    if thc_inc: rem += ", INCLUDED O.THC"
    
    if is_surcharge_error:
        rem += " (SURCHARGE ERROR)"
        
    return rem

# ===================================================================================
# --- MAIN RUN ---
# ===================================================================================
print("\n>>> BOT COSCO - FIX SURCHARGE v2 <<<")

EXCEL_FILE = os.environ.get("EXCEL_PATH", "input_gia.xlsx")
FILTER_POL = os.environ.get("FILTER_POL", "").strip().upper()
FILTER_POD = os.environ.get("FILTER_POD", "").strip().upper()
wb = openpyxl.load_workbook(EXCEL_FILE)
sheet = wb.active

# --- BẮT ĐẦU THÊM: Đọc dữ liệu từ file FREE TIME.xlsx ---
free_time_dict = {}
try:
    wb_ft = openpyxl.load_workbook("FREE TIME.xlsx", data_only=True)
    sheet_ft = wb_ft.active
    for r in range(2, sheet_ft.max_row + 1):
        key = str(sheet_ft.cell(row=r, column=1).value or "").strip().upper()
        val = str(sheet_ft.cell(row=r, column=2).value or "").strip()
        if key:
            free_time_dict[key] = val
except Exception as e:
    print(f"⚠️ Lỗi khi đọc file FREE TIME.xlsx: {e}")
# --- KẾT THÚC THÊM ---

failed_rows = []  # Mảng lưu lại các dòng bị báo lỗi NO SERVICE để chạy phiên 2

# ---------------------------------------------------------
# PHIÊN CHẠY 1: CHẠY ĐỒNG THỜI SYNCONHUB & ELINES
# ---------------------------------------------------------
print(f"\n{'='*52}")
print("▶️ BẮT ĐẦU PHIÊN CHẠY 1 (SYNCONHUB + ELINES)")
print(f"{'='*52}")

login_cosco(driver) # <--- THÊM DÒNG NÀY VÀO ĐÂY
# --- BẮT ĐẦU THÊM: LOGIC ĐỌC LỆNH TỪ TERMINAL ---
import sys
target_row = None
if len(sys.argv) > 1:
    try:
        target_row = int(sys.argv[1])
        print(f"🛠️ [CHẾ ĐỘ TEST] Bỏ qua toàn bộ, CHỈ CHẠY DUY NHẤT DÒNG {target_row}!")
    except: pass

# Nếu có nhập số dòng trên Terminal thì chỉ chạy mảng có 1 dòng đó, nếu không thì chạy từ dòng 2 đến hết
danh_sach_dong = [target_row] if target_row else range(2, sheet.max_row + 1)

for row in danh_sach_dong:
    pod_country = str(sheet.cell(row=row, column=2).value or "").strip()
    if not pod_country:
        pod_country = os.environ.get("FILTER_COUNTRY", "").strip()
    pol = str(sheet.cell(row=row, column=3).value or "").strip()
    pod = str(sheet.cell(row=row, column=4).value or "").strip()
    carrier = str(sheet.cell(row=row, column=5).value or "").strip().upper()

    # --- BẮT ĐẦU THÊM: Chuẩn hóa tên cảng đặc biệt COSCO ---
    PORT_MAPPING = {
        "TIANJIN": "XINGANG",
    }
    
    # Ép kiểu viết hoa, xóa khoảng trắng thừa và tra từ điển
    if pol:
        pol = PORT_MAPPING.get(str(pol).upper().strip(), str(pol).strip())
    if pod:
        pod = PORT_MAPPING.get(str(pod).upper().strip(), str(pod).strip())
    # --- KẾT THÚC THÊM ---

    if not pol or not pod or carrier != "COSCO":
        continue
    if FILTER_POL and str(sheet.cell(row=row, column=3).value or "").strip().upper() != FILTER_POL:
        continue
    if FILTER_POD and str(sheet.cell(row=row, column=4).value or "").strip().upper() != FILTER_POD:
        continue

    print(f"\n{'='*52}")
    print(f"🚀 SO SÁNH GIÁ (LẦN 1): {pol} -> {pod} (Row {row})")
    
    # Kiểm tra nếu POD thuộc AUSTRALIA thì bỏ qua Synconhub
    if pod_country and "AUSTRALIA" in pod_country.upper():
        print("   [1] Synconhub... ⏩ BỎ QUA (Tuyến Úc chỉ check Elines)")
        res_s = None
    else:
        print("   [1] Synconhub...")
        focus_tab_by_url("synconhub.coscoshipping.com", "https://synconhub.coscoshipping.com/spot")
        time.sleep(1)
        if "login" in (driver.current_url or "").lower() or "auth" in (driver.current_url or "").lower():
            print("   ⚠️ Phát hiện Session Synconhub hết hạn! Tiến hành đăng nhập lại...")
            login_cosco(driver)
            focus_tab_by_url("synconhub.coscoshipping.com", "https://synconhub.coscoshipping.com/spot")
        res_s = run_synconhub(pol, pod, pod_country)

    print("   [2] Elines...")
    focus_tab_by_url("elines.coscoshipping.com", "https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")
    time.sleep(1)
    if "login" in (driver.current_url or "").lower() or "auth" in (driver.current_url or "").lower():
        print("   ⚠️ Phát hiện Session Elines hết hạn! Tiến hành đăng nhập lại...")
        login_cosco(driver)
        focus_tab_by_url("elines.coscoshipping.com", "https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")
    res_e = run_elines(pol, pod, pod_country)

    chot_deal = None
    source = ""

    def get_compare_price(res):
        if res is None: return float('inf')
        p20 = res['rates']['20GP']
        p40 = res['rates']['40GP']
        if p20 is not None: return p20
        if p40 is not None: return p40
        return float('inf')

    price_s = get_compare_price(res_s)
    price_e = get_compare_price(res_e)

    if price_s == float('inf') and price_e == float('inf'):
        chot_deal = None
    elif price_s <= price_e:
        chot_deal, source = res_s, "SYNCONHUB"
    else:
        chot_deal, source = res_e, "ELINES"

    if chot_deal:
        r = chot_deal['rates']
        print(f"   🏆 CHỐT TỪ: {source}")
        v20 = f"${round(r['20GP'],2)}" if r['20GP'] is not None else "N/A"
        v40 = f"${round(r['40GP'],2)}" if r['40GP'] is not None else "N/A"
        vHQ = f"${round(r['40HQ'],2)}" if r['40HQ'] is not None else "N/A"
        print(f"      20'={v20} | 40'={v40} | 40H={vHQ}")
        
        sheet.cell(row=row, column=6).value = round(r['20GP'], 2) if r['20GP'] is not None else ""
        sheet.cell(row=row, column=7).value = round(r['40GP'], 2) if r['40GP'] is not None else ""
        sheet.cell(row=row, column=8).value = round(r['40HQ'], 2) if r['40HQ'] is not None else ""
        sheet.cell(row=row, column=9).value = chot_deal['etd']
        sheet.cell(row=row, column=10).value = chot_deal['tt']
        sheet.cell(row=row, column=11).value = chot_deal['valid'] # <-- ĐIỀN VALID CỘT K
        
        # BẮT ĐẦU THÊM: Ghi Remark có truyền thêm cờ Surcharge Error
        is_err = chot_deal.get('surcharge_error', False)
        sheet.cell(row=row, column=13).value = generate_remark(chot_deal['thc_inc'], chot_deal['ows'], chot_deal['manifest_fee'], pod, source, is_err)
        
        # Ghi Free Time cột N (14)
        ft_value = free_time_dict.get(pod_country.upper(), "") 
        sheet.cell(row=row, column=14).value = ft_value

        # Ghi Thông tin Tàu (Cột O) và Chuyển tải (Cột P)
        sheet.cell(row=row, column=15).value = chot_deal.get('vessel_info', '')
        sheet.cell(row=row, column=16).value = chot_deal.get('transshipment', '')

        # --- BẮT ĐẦU THÊM: Ghi Free Time cột N (14) tra theo QUỐC GIA ---
        ft_value = free_time_dict.get(pod_country.upper(), "") 
        print(f"      -> [Free Time] Đã tìm thấy cho quốc gia {pod_country.upper()}: '{ft_value}'")
        sheet.cell(row=row, column=14).value = ft_value
        # --- KẾT THÚC THÊM ---
    else:
        print("   ❌ Cả 2 hệ thống đều móm ở Lần 1!")
        sheet.cell(row=row, column=13).value = "NO SERVICE / SOLD OUT"
        failed_rows.append(row)  # Đưa dòng này vào danh sách đen chờ xử ở Phiên 2

    wb.save(EXCEL_FILE)

# ---------------------------------------------------------
# PHIÊN CHẠY 2: CHỈ VÉT LẠI CÁC DÒNG LỖI BẰNG ELINES
# ---------------------------------------------------------
if failed_rows:
    print(f"\n{'='*52}")
    print(f"🔄 BẮT ĐẦU PHIÊN CHẠY 2: VÉT LẠI {len(failed_rows)} TUYẾN BỊ LỖI TRÊN ELINES 🔄")
    print(f"{'='*52}")

    for row in failed_rows:
        pod_country = str(sheet.cell(row=row, column=2).value or "").strip()
        if not pod_country:
            pod_country = os.environ.get("FILTER_COUNTRY", "").strip()
        pol = str(sheet.cell(row=row, column=3).value or "").strip()
        pod = str(sheet.cell(row=row, column=4).value or "").strip()
        # --- BẮT ĐẦU THÊM: Chuẩn hóa tên cảng đặc biệt COSCO ---
        PORT_MAPPING = {
            "FOS SUR MER": "FOS",
            "GENOA": "GENOVA",
            "NAPOLI": "NAPLES",
            "COCHIN": "KOCHI"
        }
        
        # Ép kiểu viết hoa, xóa khoảng trắng thừa và tra từ điển
        if pol:
            pol = PORT_MAPPING.get(str(pol).upper().strip(), str(pol).strip())
        if pod:
            pod = PORT_MAPPING.get(str(pod).upper().strip(), str(pod).strip())
        # --- KẾT THÚC THÊM ---
        
        print(f"\n🚀 SO SÁNH GIÁ (LẦN 2): {pol} -> {pod} (Row {row})")
        
        print("   [2] Elines (Retry) - Đang reload lại trang để clear lag...")
        
        # 1. Trỏ vào đúng tab Elines (hoặc mở mới nếu lỡ tắt)
        focus_tab_by_url("elines.coscoshipping.com", "https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")
        
        # 2. Ép trình duyệt Reload thẳng lại link gốc để dọn sạch rác/state cũ
        driver.get("https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/")
        time.sleep(4)  # Chờ 4s cho iframe và các component Vue của hãng tàu load xong hoàn toàn
        
        # 3. Bắt đầu chạy lại hàm rút giá
        res_e = run_elines(pol, pod, pod_country)

        if res_e:
            r = res_e['rates']
            source = "ELINES (RETRY)"
            print(f"   🏆 PHIÊN 2 ĐÃ VỚT ĐƯỢC DEAL TỪ: {source}")
            
            v20 = f"${round(r['20GP'],2)}" if r['20GP'] is not None else "N/A"
            v40 = f"${round(r['40GP'],2)}" if r['40GP'] is not None else "N/A"
            vHQ = f"${round(r['40HQ'],2)}" if r['40HQ'] is not None else "N/A"
            print(f"      20'={v20} | 40'={v40} | 40H={vHQ}")
            
            # Ghi đè lại dữ liệu (xoá cái NO SERVICE cũ)
            sheet.cell(row=row, column=6).value = round(r['20GP'], 2) if r['20GP'] is not None else ""
            sheet.cell(row=row, column=7).value = round(r['40GP'], 2) if r['40GP'] is not None else ""
            sheet.cell(row=row, column=8).value = round(r['40HQ'], 2) if r['40HQ'] is not None else ""
            sheet.cell(row=row, column=9).value = res_e['etd']
            sheet.cell(row=row, column=10).value = res_e['tt']
            sheet.cell(row=row, column=11).value = res_e['valid'] # <-- ĐIỀN VALID CỘT K
        
            # BẮT ĐẦU THÊM: Ghi Remark có truyền thêm cờ Surcharge Error
            is_err = res_e.get('surcharge_error', False)
            sheet.cell(row=row, column=13).value = generate_remark(res_e['thc_inc'], res_e['ows'], res_e['manifest_fee'], pod, source, is_err)
            
            # Ghi Free Time cột N (14)
            ft_value = free_time_dict.get(pod_country.upper(), "") 
            sheet.cell(row=row, column=14).value = ft_value

            # Ghi Thông tin Tàu (Cột O) và Chuyển tải (Cột P)
            from openpyxl.styles import Alignment
            
            vessel_cell = sheet.cell(row=row, column=15)
            vessel_cell.value = res_e.get('vessel_info', '')
            vessel_cell.alignment = Alignment(wrap_text=True) # Kích hoạt tính năng xuống dòng trong ô Excel
            
            sheet.cell(row=row, column=16).value = res_e.get('transshipment', '')
            # --- BẮT ĐẦU THÊM: Ghi Free Time cột N (14) tra theo QUỐC GIA ---
            ft_value = free_time_dict.get(pod_country.upper(), "") 
            print(f"      -> [Free Time] Đã tìm thấy cho quốc gia {pod_country.upper()}: '{ft_value}'")
            sheet.cell(row=row, column=14).value = ft_value
            # --- KẾT THÚC THÊM ---
        else:
            print("   ❌ Phiên 2 vẫn móm, bỏ qua!")

        wb.save(EXCEL_FILE)

print("\n🎉 XONG TOÀN BỘ TIẾN TRÌNH! CHECK FILE EXCEL NHÉ!")
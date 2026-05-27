from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import openpyxl
import os
import time
import re
import random
import subprocess
import socket
import sys
import io
import calendar

# ===================================================================================
# ── Timestamp print ──
# ===================================================================================
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
_orig_print = print
def print(*args, **kwargs):
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    _orig_print(f"[{ts}]", *args, **kwargs)
# ===================================================================================
# --- CONFIG ---
# ===================================================================================
def sleep_human(*args):
    # Phớt lờ các tham số thời gian cũ, ép chạy random siêu nhanh từ 0.01s đến 0.04s
    time.sleep(random.uniform(0.01, 0.04))
current_folder     = os.getcwd()
driver_path        = os.path.join(current_folder, "msedgedriver.exe")  # ← THÊM
excel_path         = os.environ.get("EXCEL_PATH", os.path.join(current_folder, "input_gia.xlsx"))
FILTER_POL         = os.environ.get("FILTER_POL", "").strip().upper()
FILTER_POD         = os.environ.get("FILTER_POD", "").strip().upper()
EDGE_EXE           = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
EDGE_DEBUG_PORT    = 9524
EDGE_USER_DATA_DIR = r"C:\edge_cma"  # ← profile riêng cho CMA
CMA_URL            = "https://www.cma-cgm.com/ebusiness/pricing/instant-quoting"
LOGIN_URL          = "https://www.cma-cgm.com/myCmaCgm/login"
LOGIN_EMAIL        = "celine@pio-logistics.vn"    # ← sửa lại
LOGIN_PASSWORD     = "Xvnt@686868"             # ← sửa lại

# ===================================================================================
# --- AUTO LAUNCH EDGE ---
# ===================================================================================
def is_port_open(port, host="127.0.0.1"):
    try:
        with socket.create_connection((host, port), timeout=1):
            return True
    except (OSError, ConnectionRefusedError):
        return False

def launch_edge_if_needed():
    if is_port_open(EDGE_DEBUG_PORT):
        print(f"[OK] Edge đã mở (port {EDGE_DEBUG_PORT}).")
        return True
    print("[INFO] Khởi động Edge...")
    try:
        subprocess.Popen([
            EDGE_EXE,
            f"--remote-debugging-port={EDGE_DEBUG_PORT}",
            f"--user-data-dir={EDGE_USER_DATA_DIR}",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-backgrounding-occluded-windows",
        ])
        for i in range(30):
            time.sleep(0.5)
            if is_port_open(EDGE_DEBUG_PORT):
                print(f"[OK] Edge sẵn sau {(i+1)*0.5:.1f}s.")
                return True
        print("[ERROR] Edge không khởi động được.")
        return False
    except FileNotFoundError:
        print(f"[ERROR] Không thấy Edge: {EDGE_EXE}")
        return False

def init_browser():
    opts = Options()
    opts.use_chromium = True
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{EDGE_DEBUG_PORT}")
    try:
        drv = webdriver.Edge(service=Service(driver_path), options=opts)
        drv.set_page_load_timeout(30)
        print("[OK] Kết nối Selenium thành công.")
        return drv
    except Exception as e:
        print(f"[ERROR] Không kết nối được Selenium: {e}")
        return None

# ===================================================================================
# --- AUTO LOGIN CMA ---
# ===================================================================================
def check_and_login(drv):
    """Truy cập Spot-On, nếu bị redirect sang trang auth thì đăng nhập."""
    print("[INFO] Truy cập trang Spot-On...")
    try:
        drv.get(CMA_URL)
    except TimeoutException:
        drv.execute_script("window.stop();")
    
    # --- FIX 13s: THAY time.sleep(4) BẰNG CHỜ URL THÔNG MINH ---
    try:
        WebDriverWait(drv, 5, poll_frequency=0.1).until(
            lambda d: "auth.cma-cgm.com" in d.current_url or "instant-quoting" in d.current_url or "ebusiness" in d.current_url
        )
    except:
        pass # Nếu lag quá 5s chưa đổi URL thì cứ đi tiếp để code dưới xử lý
        
    cur_url = drv.current_url

    if "auth.cma-cgm.com" in cur_url:
        print("[INFO] Bị chuyển hướng sang trang đăng nhập, tiến hành clear dữ liệu và login...")
        try:
            # 1. Xử lý Email
            email_xpath = "/html/body/main/section/div/div/section/form/fieldset/div[1]/input"
            email_inp = WebDriverWait(drv, 10).until(
                EC.presence_of_element_located((By.XPATH, email_xpath))
            )
            email_inp.click()
            time.sleep(0.2)
            email_inp.send_keys(Keys.CONTROL + "a")
            time.sleep(0.2)
            email_inp.send_keys(Keys.DELETE)
            email_inp.send_keys(LOGIN_EMAIL)
            time.sleep(0.5)

            # 2. Xử lý Password
            pwd_xpath = "/html/body/main/section/div/div/section/form/fieldset/div[2]/input[1]"
            pwd_inp = WebDriverWait(drv, 10).until(
                EC.presence_of_element_located((By.XPATH, pwd_xpath))
            )
            pwd_inp.click()
            time.sleep(0.2)
            pwd_inp.send_keys(Keys.CONTROL + "a")
            time.sleep(0.2)
            pwd_inp.send_keys(Keys.DELETE)
            pwd_inp.send_keys(LOGIN_PASSWORD)
            time.sleep(0.5)

            # 3. Bấm Log In
            btn_xpath = "/html/body/main/section/div/div/section/form/div/button"
            login_btn = WebDriverWait(drv, 5).until(
                EC.element_to_be_clickable((By.XPATH, btn_xpath))
            )
            login_btn.click()

            # Chờ web trả về lại trang báo giá
            print("[INFO] Đang chờ web quay lại trang Spot-On...")
            try:
                WebDriverWait(drv, 1).until(
                    lambda d: "instant-quoting" in d.current_url or "ebusiness/pricing" in d.current_url
                )
                print("[OK] Đăng nhập thành công và đã tự về trang Spot-On!")
            except Exception:
                # NẾU FAIL HOẶC TIMEOUT THÌ ÉP CHUYỂN HƯỚNG THEO LỆNH CỦA BẠN
                print("[WARN] Web không tự chuyển về Spot-On (hoặc bị lag). Tiến hành ép chuyển hướng...")
                drv.get(CMA_URL)
                time.sleep(5)
                print("[OK] Đã ép trình duyệt về lại trang Spot-On thành công!")
                
            return True

        except Exception as e:
            print(f"[ERROR] Quá trình nhập liệu login thất bại: {e}")
            print("[INFO] Thử ép chuyển hướng lại trang Spot-On lần cuối...")
            try:
                drv.get(CMA_URL)
                time.sleep(5)
                return True
            except:
                return False

    elif "instant-quoting" in cur_url or "ebusiness" in cur_url:
        print("[OK] Đã ở trang Spot-On, không cần đăng nhập.")
        return True
    else:
        print(f"[WARN] URL hiện tại không xác định: {cur_url}")
        return True

def ensure_on_cma_tab(drv):
    print("[DEBUG] BẮT ĐẦU KHÓA MỤC TIÊU VÀO TAB CMA (BỎ QUA TAB RÁC)...")
    try:
        all_tabs = drv.window_handles
        print(f"[DEBUG] Tổng số tab đang mở: {len(all_tabs)} -> Danh sách ID: {all_tabs}")
    except Exception as e:
        print(f"[ERROR] [DEBUG] Không lấy được window_handles: {e}")
        return

    if not all_tabs:
        print("[WARN] [DEBUG] Không có tab nào cả!")
        return

    target = None
    # 1. Quét từng tab để tìm tab CMA
    for h in all_tabs:
        try:
            drv.switch_to.window(h)
            cur_url = drv.current_url
            print(f"[DEBUG] Đang xét tab ID [{h[-6:]}] | URL hiện tại: {cur_url}")
            if "cma-cgm.com" in cur_url:
                target = h
                print(f"   -> [DEBUG] TÌM THẤY tab CMA! Đặt tab [{h[-6:]}] làm Target.")
                break
        except Exception as e:
            print(f"[ERROR] [DEBUG] Lỗi khi đọc tab {h[-6:]}: {e}")

    # 2. Nếu không thấy trang CMA nào, lấy tab đầu tiên làm gốc
    if not target:
        target = all_tabs[0]
        print(f"[WARN] [DEBUG] Không thấy URL CMA nào! Đặt tab đầu tiên [{target[-6:]}] làm Target mặc định.")

    # 3. Trở về tab Target và phớt lờ hoàn toàn các tab khác (Bỏ lệnh close)
    try:
        print(f"[DEBUG] Đang khóa mục tiêu vào tab Target [{target[-6:]}]...")
        drv.switch_to.window(target)
        print(f"[DEBUG] HOÀN TẤT KHÓA MỤC TIÊU! Đang ở URL: {drv.current_url}")
    except Exception as e:
        print(f"[FATAL] [DEBUG] LỖI MẤT SESSION KHI TRỞ VỀ TARGET: {e}")
# ===================================================================================
# --- 1. KHỞI ĐỘNG ---
# ===================================================================================
from selenium.common.exceptions import TimeoutException

if not launch_edge_if_needed():
    print("[ERROR] Không khởi động được Edge. Dừng.")
    exit()

driver = init_browser()
if not driver:
    print("[ERROR] Không kết nối được Selenium. Dừng.")
    exit()

# Bùa tàng hình
try:
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
except:
    pass

ensure_on_cma_tab(driver)

if not check_and_login(driver):
    print("[ERROR] Đăng nhập thất bại. Dừng.")
    exit()


print("[OK] CMA Bot sẵn sàng!")

# ===================================================================================
# --- 2. CÁC HÀM XỬ LÝ NHẬP LIỆU ---
# ===================================================================================
def check_and_kill_popup():
    print("   -> Kiểm tra Popup...")

    # 1. Popup cũ
    try:
        remind_btn = WebDriverWait(driver, 1.5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Remind Me Later')]"))
        )
        print("      [CÓ POPUP CŨ] -> Bấm Remind Me Later!")
        driver.execute_script("arguments[0].click();", remind_btn)
        time.sleep(random.uniform(1.0, 1.5))
    except:
        pass

    # 2. Popup Concierge
    try:
        concierge_close_btn = driver.find_element(By.CSS_SELECTOR, "button.mec-me-popup__close")
        if concierge_close_btn.is_displayed():
            print("      [CÓ POPUP CONCIERGE] -> Đóng popup!")
            driver.execute_script("arguments[0].click();", concierge_close_btn)
            time.sleep(random.uniform(1.0, 1.5))
    except:
        pass

    # 3. Popup lạ
    try:
        generic_close_btn = driver.find_element(
            By.CSS_SELECTOR, "button[aria-label='Close'], button[aria-label='Close dialog']"
        )
        if generic_close_btn.is_displayed():
            print("      [CÓ POPUP LẠ] -> Đóng popup!")
            driver.execute_script("arguments[0].click();", generic_close_btn)
            time.sleep(random.uniform(1.0, 1.5))
    except:
        pass


def perform_modify_search_action():
    print("   -> Bấm Modify Search...")
    try:
        modify_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH,
                "//button[contains(., 'Modify Search')] | //a[contains(., 'Modify Search')]"))
        )
        driver.execute_script("arguments[0].click();", modify_btn)

        # Chờ form hiện ra thay vì sleep cứng
        WebDriverWait(driver, 8).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//button[contains(., 'Get My Quote') or @id='SearchQuote']"))
        )
        check_and_kill_popup()
        return True
    except:
        return False
    


def smart_update_field(target_type, new_value):
    print(f"   -> Đang sửa {target_type} thành: {new_value}...")
    try:
        wrapper = None
        if target_type == 'POD':
            try:
                dest_div = driver.find_element(By.CSS_SELECTOR, ".destination-search")
                wrapper = dest_div.find_element(By.CSS_SELECTOR, ".selected-value-wrapper")
            except: pass
        else:
            try:
                wrappers = driver.find_elements(By.CSS_SELECTOR, ".selected-value-wrapper")
                if wrappers: wrapper = wrappers[0]
            except: pass

        if wrapper and wrapper.is_displayed():
            driver.execute_script("arguments[0].click();", wrapper)
            try:
                WebDriverWait(driver, 2, poll_frequency=0.05).until(
                    EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Name / Code / Port']")))
            except: time.sleep(0.3)

        all_inputs = driver.find_elements(By.XPATH, "//input[@placeholder='Name / Code / Port']")
        visible_inputs = [inp for inp in all_inputs if inp.is_displayed()]
        if not visible_inputs: return False
        target_input = visible_inputs[-1] if target_type == 'POD' else visible_inputs[0]

        driver.execute_script("arguments[0].click();", target_input)
        time.sleep(0.15)
        target_input.send_keys(Keys.CONTROL + "a")
        target_input.send_keys(Keys.DELETE)
        if target_input.get_attribute("value"): target_input.clear()

        # Xử lý Logic tách PORT và RAMP
        search_key = new_value
        pick_ramp = False
        pick_port = False
        if new_value == "VNSGN-RAMP":
            search_key = "VNSGN"
            pick_ramp = True
        elif new_value == "VNSGN-PORT":
            search_key = "VNSGN"
            pick_port = True
        elif "VNSGN" in new_value:
            pick_ramp = True # Mặc định tuyến xa lấy RAMP (hoặc Vũng Tàu sau đó)

        target_input.send_keys(search_key)

        try:
            WebDriverWait(driver, 4, poll_frequency=0.05).until(
                EC.visibility_of_any_elements_located((By.CSS_SELECTOR, "li.place-suggestion")))
        except: time.sleep(0.8)

        try:
            suggestions = driver.find_elements(By.CSS_SELECTOR, "li.place-suggestion")
            visible = [s for s in suggestions if s.is_displayed()]
            
            chosen = None
            if search_key == "VNSGN":
                if pick_ramp:
                    chosen = next((s for s in visible if "RAMP" in s.text.upper()), None)
                    if not chosen and len(visible) > 1: chosen = visible[1]
                elif pick_port:
                    chosen = next((s for s in visible if "RAMP" not in s.text.upper()), None)
                    if not chosen and visible: chosen = visible[0]
                else:
                    chosen = visible[1] if len(visible) > 1 else (visible[0] if visible else None)
            else:
                chosen = visible[0] if visible else None

            if chosen: driver.execute_script("arguments[0].click();", chosen)
            else: target_input.send_keys(Keys.ENTER)
        except:
            target_input.send_keys(Keys.ENTER)

        try:
            WebDriverWait(driver, 1.5, poll_frequency=0.05).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, "li.place-suggestion")))
        except: time.sleep(0.3)

        return True

    except Exception as e:
        print(f"   -> [ERROR] smart_update_field thất bại: {e}")
        return False


def select_port_full(element, text):
    try:
        driver.execute_script("arguments[0].click();", element)
        element.send_keys(Keys.CONTROL + "a"); time.sleep(0.1)
        element.send_keys(Keys.DELETE)
        
        search_key = text
        pick_ramp = False
        pick_port = False
        if text == "VNSGN-RAMP":
            search_key = "VNSGN"
            pick_ramp = True
        elif text == "VNSGN-PORT":
            search_key = "VNSGN"
            pick_port = True
        elif "VNSGN" in text:
            pick_ramp = True

        element.send_keys(search_key)

        try:
            WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.place-suggestion")))
            WebDriverWait(driver, 5).until(EC.visibility_of_any_elements_located((By.CSS_SELECTOR, "li.place-suggestion")))
        except:
            sleep_human(1.5, 2.0)

        try:
            suggestions = driver.find_elements(By.CSS_SELECTOR, "li.place-suggestion")
            visible = [s for s in suggestions if s.is_displayed()]
            
            chosen = None
            if search_key == "VNSGN":
                if pick_ramp:
                    chosen = next((s for s in visible if "RAMP" in s.text.upper()), None)
                    if not chosen and len(visible) > 1: chosen = visible[1]
                elif pick_port:
                    chosen = next((s for s in visible if "RAMP" not in s.text.upper()), None)
                    if not chosen and visible: chosen = visible[0]
                else:
                    chosen = visible[1] if len(visible) > 1 else (visible[0] if visible else None)
            else:
                chosen = visible[0] if visible else None

            if chosen: driver.execute_script("arguments[0].click();", chosen)
            else: element.send_keys(Keys.ENTER)
        except:
            element.send_keys(Keys.ENTER)
    except:
        pass

def select_vung_tau_mandatory():
    print("   -> Đang móc Vũng Tàu (JS Bypass)...")
    try:
        pol_inp = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Select']"))
        )
        driver.execute_script("arguments[0].click();", pol_inp)

        # Chờ dropdown hiện
        try:
            WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "li.el-select-dropdown__item"))
            )
        except:
            time.sleep(random.uniform(1.0, 1.5))

        js_find_vung_tau = """
        var items = document.querySelectorAll('li.el-select-dropdown__item');
        for(var i=0; i<items.length; i++) {
            var parent = items[i].closest('.el-select-dropdown');
            if(parent && parent.style.display !== 'none') {
                var text = (items[i].textContent || items[i].innerText).toUpperCase();
                if(text.indexOf('VUNG TAU') !== -1 || text.indexOf('VNVUT') !== -1) {
                    items[i].click();
                    return text.trim();
                }
            }
        }
        return null;
        """
        result = driver.execute_script(js_find_vung_tau)

        if not result:
            pol_inp.send_keys("VUNG")
            time.sleep(random.uniform(1.0, 1.5))
            items2 = driver.find_elements(By.CSS_SELECTOR, "li.el-select-dropdown__item")
            for item in items2:
                if item.is_displayed() and (
                    "VUNG TAU" in item.text.upper() or "VNVUT" in item.text.upper()
                ):
                    driver.execute_script("arguments[0].click();", item)
                    break
            else:
                pol_inp.send_keys(Keys.ENTER)
        # ✅ THÊM: Chờ React re-render xong sau khi chọn Vung Tau
        # Đợi dropdown biến mất = form đã stabilize
        try:
            WebDriverWait(driver, 2, poll_frequency=0.05).until_not(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "li.el-select-dropdown__item"))
            )
        except:
            time.sleep(0.4)
    except Exception as e:
        pass


def handle_pod_selection_popup(prefer_port="JEDDAH"):
    """
    Xử lý dropdown 'Select' xuất hiện sau khi chọn port inland (VD: Riyadh).
    Ưu tiên chọn prefer_port (Jeddah), fallback về option đầu tiên.
    """
    try:
        # Chờ input Select xuất hiện tối đa 4s
        dropdown = WebDriverWait(driver, 4).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input.el-input__inner[placeholder='Select']")
            )
        )
        print(f"   -> Phát hiện dropdown POD → mở lên chọn {prefer_port}...")

        driver.execute_script("arguments[0].click();", dropdown)

        # Chờ options hiện ra
        try:
            WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "li.el-select-dropdown__item"))
            )
        except:
            time.sleep(random.uniform(1.0, 1.5))

        options = driver.find_elements(By.CSS_SELECTOR, "li.el-select-dropdown__item")
        visible = [o for o in options if o.is_displayed()]
        print(f"      Có {len(visible)} options: {[o.text.strip() for o in visible[:5]]}")

        target = next(
            (o for o in visible if prefer_port.upper() in o.text.upper()), None
        )

        if target:
            print(f"      → Chọn: {target.text.strip()}")
            driver.execute_script("arguments[0].click();", target)
        elif visible:
            print(f"      → Không thấy {prefer_port}, chọn đầu tiên: {visible[0].text.strip()}")
            driver.execute_script("arguments[0].click();", visible[0])
        else:
            print(f"      ⚠️ Không có options nào hiện ra")

        sleep_human(0.5, 0.8)

    except Exception:
        pass

def debug_vessel_structure():
    """Debug: In ra cấu trúc HTML của vessel để kiểm tra"""
    try:
        vessel_items = driver.find_elements(By.CSS_SELECTOR, "li.more-infos.vessel")
        if vessel_items:
            first = vessel_items[0]
            html = driver.execute_script("return arguments[0].innerHTML;", first)
            print(f"\n[DEBUG] HTML cấu trúc vessel:\n{html[:500]}...\n")
    except:
        pass

def debug_vessel_html():
    """In ra HTML cấu trúc vessel để kiểm tra"""
    try:
        vessel_items = driver.find_elements(By.CSS_SELECTOR, "li.more-infos.vessel")
        if vessel_items:
            first = vessel_items[0]
            lis = first.find_elements(By.XPATH, ".//ul/li")
            print(f"      [DEBUG] Số <li>: {len(lis)}")
            for i, li in enumerate(lis[:3]):
                html = driver.execute_script("return arguments[0].innerHTML;", li)
                print(f"      [DEBUG] Li #{i}: {html[:200]}...")
    except Exception as e:
        print(f"      [DEBUG] Error: {e}")

def select_commodity_refresh():
    print("   -> Chọn lại Commodity...")
    try:
        # Thử nhiều selector để tìm dropdown Commodity
        comm = None
        selectors = [
            (By.ID, "DdlCommodity"),
            (By.XPATH, "//div[contains(@class,'commodity')]//input"),
            (By.XPATH, "//*[contains(@placeholder,'commodity') or contains(@placeholder,'Commodity')]"),
            (By.XPATH, "//label[contains(.,'Commodity')]/following::input[1]"),
            (By.XPATH, "//label[contains(.,'Commodity')]/following::div[contains(@class,'el-select')][1]//input"),
        ]
        for by, sel in selectors:
            try:
                el = WebDriverWait(driver, 2).until(EC.presence_of_element_located((by, sel)))
                if el.is_displayed():
                    comm = el
                    print(f"      -> Tìm thấy Commodity bằng selector: {sel}")
                    break
            except:
                continue

        if not comm:
            print("      ⚠️ Không tìm thấy ô Commodity!")
            return

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", comm)
        driver.execute_script("arguments[0].click();", comm)
        time.sleep(random.uniform(0.3, 0.5))

        # Chờ dropdown options hiện ra
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "li.el-select-dropdown__item"))
            )
        except:
            # Thử click lần 2 nếu dropdown chưa mở
            driver.execute_script("arguments[0].click();", comm)
            time.sleep(random.uniform(0.5, 0.8))

        # Ưu tiên chọn "Freight All Kinds", fallback lấy item đầu tiên
        try:
            fak_xpath = "//li[contains(@class,'el-select-dropdown__item') and contains(.,'Freight All Kinds')]"
            fak_item = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, fak_xpath))
            )
            driver.execute_script("arguments[0].click();", fak_item)
            print("      -> [OK] Đã chọn Freight All Kinds!")
        except:
            js_force = """
            var items = document.querySelectorAll('li.el-select-dropdown__item');
            for(var i=0; i<items.length; i++) {
                if(items[i].offsetParent !== null) {
                    items[i].click();
                    return items[i].innerText.trim();
                }
            }
            return null;
            """
            result = driver.execute_script(js_force)
            if result:
                print(f"      -> [OK] Không thấy FAK, đã chọn: {result}")
            else:
                print("      ⚠️ Dropdown không hiện options!")

        time.sleep(random.uniform(0.2, 0.3))

    except Exception as e:
        print(f"      ⚠️ Lỗi chốt Commodity: {e}")

def get_main_vessel_and_voyage(search_context=None):
    """
    Lấy tàu chính (Leg 1) từ route detail.
    Format: <TÊN TÀU> (<SERVICE CODE>)
    search_context: card element để tìm kiếm (tránh lấy nhầm từ card khác)
    """
    try:
        # ✅ TÌM TRONG CARD ELEMENT CỤ THỀ (TỪ PARAMETER)
        if search_context is None:
            search_context = driver
        
        vessel_items = search_context.find_elements(By.CSS_SELECTOR, "li.more-infos.vessel")
        
        if not vessel_items:
            print(f"      ⚠️ Không tìm thấy vessel items")
            return "TBA", "N/A"
        
        # ✅ LUÔN LẤYA LI.MORE-INFOS.VESSEL ĐẦU TIÊN TRONG CONTEXT
        first_vessel_li = vessel_items[0]
        
        # ✅ LẤY TÊN TÀU + SERVICE CODE (dùng dt text thay vì index cứng)
        vessel_name = "TBA"
        service_code = "N/A"
        try:
            lis = first_vessel_li.find_elements(By.XPATH, ".//ul/li")
            for li_item in lis:
                dts = li_item.find_elements(By.XPATH, ".//dt")
                dds = li_item.find_elements(By.XPATH, ".//dd")
                for idx_dt, dt in enumerate(dts):
                    dt_text = dt.text.strip().upper()
                    if dt_text == "VESSEL" and idx_dt < len(dds):
                        vessel_name = dds[idx_dt].text.strip()
                    elif dt_text == "SERVICE" and idx_dt < len(dds):
                        service_code = dds[idx_dt].text.strip()
        except:
            pass
        
        result = f"{vessel_name} ({service_code})"
        print(f"      ✅ Tàu chính: {result}")
        return vessel_name, service_code
        
    except Exception as e:
        print(f"      ⚠️ Lỗi bóc vessel tổng quát: {type(e).__name__}")
        return "TBA", "N/A"
              
        
def pick_date_plus_7():
    print("   -> Chọn ngày hôm nay +7...")
    try:
        today = datetime.now()
        target = today + timedelta(days=7)
        target_day = str(target.day)
        is_next = target.month != today.month

        # Mở lịch
        opened = driver.execute_script("""
            var inp = document.getElementById('DepartureFrom');
            if(!inp) return false;
            var parent = inp.closest('.el-date-editor') || inp.parentElement;
            var icon = parent ? parent.querySelector('i.el-icon, i.el-input__icon') : null;
            if(icon) { icon.click(); return true; }
            inp.click(); inp.dispatchEvent(new MouseEvent('click', {bubbles:true}));
            return true;
        """)

        if not opened:
            print(f"      ⚠️ Không tìm thấy icon calendar")
            return

        # Chờ bảng lịch hiện ra
        try:
            WebDriverWait(driver, 3, poll_frequency=0.05).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "table.el-date-table"))
            )
            time.sleep(0.2)
        except:
            print(f"      ⚠️ Calendar không mở kịp")
            return

        # Nếu là tháng sau, click vào ô chứa class 'next-month'. Cùng tháng thì click 'available'
        js = f"""
        var d='{target_day}';
        var isNext = {'true' if is_next else 'false'};
        var selector = isNext ? 'td.next-month span.el-date-table-cell__text' : 'td.available span.el-date-table-cell__text';
        var s = document.querySelectorAll(selector);
        
        for(var i=0; i<s.length; i++){{
            if(s[i].innerText.trim() === d && s[i].offsetParent !== null){{
                s[i].click();
                return true;
            }}
        }}
        return false;
        """
        result = driver.execute_script(js)

        # Fallback: Đề phòng web không nạp thẻ next-month, ép click nút Next Arrow rồi kiếm lại
        if not result and is_next:
            try:
                next_btn = driver.find_element(By.CSS_SELECTOR, "button.el-picker-panel__icon-btn.arrow-right")
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(0.3)
                result = driver.execute_script(js.replace("td.next-month", "td.available"))
            except:
                pass

        print(f"      -> Chọn ngày {target.strftime('%d-%b-%Y')}: {'✅ OK' if result else '⚠️ Không tìm thấy ngày'}")

    except Exception as e:
        print(f"      ⚠️ Lỗi pick_date: {e}")  
    

def select_containers():
    print("   -> Chọn Container...")
    try:
        driver.execute_script("window.scrollBy(0, 400);")
        time.sleep(random.uniform(0.3, 0.5))
    except:
        pass
    for cont in ["20' Dry Standard", "40' Dry Standard", "40' Dry High Cube"]:
        try:
            li = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//ul[contains(@class,'container')]//li[contains(., \"{cont}\")]"))
            )
            if "is-checked" not in li.get_attribute("class"):
                driver.execute_script("arguments[0].click();", li)
                sleep_human(0.2, 0.4)
            inp = li.find_elements(By.TAG_NAME, "input")[-1]
            driver.execute_script("arguments[0].click();", inp)
            inp.send_keys(Keys.CONTROL + "a")
            inp.send_keys(Keys.DELETE)
            inp.send_keys("22222")
        except:
            pass

def ensure_containers_filled():
    print("   -> Đảm bảo Container đã điền 22222kgs...")
    for cont in ["20' Dry Standard", "40' Dry Standard", "40' Dry High Cube"]:
        try:
            li = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//ul[contains(@class,'container')]//li[contains(., \"{cont}\")]"))
            )
            # Đảm bảo checkbox đã được tick
            if "is-checked" not in li.get_attribute("class"):
                driver.execute_script("arguments[0].click();", li)
                sleep_human(0.2, 0.4)

            # Chỉ điền lại nếu giá trị CHƯA đúng
            inp = li.find_elements(By.TAG_NAME, "input")[-1]
            current_val = (inp.get_attribute("value") or "").strip()
            if current_val != "22222":
                print(f"      -> Ô [{cont}] đang là '{current_val}', đang điền lại 22222...")
                driver.execute_script("arguments[0].click();", inp)
                inp.send_keys(Keys.CONTROL + "a")
                inp.send_keys(Keys.DELETE)
                inp.send_keys("22222")
            else:
                print(f"      -> [{cont}] ✅ Đã có 22222, bỏ qua.")
        except Exception as e:
            print(f"      ⚠️ Không kiểm tra được [{cont}]: {e}")

# ===================================================================================
# --- 3. BỘ NÃO LỌC ETD & BÓC GIÁ (9 QUY TẮC VÀNG) ---
# ===================================================================================
def parse_cma_date(date_str):
    try:
        return datetime.strptime(date_str, "%d-%b-%Y")
    except:
        return None

def calculate_cma_validity(last_etd):
    """Tính mốc Valid (7, 14, 21, cuối tháng) dựa trên ETD cuối."""
    import calendar
    year, month = last_etd.year, last_etd.month
    last_day = calendar.monthrange(year, month)[1]
    
    milestones = [7, 14, 21, last_day]
    for m in milestones:
        if m >= last_etd.day:
            dt = datetime(year, month, m)
            return f"{dt.day}-{dt.strftime('%b')}"
    return f"{last_day}-{last_etd.strftime('%b')}"

def calculate_cma_validity(last_etd):
    """Tính mốc Valid (7, 14, 21, cuối tháng) dựa trên ETD cuối."""
    import calendar
    year, month = last_etd.year, last_etd.month
    last_day = calendar.monthrange(year, month)[1]
    
    milestones = [7, 14, 21, last_day]
    for m in milestones:
        if m >= last_etd.day:
            dt = datetime(year, month, m)
            return f"{dt.day}-{dt.strftime('%b')}"
    return f"{last_day}-{last_etd.strftime('%b')}"

def scrape_multi_etd_and_save(row_index, ws, pol_text_excel):
    print("8. Đang quét danh sách tàu và bóc tách dữ liệu chi tiết...")
    
    # --- THÊM BƯỚC: POLLING JAVASCRIPT ĐỂ CLICK AVAILABLE ---
    try:
        print("      -> Đang Polling tìm và click chọn 'Available solutions only'...")
        
        js_polling = """
        // Tìm trực tiếp input có value là 'OnlyAvailable'
        var input = document.querySelector('input[value="OnlyAvailable"]');
        if (input) {
            var label = input.closest('label');
            if (label) {
                if (!label.className.includes('is-checked')) {
                    label.click();
                    return 'CLICKED';
                }
                return 'ALREADY_CHECKED';
            }
        }
        
        // Fallback: Tìm thẻ label chứa chữ 'Available'
        var labels = document.querySelectorAll('label.el-radio');
        for(var i=0; i<labels.length; i++) {
            if(labels[i].innerText.trim() === 'Available') {
                if(!labels[i].className.includes('is-checked')) {
                    labels[i].click();
                    return 'CLICKED';
                }
                return 'ALREADY_CHECKED';
            }
        }
        // Trả về chuỗi rỗng để WebDriverWait tiếp tục quét
        return ''; 
        """

        # Chạy Polling liên tục tối đa 5s, mỗi 0.2s quét JS 1 lần. 
        # Nếu JS trả về 'CLICKED' hoặc 'ALREADY_CHECKED' (Truthy) thì vòng lặp dừng ngay lập tức.
        result = WebDriverWait(driver, 5, poll_frequency=0.2).until(
            lambda d: d.execute_script(js_polling)
        )
        
        if result == 'CLICKED':
            print("      -> Đã click chọn bộ lọc 'Available' thành công!")
            # Chỉ sleep 1 nhịp ngắn (1-1.5s) SAU KHI click để chờ danh sách web lọc bớt card Sold Out
            time.sleep(1.5) 
        elif result == 'ALREADY_CHECKED':
            print("      -> Form đã tự động chọn sẵn 'Available', đi tiếp...")
            
    except TimeoutException:
        print("      [WARN] Hết 5s Polling không tìm thấy nút Available, tiếp tục kịch bản...")
    except Exception as e:
        print(f"      [WARN] Lỗi không xác định khi Polling Available: {type(e).__name__}")
    # ---------------------------------------------S

    try:
       # --- MỎ NEO THỜI GIAN: ÉP CHỜ TẤT CẢ CÁC GIÁ LOAD XONG ---
        print("      -> [WAIT] Đang ép bot đứng chờ web tải giá tiền từ server...")
        try:
            # Dùng 'all' để ép bot chờ tất cả các card đều nhả chữ USD hoặc SOLD OUT
            WebDriverWait(driver, 15, poll_frequency=0.5).until(
                lambda d: (
                    len(d.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")) > 0 and
                    all(
                        re.search(r"([\d,]+)\s*(?:USD|US\$|\$)|SOLD OUT", c.text, re.IGNORECASE)
                        for c in d.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")
                    )
                )
            )
            time.sleep(1.0) # Nghỉ thêm 1 giây cho React bung hết text
        except:
            print("      [WARN] Hết 15s chờ, một số card có thể bị lỗi không tải được giá.")

        cards = driver.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")
        if not cards:
            print("      ⚠️ Sold Out toàn bộ (Không có chuyến tàu nào)!")
            ws.cell(row=row_index, column=6).value = "Sold Out"
            try: wb.save(excel_path)
            except: pass
            return "SOLD_OUT"

        # ══════════════════════════════════════════════════════════════════════════════
        # BƯỚC 1: QUÉT TẤT CẢ CARD, LỌC THEO 9 QUY TẮC VÀNG (BẢN GẮN RADAR)
        # ══════════════════════════════════════════════════════════════════════════════
        all_options = []
        for idx, card in enumerate(cards[:8]):
            try:
                print(f"      [DEBUG] Đang xét Card #{idx+1}...")
                
                # 1. Bóc Date
                try:
                    etd_elm = card.find_element(By.CSS_SELECTOR, "span.date, .date")
                    raw_etd = etd_elm.text.strip()
                    etd_str = raw_etd.split(",")[1].strip() if "," in raw_etd else raw_etd
                    etd_date = parse_cma_date(etd_str)
                    
                    if not etd_date:
                        clean_str = etd_str.replace("-", " ")
                        try:
                            etd_date = datetime.strptime(clean_str, "%d %b %Y")
                        except: pass
                except Exception as e:
                    print(f"         -> Bỏ qua: Không tìm thấy thẻ Ngày/Tháng ({type(e).__name__})")
                    continue

                if not etd_date:
                    print(f"         -> Bỏ qua: Lỗi format ngày '{raw_etd}'")
                    continue

                # 2. Bóc Transit
                try:
                    transit_elm = card.find_element(By.XPATH, ".//*[contains(@class, 'transit')]")
                    match_tr = re.search(r"(\d+)", transit_elm.text)
                    transit_val = int(match_tr.group(1)) if match_tr else 99
                except:
                    transit_val = 99

                # 3. Bóc Transshipment
                ts_text = "DIRECT"
                try:
                    ts_div = card.find_element(By.XPATH, ".//div[contains(@class,'transit') and (contains(@class,'transshipment') or contains(@class,'direct'))]")
                    if "direct" not in ts_div.get_attribute("class").lower():
                        raw_ts = ts_div.text.replace("via", "").strip()
                        ports  = [p.split(",")[0].strip() for p in raw_ts.split("•")]
                        ts_text = " + ".join(ports)
                except: pass

                # 4. Bóc Giá (Dùng JS ép đọc text)
                raw_text_upper = driver.execute_script("return arguments[0].innerText || arguments[0].textContent;", card).upper()
                price_match = re.search(r"([\d\s,]+)\s*(?:USD|US\$|\$|EUR)", raw_text_upper)
                fallback_match = re.search(r"PER\s*(?:20|40)[A-Z0-9]*\s*([\d\s,]+)", raw_text_upper)
                
                if price_match:
                    current_price = int(re.sub(r"[^\d]", "", price_match.group(1)))
                elif fallback_match:
                    current_price = int(re.sub(r"[^\d]", "", fallback_match.group(1)))
                else:
                    current_price = 0

                # KIỂM TRA ĐIỀU KIỆN LOẠI BỎ
                if "SOLD OUT" in raw_text_upper and current_price == 0:
                    print("         -> Bỏ qua: Card này bị dính chữ SOLD OUT.")
                    continue
                
                if current_price == 0:
                    print("         -> Bỏ qua: Price = 0 (Giá chưa tải hoặc ẩn).")
                    continue

                print(f"         => HỢP LỆ! Ghi nhận: Giá {current_price}, ETD {etd_str}, Transit {transit_val} Ngày")
                
                all_options.append({
                    'date': etd_date, 'price': current_price, 'transit': transit_val,
                    'card_idx': idx, 'ts_port': ts_text
                })
            except Exception as ex:
                print(f"         -> Lỗi văng code không xác định: {ex}")
                continue

        if not all_options:
            print("      ⚠️ Toàn bộ bảng giá đều đã Sold Out hoặc không lấy được dữ liệu!")
            ws.cell(row=row_index, column=6).value = "Sold Out"
            try: wb.save(excel_path)
            except: pass
            return "SOLD_OUT"

        # --- BƯỚC 1.5: TÌM GIÁ ĐÁY VÀ VỨT BỎ GIÁ CAO ---
        min_price = min(opt['price'] for opt in all_options)
        best_price_options = [opt for opt in all_options if opt['price'] == min_price]

        # Lọc cùng ngày (trong nhóm giá rẻ nhất) → giữ transit ngắn nhất
        unique_dates = {}
        for opt in best_price_options:
            d = opt['date']
            if d not in unique_dates:
                unique_dates[d] = opt
            else:
                if opt['transit'] < unique_dates[d]['transit']:
                    unique_dates[d] = opt

        sorted_opts = sorted(unique_dates.values(), key=lambda x: x['date'])
        e_min = sorted_opts[0]['date']

        # Lọc cửa sổ 9 ngày, cách nhau >= 2 ngày, tối đa 3 card
        valid_opts = []
        for opt in sorted_opts:
            if (opt['date'] - e_min).days > 9:
                break
            if not valid_opts:
                valid_opts.append(opt)
            else:
                if (opt['date'] - valid_opts[-1]['date']).days >= 2:
                    valid_opts.append(opt)
            if len(valid_opts) == 3:
                break

        if not valid_opts:
            return

        # ══════════════════════════════════════════════════════════════════════════════
        # BƯỚC 2: BÓC CHI TIẾT TỪNG CARD (VESSEL, FREE TIME)
        #         + LẤY GIÁ TỪ CARD ĐẦU TIÊN
        # ══════════════════════════════════════════════════════════════════════════════
        vessel_entries = []
        ts_entries     = []
        free_time_pod  = "N/A"
        total_20 = total_40 = total_40hc = 0
        othc_included  = False
        ows_found      = False
        manifest_fee_found = False

        for i, opt in enumerate(valid_opts):
            idx = opt['card_idx']

            fresh_cards = driver.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")
            if idx >= len(fresh_cards):
                continue
            card_el = fresh_cards[idx]

            print(f"   -> Đang bóc chi tiết Card #{idx + 1}...")

            # Click mở Details
            try:
                # Tăng thời gian chờ từ 2s lên 8s để web có đủ thời gian thở và render nút
                details_btn = WebDriverWait(driver, 8, poll_frequency=0.2).until(
                    lambda d: card_el.find_element(
                        By.XPATH, ".//label[contains(@class,'o-button') and contains(.,'Details')]")
                )
                switches  = card_el.find_elements(By.XPATH, ".//input[@data-role='switch']")
                is_opened = switches and switches[0].is_selected()

                if not is_opened:
                    driver.execute_script(
                        "arguments[0].scrollIntoView({behavior:'instant',block:'center'});"
                        "window.scrollBy(0,-100);", details_btn)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", details_btn)
                    
                    # Chờ một nhịp để bảng Details thực sự bung ra
                    time.sleep(1.0)
                    
                    # ✅ LẤY FRESH CARD SAU KHI CLICK DETAILS
                    fresh_cards = driver.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")
                    if idx < len(fresh_cards):
                        card_el = fresh_cards[idx]

            except Exception as e:
                print(f"      ⚠️ Nút Details load chậm ({type(e).__name__}). Đang dùng JS ép click...")
                try:
                    js_force_click = f"""
                    var cards = document.querySelectorAll('article.card-route-horizontal');
                    if(cards.length > {idx}) {{
                        var labels = cards[{idx}].querySelectorAll('label');
                        for(var j=0; j<labels.length; j++) {{
                            if(labels[j].innerText.includes('Details')) {{
                                labels[j].scrollIntoView({{behavior:'instant',block:'center'}});
                                window.scrollBy(0,-100);
                                labels[j].click();
                                return true;
                            }}
                        }}
                    }}
                    return false;
                    """
                    if driver.execute_script(js_force_click):
                        time.sleep(1.5)
                    else:
                        raise Exception("JS không tìm thấy nút Details")
                except Exception as ex:
                    print(f"      ❌ Bất lực với nút Details: {ex}. Bỏ qua card này.")
                    # CÚ PHÁP MỚI CHO FALLBACK
                    vessel_entries.append(
                        f"TBA / ETD: {opt['date'].day}-{opt['date'].strftime('%b')} "
                        f"/ Transit time: {opt['transit']} Days / Transshipment: {opt['ts_port']}")
                    ts_entries.append(opt['ts_port'])
                    continue

            # Lấy lại card tươi sau khi React bung tab
            fresh_cards = driver.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")
            card_el = fresh_cards[idx]

            # ── GIÁ: Chỉ lấy từ card đầu tiên ──────────────────────
            if i == 0:
                try:
                    print("   -> Đang bóc bảng giá chi tiết...")

                    # Chờ bảng giá xuất hiện
                    try:
                        WebDriverWait(driver, 8, poll_frequency=0.1).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.el-table__row"))
                        )
                        time.sleep(0.3)
                    except:
                        time.sleep(2.0)

                    # Nếu vẫn không thấy → thử click Details lại một lần
                    if not driver.find_elements(By.CSS_SELECTOR, "tr.el-table__row"):
                        driver.execute_script("arguments[0].click();", details_btn)
                        try:
                            WebDriverWait(driver, 6, poll_frequency=0.1).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "tr.el-table__row"))
                            )
                        except:
                            time.sleep(2.0)

                    # Expand các dòng nhóm (Ocean Freight + Charges payable)
                    try:
                        expand_rows = driver.find_elements(By.XPATH,
                            "//tr[.//td[contains(.,'Charges payable as per freight')] "
                            "or .//td[contains(.,'Ocean Freight')]]")
                        for r in expand_rows:
                            try:
                                icon = r.find_element(By.CSS_SELECTOR, ".el-table__expand-icon")
                                if "expanded" not in icon.get_attribute("class"):
                                    driver.execute_script("arguments[0].click();", icon)
                                    time.sleep(0.3)
                            except: pass
                    except: pass

                    # JS quét bảng giá
                    # JS quét bảng giá
                    js_price = """
                    var result = {ocean:[0,0,0], surcharge:[0,0,0], ows:[0,0,0], othc_included:false, manifest_fee_found:false};

                    var tableBody = document.querySelector('.el-table__body tbody');
                    if(tableBody) {
                        var fullText = (tableBody.innerText || tableBody.textContent).toUpperCase();
                        if(fullText.includes('TERMINAL HANDLING CHARGE (OTHC)') ||
                           fullText.includes('(OTHC) AT ORIGIN')) {
                            result.othc_included = true;
                        }
                        // Bắt từ khóa Advanced Manifest (Trung, Nhật) hoặc ENS (Châu Âu)
                    if(fullText.includes('ADVANCED MANIFEST DECLARATION') || 
                       fullText.includes('ENTRY SUMMARY DECLARATION') || 
                       fullText.includes(' ENS ') || 
                       fullText.includes('ENS SURCHARGE')) {
                        result.manifest_fee_found = true;
                    }
                    }

                    var rows = document.querySelectorAll('.el-table__body tbody tr');
                    for(var i=0; i<rows.length; i++) {
                        var cols = rows[i].querySelectorAll('td');
                        if(cols.length >= 5) {
                            var name  = (cols[1].textContent || cols[1].innerText || "").trim().toUpperCase();
                            var v20   = parseInt((cols[2].textContent || "").replace(/[^0-9]/g,'')) || 0;
                            var v40   = parseInt((cols[3].textContent || "").replace(/[^0-9]/g,'')) || 0;
                            var v40hc = parseInt((cols[4].textContent || "").replace(/[^0-9]/g,'')) || 0;

                            if(name.indexOf('OCEAN FREIGHT') !== -1) {
                                if(v20>0) result.ocean[0]=v20;
                                if(v40>0) result.ocean[1]=v40;
                                if(v40hc>0) result.ocean[2]=v40hc;
                            }
                            else if(name.indexOf('CHARGES PAYABLE AS PER FREIGHT') !== -1) {
                                if(v20>0) result.surcharge[0]=v20;
                                if(v40>0) result.surcharge[1]=v40;
                                if(v40hc>0) result.surcharge[2]=v40hc;
                            }
                            else if(name.indexOf('OVERWEIGHT SURCHARGE') !== -1) {
                                if(v20>0) result.ows[0]=v20;
                                if(v40>0) result.ows[1]=v40;
                                if(v40hc>0) result.ows[2]=v40hc;
                            }
                        }
                    }
                    return result;
                    """
                    data = driver.execute_script(js_price)

                    total_20   = data['ocean'][0] + data['surcharge'][0] - data['ows'][0]
                    total_40   = data['ocean'][1] + data['surcharge'][1] - data['ows'][1]
                    total_40hc = data['ocean'][2] + data['surcharge'][2] - data['ows'][2]
                    othc_included = data['othc_included']
                    manifest_fee_found = data['manifest_fee_found']
                    ows_found  = any(data['ows'][j] > 0 for j in range(3))

                    print(f"   => GIÁ CHỐT: 20'={total_20} | 40'={total_40} | 40HC={total_40hc}"
                          f"{' | OTHC included' if othc_included else ''}"
                          f"{' | Manifest/ENS detected' if manifest_fee_found else ''}"
                          f"{' | OWS detected' if ows_found else ''}")

                except Exception as e:
                    print(f"      ⚠️ Lỗi bóc bảng giá: {e}")

            # ── FREE TIME: Chỉ lấy từ card đầu tiên ────────────────
            if i == 0:
                try:
                    print("   -> Đang tìm và mở tab Free Time (D&D)...")
                    tab_dd = None
                    
                    # Ưu tiên tìm tab D&D bằng text
                    dd_tabs = card_el.find_elements(By.XPATH, ".//*[contains(@class,'el-tabs__item') or @role='tab']")
                    for t in dd_tabs:
                        txt = driver.execute_script("return arguments[0].innerText;", t)
                        if txt and any(k in txt.upper() for k in ["D&D", "DND", "DETENTION", "DEMURRAGE", "FREE TIME"]):
                            tab_dd = t
                            break

                    if tab_dd:
                        # Cuộn mượt và click ép bằng JS
                        driver.execute_script(
                            "arguments[0].scrollIntoView({behavior:'instant',block:'center'});"
                            "window.scrollBy(0,-100);", tab_dd)
                        time.sleep(0.3)
                        driver.execute_script("arguments[0].click();", tab_dd)

                        # Chờ bảng dữ liệu hiển thị
                        try:
                            active_dd_pane = WebDriverWait(driver, 4, poll_frequency=0.1).until(
                                lambda d: next((
                                    p for p in card_el.find_elements(By.XPATH, ".//div[contains(@class,'el-tab-pane')]")
                                    if p.is_displayed() and p.find_elements(By.XPATH, ".//table")
                                ), None)
                            )
                        except Exception:
                            print("      [WARN] Click được tab Free Time nhưng không tìm thấy thẻ <table>.")
                            active_dd_pane = None

                        dem, det, merged = "", "", ""
                        debug_headers = []
                        debug_rows = []

                        if active_dd_pane:
                            # TÌM BẢNG IMPORT FREE TIME (Bỏ qua Export)
                            tables = active_dd_pane.find_elements(By.XPATH, ".//table")
                            target_table = None
                            
                            if len(tables) > 1:
                                for tbl in tables:
                                    # Quét text của thẻ cha bọc table để xem có chữ IMPORT không
                                    parent_text = driver.execute_script("return arguments[0].parentElement.innerText;", tbl).upper()
                                    if "IMPORT" in parent_text or "DESTINATION" in parent_text:
                                        target_table = tbl
                                        break
                                # Nếu không thấy chữ Import, mặc định lấy bảng cuối cùng (bên phải)
                                if not target_table:
                                    target_table = tables[-1] 
                            elif tables:
                                target_table = tables[0]

                            if target_table:
                                # THUẬT TOÁN ĐỌC HEADER TRÊN BẢNG IMPORT ĐÃ CHỌN
                                headers = target_table.find_elements(By.XPATH, ".//thead/tr/th")
                                charge_idx, duration_idx = 0, 2 # Mặc định
                                
                                for idx, th in enumerate(headers):
                                    th_text = driver.execute_script("return arguments[0].innerText;", th).upper().strip()
                                    debug_headers.append(th_text)
                                    
                                    if th_text in ["CHARGE", "CHARGE TYPE", "FEE", "DESCRIPTION", "TYPE"]: 
                                        charge_idx = idx
                                    elif th_text in ["DURATION", "PERIOD", "FREE TIME", "LIMIT", "DAYS"]: 
                                        duration_idx = idx

                                # Quét các dòng dữ liệu trong tbody của ĐÚNG BẢNG ĐÓ
                                rows_ft = target_table.find_elements(By.XPATH, ".//tbody/tr")
                                
                                for r in rows_ft:
                                    cols = r.find_elements(By.XPATH, ".//*[self::td or self::th]")
                                    row_texts = [driver.execute_script("return arguments[0].innerText;", c).strip() for c in cols]
                                    if row_texts and any(row_texts):
                                        debug_rows.append(" | ".join(row_texts))
                                    
                                    if len(cols) > max(charge_idx, duration_idx):
                                        charge_type = driver.execute_script("return arguments[0].innerText;", cols[charge_idx]).upper()
                                        duration_val = driver.execute_script("return arguments[0].innerText;", cols[duration_idx])

                                        match = re.search(r"(\d+)", duration_val)
                                        if match:
                                            val = match.group(1)
                                            if any(k in charge_type for k in ["MERGED", "COMBINED", "D&D", "FREE TIME", "D & D"]) and not merged:
                                                merged = val
                                            elif "DEMURRAGE" in charge_type and not dem:
                                                dem = val
                                            elif "DETENTION" in charge_type and not det:
                                                det = val

                        # Tổng hợp và xuất chuỗi format chuẩn chỉnh
                        if merged:
                            free_time_pod = f"{merged} COMBINED"
                        elif dem and det:
                            free_time_pod = f"{dem} DEM + {det} DET"
                        elif dem:
                            free_time_pod = f"{dem} DEM"
                        elif det:
                            free_time_pod = f"{det} DET"
                        else:
                            print("      [WARN] Có bảng nhưng không bóc được số liệu.")
                            print(f"      [DEBUG] Headers: {debug_headers}")
                            print(f"      [DEBUG] Rows: {debug_rows}")
                    else:
                        print("      [WARN] Tuyến này hãng không hiển thị tab Free Time/D&D.")

                except Exception as e:
                    print(f"      [WARN] Lỗi bóc Free Time: {type(e).__name__} - {e}")

            # ── VESSEL & VOYAGE: Tab Route ─────────────────────────
            try:
                tab_route = card_el.find_element(By.XPATH,
                    ".//div[contains(@class,'el-tabs__item') and contains(.,'Route')]")
                driver.execute_script(
                    "arguments[0].scrollIntoView({behavior:'smooth',block:'center'});"
                    "window.scrollBy(0,-150);", tab_route)
                time.sleep(0.1)
                driver.execute_script("arguments[0].click();", tab_route)

                active_pane_xpath = (".//div[contains(@class,'el-tab-pane') "
                                     "and not(contains(@style,'display: none'))]")
                WebDriverWait(driver, 3, poll_frequency=0.05).until(
                    lambda d: card_el.find_elements(By.XPATH,
                        f"{active_pane_xpath}//ul[contains(@class,'list-main') "
                        f"or contains(@class,'route') or contains(@class,'vessel')]")
                )
                time.sleep(0.5)  # ← Tăng từ 0.1 lên 0.5s để React render xong

                # ✅ LẤY FRESH CARD SAU KHI CLICK TAB ROUTE
                fresh_cards = driver.find_elements(By.CSS_SELECTOR, "article.card-route-horizontal")
                if idx < len(fresh_cards):
                    card_el = fresh_cards[idx]

                # ✅ LẤY TÀU CHÍNH (LEG 1) - DÙNG INDEX TỬ LOOP
                v_name = "TBA"
                v_voyage = "N/A"
                try:
                    # Dùng thẳng idx từ loop (đã chính xác)
                    card_idx = idx
                    
                    # JS lấy vessel từ card thứ card_idx
                    js_get_vessel = """
                    var cardIdx = arguments[0];
                    var cards = document.querySelectorAll('article.card-route-horizontal');
                    if(cardIdx >= cards.length) return null;
                    
                    var card = cards[cardIdx];
                    var wrap = card.querySelector('.wrap-routes');
                    if(!wrap) return null;
                    
                    var lis = wrap.querySelectorAll('ul.route > li');
                    var firstVessel = null;
                    
                    // Tìm li.more-infos.vessel đầu tiên
                    for(var i=0; i<lis.length; i++) {
                        if(lis[i].className.includes('more-infos') && lis[i].className.includes('vessel')) {
                            firstVessel = lis[i];
                            break;
                        }
                    }
                    
                    if(!firstVessel) return null;
                    
                    var ulLis = firstVessel.querySelectorAll('ul > li');
                    var result = {service: '', vessel: ''};
                    
                    // Quét tất cả li, dùng dt text để xác định đúng field
                    for(var k=0; k<ulLis.length; k++) {
                        var dts = ulLis[k].querySelectorAll('dt');
                        var dds = ulLis[k].querySelectorAll('dd');
                        for(var m=0; m<dts.length; m++) {
                            var dtText = (dts[m].innerText || dts[m].textContent || '').trim().toUpperCase();
                            if(dtText === 'VESSEL' && m < dds.length && !result.vessel) {
                                result.vessel = (dds[m].innerText || dds[m].textContent || '').trim();
                            } else if(dtText === 'SERVICE' && m < dds.length && !result.service) {
                                result.service = (dds[m].innerText || dds[m].textContent || '').trim();
                            }
                        }
                    }
                    
                    return result;
                    """
                    
                    data = driver.execute_script(js_get_vessel, card_idx)
                    if data and data.get('vessel'):
                        v_name = data['vessel']
                        v_voyage = data.get('service', 'N/A')
                        print(f"      ✅ Tàu chính: {v_name} ({v_voyage})")
                    else:
                        print(f"      ⚠️ Không tìm thấy vessel qua JS (card_idx={card_idx})")
                        
                except Exception as e:
                    print(f"      ⚠️ Lỗi bóc vessel JS: {type(e).__name__}")

                vessel_entries.append(
                    f"{v_name} ({v_voyage}) / ETD: {opt['date'].day}-{opt['date'].strftime('%b')}"
                    f" / Transit time: {opt['transit']} Days / Transshipment: {opt['ts_port']}")

            except Exception as e:
                print(f"      ⚠️ Không thể bóc Vessel: {e}")
                vessel_entries.append(
                    f"TBA / ETD: {opt['date'].day}-{opt['date'].strftime('%b')}"
                    f" / Transit time: {opt['transit']} Days / Transshipment: {opt['ts_port']}")

            ts_entries.append(opt['ts_port'])

        # ══════════════════════════════════════════════════════════════════════════════
        # BƯỚC 3: GHI VÀO EXCEL
        # ══════════════════════════════════════════════════════════════════════════════

        # Cột F/G/H: Giá 20GP / 40GP / 40HQ
        ws.cell(row=row_index, column=6).value = total_20   if total_20   else "Check"
        ws.cell(row=row_index, column=7).value = total_40   if total_40   else "Check"
        ws.cell(row=row_index, column=8).value = total_40hc if total_40hc else "Check"

        # Cột I: ETD
        def fmt_date(d):
            return f"{d.day}-{d.strftime('%b')}"

        final_dates = [o['date'] for o in valid_opts]
        if len(final_dates) == 1:
            final_etd_str = fmt_date(final_dates[0])
        elif len(final_dates) == 2:
            final_etd_str = f"{fmt_date(final_dates[0])} & {fmt_date(final_dates[1])}"
        elif len(final_dates) >= 3:
            if all(d.month == final_dates[0].month for d in final_dates):
                final_etd_str = (f"{final_dates[0].day}, {final_dates[1].day},"
                                 f" {fmt_date(final_dates[2])}")
            else:
                final_etd_str = " & ".join(fmt_date(d) for d in final_dates)
        else:
            final_etd_str = "N/A"
        ws.cell(row=row_index, column=9).value = final_etd_str

        # Cột J: Transit
        transits = [o['transit'] for o in valid_opts]
        ws.cell(row=row_index, column=10).value = (str(min(transits))
            if min(transits) == max(transits) else f"{min(transits)}-{max(transits)}")

        # Cột K: Valid
        ws.cell(row=row_index, column=11).value = calculate_cma_validity(valid_opts[-1]['date'])

        # Cột M: Remark
        # Lấy POD từ file Excel (cột D) để định vị khu vực
        pod_text = str(ws.cell(row=row_index, column=4).value or "").strip().upper()
        manifest_acronym = "ENS" # Mặc định là ENS cho Châu Âu và các khu vực khác
        
        # Nếu POD thuộc Trung Quốc -> Đổi thành AMS
        if any(k in pod_text for k in ["SHANGHAI", "NINGBO", "QINGDAO", "XINGANG", "TIANJIN", "YANTIAN", "SHEKOU", "NANSHA", "XIAMEN", "DALIAN", "CHINA"]):
            manifest_acronym = "AMS"
        # Nếu POD thuộc Nhật Bản -> Đổi thành AFS
        elif any(k in pod_text for k in ["TOKYO", "YOKOHAMA", "NAGOYA", "OSAKA", "KOBE", "MOJI", "HAKATA", "JAPAN"]):
            manifest_acronym = "AFS"

        # Cột M: Remark
        remark = "INCLUDED THC, SUBJECT TO BILL, SEAL, TELEX" if othc_included \
                 else "SUBJECT TO THC, BILL, SEAL, TLX"
        
        if manifest_fee_found:
            remark += f", {manifest_acronym}"
        if ows_found:
            remark += ", OWS"
            
        ws.cell(row=row_index, column=13).value = remark

        # Cột N: Free Time
        ws.cell(row=row_index, column=14).value = free_time_pod

        # Cột O: Vessel (xuống dòng)
        ws.cell(row=row_index, column=15).value = "\n".join(vessel_entries)
        ws.cell(row=row_index, column=15).alignment = openpyxl.styles.Alignment(wrapText=True)

        # Cột P: Transshipment (xuống dòng)
        unique_ts = []
        for ts in ts_entries:
            if ts not in unique_ts: unique_ts.append(ts)
        ws.cell(row=row_index, column=16).value = " or\n".join(unique_ts)
        ws.cell(row=row_index, column=16).alignment = openpyxl.styles.Alignment(wrapText=True)

        try:
            wb.save(excel_path)
        except PermissionError:
            print("      ❌ LỖI GHI FILE: TẮT FILE EXCEL ĐI SẾP ƠI!!!")

        print(f"   [OK] Đã lưu dữ liệu dòng {row_index} thành công!")
        return "SUCCESS"

    except Exception as e:
        print(f"      ❌ Lỗi Scrape chi tiết tổng quát: {e}")
        ws.cell(row=row_index, column=6).value = "Error"
        try: wb.save(excel_path)
        except: pass
        return "ERROR"

# ===================================================================================
# --- 4. MAIN LOOP ---
# ===================================================================================
# Danh sách nhận diện các cảng/quốc gia thuộc Đông Á và Đông Nam Á
INTRA_ASIA_KEYWORDS = [
    "SHANGHAI", "NINGBO", "QINGDAO", "XINGANG", "TIANJIN", "YANTIAN", "SHEKOU", "NANSHA", "XIAMEN", "DALIAN",
    "HONG KONG", "BUSAN", "INCHEON", "KWANGYANG", "TOKYO", "YOKOHAMA", "NAGOYA", "OSAKA", "KOBE", "MOJI", "HAKATA",
    "KEELUNG", "KAOHSIUNG", "TAICHUNG", "MANILA", "CEBU", "SUBIC", "BATANGAS", "DAVAO",
    "PORT KELANG", "PORT KLANG", "PENANG", "PASIR GUDANG", "TANJUNG PELEPAS", "KUANTAN",
    "JAKARTA", "SURABAYA", "SEMARANG", "BELAWAN", "PANJANG",
    "BANGKOK", "LAEM CHABANG", "LAT KRABANG", "SONGKHLA", "YANGON", "SIHANOUKVILLE", "PHNOM PENH",
    # FIX: bổ sung Myanmar và các alias
    "MYANMAR", "RANGOON", "THILAWA",
    "CHINA", "JAPAN", "KOREA", "TAIWAN", "MALAYSIA", "THAILAND", "INDONESIA", "PHILIPPINES", "SINGAPORE",
    "MYANMAR", "VIETNAM", "CAMBODIA",  # quốc gia gần VN
]

def extract_row_data(ws, r_idx):
    return {c: ws.cell(row=r_idx, column=c).value for c in range(6, 17)}

def write_row_data(ws, r_idx, data):
    for c, val in data.items():
        ws.cell(row=r_idx, column=c).value = val

def execute_single_search(row_index, pol_val, pod, is_riyadh, is_intra_hcm, pol_excel):
    global is_first_run_in_session, previous_pol
    wait = WebDriverWait(driver, 15)

    def do_search_steps():
        if is_riyadh: handle_pod_selection_popup(prefer_port="JEDDAH")
        # Fix: RAMP luôn cần Vũng Tàu, không phân biệt Intra hay không
        if "RAMP" in pol_val or (not is_intra_hcm and "VNSGN" in pol_val):
            select_vung_tau_mandatory()

        # === CHECK BANNER NO ROUTE ===
        time.sleep(0.5) # Nghỉ nửa nhịp cho web render banner
        try:
            alerts = driver.find_elements(By.CSS_SELECTOR, "span.el-alert__title")
            for alert in alerts:
                if alert.is_displayed() and "SpotOn hasn't found possible route" in alert.text:
                    print("      ⚠️ Phát hiện banner: No Route. Bỏ qua luôn!")
                    ws.cell(row=row_index, column=6).value = "No Route"
                    try: wb.save(excel_path)
                    except: pass
                    return "NO_ROUTE"
        except:
            pass
        # =============================

        pick_date_plus_7()
        if is_first_run_in_session:
            select_containers()
        else:
            ensure_containers_filled()

        print("   -> Ép React khóa dữ liệu và Recheck form...")
        driver.execute_script("document.body.click();")
        time.sleep(0.3) 
        js_check_form = """
        var txt = "";
        var els = document.querySelectorAll('.selected-value-wrapper, input');
        for(var i=0; i<els.length; i++) txt += (els[i].innerText || els[i].value || "").toUpperCase() + " ";
        return txt;
        """
        form_text = driver.execute_script(js_check_form)
        pod_check = pod.split(",")[0].strip().upper()
        
        if pod_check not in form_text:
            print(f"      ⚠️ Phát hiện form chưa nhận POD [{pod_check}]! Đang bơm lại...")
            smart_update_field('POD', pod)
            if "RAMP" in pol_val or (not is_intra_hcm and "VNSGN" in pol_val): select_vung_tau_mandatory()
            time.sleep(0.5)
            driver.execute_script("document.body.click();")
            time.sleep(0.3)
        print("   -> Kiểm tra Commodity trước khi Get Quote...")
        js_check_commodity = """
        var inputs = document.querySelectorAll('input');
        for(var i=0; i<inputs.length; i++) {
            var ph = (inputs[i].placeholder || '').toLowerCase();
            var val = (inputs[i].value || '').trim();
            if(ph.includes('commodity')) {
                return val;  // Trả về giá trị hiện tại của ô commodity
            }
        }
        // Kiểm tra thêm qua text của selected label
        var labels = document.querySelectorAll('.el-select .el-input__inner');
        for(var j=0; j<labels.length; j++) {
            var lval = (labels[j].value || '').trim();
            if(lval && lval.toLowerCase() !== 'choose a commodity' && lval.length > 2) {
                return lval;
            }
        }
        return '';
        """
        commodity_val = driver.execute_script(js_check_commodity)
        if not commodity_val:
            print("      ⚠️ Commodity TRỐNG! Đang chọn lại trước khi Get Quote...")
            select_commodity_refresh()
            time.sleep(0.5)
        else:
            print(f"      -> [OK] Commodity đã có: '{commodity_val}'")
        print("7. Get Quote...")
        btn = driver.find_element(By.XPATH, "//button[contains(., 'Get My Quote') or @id='SearchQuote']")
        driver.execute_script("arguments[0].click();", btn)

        print("   -> Đang chờ web load kết quả...")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "article.card-route-horizontal")))
            time.sleep(random.uniform(1.0, 1.5))
        except: sleep_human(3.0, 5.0)

        page_text = driver.execute_script("return document.body.innerText;").upper()
        if pod_check not in page_text[:3000]:
            print(f"   ⚠️ BÁO ĐỘNG LỆCH ROUTE! Không thấy {pod_check}. Đang sửa lại...")
            if perform_modify_search_action():
                smart_update_field('POD', pod)
                if is_riyadh: handle_pod_selection_popup("JEDDAH")
                if not is_intra_hcm and "VNSGN" in pol_val: select_vung_tau_mandatory()
                select_commodity_refresh()
                btn = driver.find_element(By.XPATH, "//button[contains(., 'Get My Quote') or @id='SearchQuote']")
                driver.execute_script("arguments[0].click();", btn)
                try:
                    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "article.card-route-horizontal")))
                    time.sleep(random.uniform(1.0, 1.5))
                except: sleep_human(3.0, 5.0)
            else:
                print("   ❌ Lỗi sửa Route, bỏ qua.")
                return False

        return scrape_multi_etd_and_save(row_index, ws, pol_excel)

    if is_first_run_in_session:
        print(">> FULL MODE (Lần đầu)...")
        if "SpotOn" not in driver.title:
            driver.get("https://www.cma-cgm.com/ebusiness/pricing/instant-quoting")
            time.sleep(2)
        check_and_kill_popup()
        inputs = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Name / Code / Port']")))
        select_port_full(inputs[0], pol_val)
        sleep_human(0.5, 0.8)
        inputs = driver.find_elements(By.XPATH, "//input[@placeholder='Name / Code / Port']")
        select_port_full(inputs[-1], pod)
        
        status = do_search_steps()
        is_first_run_in_session = False
        previous_pol = pol_val
        if status == "NO_ROUTE": return False
    else:
        print(">> MODIFY MODE...")
        if perform_modify_search_action():
            if pol_val != previous_pol: smart_update_field('POL', pol_val)
            smart_update_field('POD', pod)
            
            status = do_search_steps()
            if status == "NO_ROUTE":
                previous_pol = pol_val
                return False
                
            if status == "SOLD_OUT":
                print("   -> ⚠️ Bị Sold Out ở Modify Mode, tiến hành recheck lại bằng FULL MODE...")
                driver.get("https://www.cma-cgm.com/ebusiness/pricing/instant-quoting")
                time.sleep(2)
                check_and_kill_popup()
                inputs = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Name / Code / Port']")))
                select_port_full(inputs[0], pol_val)
                sleep_human(0.5, 0.8)
                inputs = driver.find_elements(By.XPATH, "//input[@placeholder='Name / Code / Port']")
                select_port_full(inputs[-1], pod)
                
                retry_status = do_search_steps()
                if retry_status == "NO_ROUTE":
                    previous_pol = pol_val
                    return False
            previous_pol = pol_val
        else:
            print("   -> Lỗi Modify! Chuyển FULL MODE...")
            driver.get("https://www.cma-cgm.com/ebusiness/pricing/instant-quoting")
            time.sleep(2)
            check_and_kill_popup()
            inputs = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='Name / Code / Port']")))
            select_port_full(inputs[0], pol_val)
            sleep_human(0.5, 0.8)
            inputs = driver.find_elements(By.XPATH, "//input[@placeholder='Name / Code / Port']")
            select_port_full(inputs[-1], pod)
            
            status = do_search_steps()
            previous_pol = pol_val
            if status == "NO_ROUTE": return False

# ✅ LẤY ROW TỪ ARGUMENT TERMINAL
import sys
target_row = None
if len(sys.argv) > 1:
    try:
        target_row = int(sys.argv[1])
        print(f"[INFO] Chạy riêng row {target_row}...")
    except ValueError:
        print(f"[WARN] Argument không hợp lệ: {sys.argv[1]}. Chạy tất cả row.")

try:
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    is_first_run_in_session = True
    previous_pol = ""

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # ✅ FILTER THEO TARGET ROW
        if target_row and i != target_row:
            continue
        pol_excel = str(row[2] or "").strip()
        pod       = str(row[3] or "").strip()
        carrier   = str(row[4] or "").strip().upper()

        if not pol_excel or not pod: continue
        if carrier not in {"CMA", "ANL", "CNC", "APL"}: continue
        if FILTER_POL and pol_excel.upper() != FILTER_POL: continue
        if FILTER_POD and pod.upper() != FILTER_POD: continue

        # Phân loại logic Intra-Asia
        is_hcm = "HO CHI MINH" in pol_excel.upper() or "VNSGN" in pol_excel.upper()
        is_intra_asia = any(k in pod.upper() for k in INTRA_ASIA_KEYWORDS)
        is_riyadh = 'RIYADH' in pod.upper()

        print(f"\n==========================================")
        print(f"--- DÒNG {i}: {pol_excel} -> {pod} | Hãng: {carrier} ---")
        if is_hcm and is_intra_asia:
            print(f"   [!] TUYẾN INTRA-ASIA TỪ HCM: Cần check cả PORT và RAMP để so giá.")

        if "Access Denied" in driver.title:
            print("!!! BỊ CHẶN !!!")
            break

        if is_hcm and is_intra_asia:
            # CHECK LẦN 1: PORT
            print("\n   >>> CHECK OPTION 1: VNSGN - PORT <<<")
            execute_single_search(i, "VNSGN-PORT", pod, is_riyadh, True, pol_excel)
            data_port = extract_row_data(ws, i)
            
            # Hàm phụ parse giá để tránh so sánh lỗi chuỗi "Check", "No Quote"
            def parse_price(val):
                try: return float(val)
                except: return float('inf')
                
            price_port = parse_price(data_port.get(6))

            # CHECK LẦN 2: RAMP
            print("\n   >>> CHECK OPTION 2: VNSGN - RAMP <<<")
            execute_single_search(i, "VNSGN-RAMP", pod, is_riyadh, True, pol_excel)
            data_ramp = extract_row_data(ws, i)
            price_ramp = parse_price(data_ramp.get(6))

            # SO SÁNH GIÁ VÀ GHI LẠI
            print(f"\n   >>> KẾT QUẢ SO SÁNH DÒNG {i}: PORT ({data_port.get(6)}) vs RAMP ({data_ramp.get(6)})")
            if price_port <= price_ramp and price_port != float('inf'):
                print("   => PORT RẺ HƠN HOẶC BẰNG! Ghi đè lại dữ liệu PORT vào Excel.")
                write_row_data(ws, i, data_port)
            elif price_ramp < price_port:
                print("   => RAMP RẺ HƠN! Đã giữ lại dữ liệu RAMP trên Excel.")
            else:
                print("   => CẢ 2 ĐỀU LỖI HOẶC SOLD OUT. Trả về báo cáo PORT.")
                write_row_data(ws, i, data_port)
            
            try: wb.save(excel_path)
            except: pass
        else:
            # Các cảng như Hải Phòng hoặc các tuyến xa (Mỹ, Úc, Âu...) chạy 1 lần như bình thường
            pol = "VNSGN" if is_hcm else ("VNHPH" if "HAI PHONG" in pol_excel.upper() else pol_excel)
            execute_single_search(i, pol, pod, is_riyadh, False, pol_excel)

        print("   -> Nghỉ 2-3s để chuyển dòng...")
        sleep_human(2, 3)

except Exception as e:
    print(f"Lỗi chung vòng lặp: {e}")
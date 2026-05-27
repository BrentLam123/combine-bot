"""
EMC Price Checker — TEST MODE (Selenium, Edge port 9521)
Chạy thử 3 cặp cảng mẫu, in kết quả ra terminal.

Cài đặt:
    pip install selenium openpyxl

Chạy:
    python emc_test.py
"""

import math
import re
import time
import os
from datetime import datetime, timedelta
from html.parser import HTMLParser

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl


current_folder = os.getcwd()
driver_path    = os.path.join(current_folder, "msedgedriver.exe")
excel_path     = os.environ.get("EXCEL_PATH", os.path.join(current_folder, "input_gia.xlsx"))
FILTER_POL     = os.environ.get("FILTER_POL", "").strip().upper()
FILTER_POD     = os.environ.get("FILTER_POD", "").strip().upper()

import subprocess
import socket

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

if not is_port_in_use(9521):
    print("[HỆ THỐNG] Edge EMC chưa mở. Đang tự động khởi động...")
    try:
        subprocess.Popen([
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            "--remote-debugging-port=9521",
            r"--user-data-dir=C:\edge_emc",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-backgrounding-occluded-windows"
        ])
        time.sleep(3)
    except: pass
else:
    print("[HỆ THỐNG] Edge EMC đã mở sẵn. Bỏ qua lệnh khởi động trình duyệt.")

edge_options = Options()
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9521")
service = Service(executable_path=driver_path)
driver  = webdriver.Edge(service=service, options=edge_options)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
})

BOOKING_FEE = 10
BASE_URL     = "https://portal.greenxtrade.com/quotes"
EMC_GROUP    = {"EMC", "EVERGREEN"}

# ===================================================================================
# --- CHỌN CẢNG ĐÚNG TỪ DROPDOWN ---
# ===================================================================================
def pick_correct_port(options_text, query):
    query_upper = query.upper().strip()

    # Ưu tiên (All Ports)
    for opt in options_text:
        if "(all ports)" in opt.lower():
            return opt

    # Quy tắc dấu phẩy ngay sau tên
    for opt in options_text:
        if opt.upper().strip().startswith(query_upper + ","):
            return opt

    # Fallback: chứa query
    for opt in options_text:
        if query_upper in opt.upper():
            return opt

    return None


def select_port_emc(xpath_input, port_name):
    print(f"      + Nạp cảng: {port_name}")
    for attempt in range(3):
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_input)))
            inp = next((e for e in driver.find_elements(By.XPATH, xpath_input) if e.is_displayed()), None)
            if not inp:
                raise Exception("Không tìm thấy input!")

            # Check giữ nguyên
            current_val = driver.execute_script("return arguments[0].value;", inp)
            if current_val and port_name.upper() in current_val.upper():
                print(f"        -> Đã có sẵn: {current_val} -> GIỮ NGUYÊN!")
                return

            # Vue hack: set value + dispatch input event
            driver.execute_script("""
                var inp = arguments[0];
                var val = arguments[1];
                inp.focus();
                var setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                setter.call(inp, val);
                inp.dispatchEvent(new Event('input', { bubbles: true }));
            """, inp, port_name)

            # Chờ ul.open xuất hiện
            WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "ul.open li[data-arrow-control]"))
            )
            time.sleep(0.2)

            opt_spans = driver.find_elements(By.CSS_SELECTOR, "ul.open li[data-arrow-control] span.item-text")
            opts_text = [s.text.strip() for s in opt_spans if s.text.strip()]
            opts      = driver.find_elements(By.CSS_SELECTOR, "ul.open li[data-arrow-control]")
            print(f"        Dropdown: {opts_text[:5]}")

            chosen = pick_correct_port(opts_text, port_name)
            if not chosen:
                raise Exception(f"Không khớp option nào cho '{port_name}'")

            for li in opts:
                try:
                    if li.find_element(By.CSS_SELECTOR, "span.item-text").text.strip() == chosen:
                        li.click()
                        print(f"        -> Đã chốt: {chosen}")
                        time.sleep(0.3)
                        return
                except: continue

        except Exception as e:
            print(f"        ⚠️ Lần {attempt+1} thất bại: {e}")
            time.sleep(0.5)

    raise Exception(f"Thất bại 3 lần: {port_name}")


# ===================================================================================
# --- 9 QUY TẮC VÀNG (copy từ bot COSCO) ---
# ===================================================================================
def apply_9_golden_rules(danh_sach_chuyen):
    danh_sach_chuyen.sort(key=lambda x: (x["etd_dt"], x["tt_days"]))

    # Lọc trùng ETD (giữ cái TT ngắn nhất)
    seen, list_loc_trung = set(), []
    for c in danh_sach_chuyen:
        if c["etd_dt"] not in seen:
            list_loc_trung.append(c)
            seen.add(c["etd_dt"])

    etd_dat_chuan = []
    if list_loc_trung:
        ngan_nhat = min(c["tt_days"] for c in list_loc_trung)
        first_date = list_loc_trung[0]["etd_dt"]
        for c in list_loc_trung:
            if len(etd_dat_chuan) >= 3: break
            if len(etd_dat_chuan) > 0 and (c["etd_dt"] - etd_dat_chuan[-1]["etd_dt"]).days < 2: continue
            if (c["etd_dt"] - first_date).days <= 9 and c["tt_days"] <= ngan_nhat + 10:
                etd_dat_chuan.append(c)

    # Format ETD string
    num = len(etd_dat_chuan)
    if num == 0:   str_etd = "N/A"
    elif num == 1: str_etd = etd_dat_chuan[0]["etd_dt"].strftime("%d-%b")
    elif num == 2: str_etd = f"{etd_dat_chuan[0]['etd_dt'].strftime('%d-%b')} & {etd_dat_chuan[1]['etd_dt'].strftime('%d-%b')}"
    else:
        d1 = etd_dat_chuan[0]["etd_dt"].strftime("%d")
        d2 = etd_dat_chuan[1]["etd_dt"].strftime("%d")
        d3 = etd_dat_chuan[2]["etd_dt"].strftime("%d-%b")
        str_etd = f"{d1}, {d2}, {d3}"

    all_tt = [c["tt_days"] for c in etd_dat_chuan]
    str_tt = f"{min(all_tt)}" if min(all_tt) == max(all_tt) else f"{min(all_tt)}-{max(all_tt)}"

    return etd_dat_chuan, str_etd, str_tt


# ===================================================================================
# --- PARSE BẢNG GIÁ TỪ PRICE DETAILS ---
# ===================================================================================
BLOCKLIST = [
    'THC', 'TERMINAL HANDLING', 'DESTINATION', 'D/O', 'DELIVERY ORDER',
    'CFS', 'PORT CONGESTION', 'PSS', 'EBS', 'ERC', 'STF',
]

def parse_usd(text):
    text = re.sub(r'\(.*?\)', '', text)
    match = re.search(r'\$?([\d,]+\.?\d*)', text.replace('USD', '').strip())
    if match:
        return float(match.group(1).replace(",", ""))
    return 0.0

def should_skip(charge_name):
    name_upper = charge_name.upper()
    return any(b in name_upper for b in BLOCKLIST)

def parse_price_details(subrow_el):
    charges = {"20": 0.0, "40": 0.0, "40hq": 0.0}
    remark = ""

    CONT_KEY = {
        "20' standard dry": "20",
        "40' standard dry": "40",
        "40' high cube":    "40hq",
    }

    SKIP_KEYWORDS = ["terminal handling", "thc", "destination", "d/o", "delivery order"]

    try:
        origin_section = subrow_el.find_elements(
            By.CSS_SELECTOR,
            ".quotes-search-results-subrow-price-details-content"
        )
        if not origin_section:
            print("      ⚠️ Không tìm thấy section giá")
            return {"20": None, "40": None, "40hq": None}, ""

        origin_table = origin_section[0]
        rows = origin_table.find_elements(By.CSS_SELECTOR, "table tbody tr")
        booking_fee_added = False

        for row in rows:
            tds = row.find_elements(By.CSS_SELECTOR, "td")
            if len(tds) < 4:
                continue

            charge_name = tds[0].find_element(By.CSS_SELECTOR, "span").text.strip() if tds[0].find_elements(By.CSS_SELECTOR, "span") else ""
            cont_raw    = tds[1].find_element(By.CSS_SELECTOR, "span").text.strip().lower() if tds[1].find_elements(By.CSS_SELECTOR, "span") else ""
            price_spans = tds[3].find_elements(By.CSS_SELECTOR, "span")
            price_text  = price_spans[0].text.strip() if price_spans else ""

            if not charge_name:
                continue

            if "booking fee" in charge_name.lower() and not booking_fee_added:
                fee = parse_usd(price_text) or BOOKING_FEE
                for k in charges:
                    charges[k] += fee
                booking_fee_added = True
                print(f"      [+] Booking Fee: ${fee}")
                continue

            if any(kw in charge_name.lower() for kw in SKIP_KEYWORDS):
                print(f"      [-] Bỏ qua: {charge_name}")
                continue

            cont_key = CONT_KEY.get(cont_raw)
            if cont_key:
                amt = parse_usd(price_text)
                charges[cont_key] += amt
                print(f"      [+] {charge_name} ({cont_raw}): ${amt}")

        if not booking_fee_added:
            for k in charges:
                charges[k] += BOOKING_FEE
            print(f"      [+] Booking Fee (default): ${BOOKING_FEE}")

        full_text = subrow_el.text.upper()
        if "NON APPLICABLE" in full_text and "HWCS" in full_text:
            remark = "SUBJECT TO THC, BILL, SEAL, TLX, OWS"

    except Exception as e:
        print(f"      ⚠️ Lỗi parse Price Details: {e}")

    totals = {k: math.ceil(v) if v > 0 else None for k, v in charges.items()}
    return totals, remark


# ===================================================================================
# --- SEARCH 1 CẶP CẢNG ---
# ===================================================================================
def search_one(port_from, port_to):
    print(f"\n{'='*55}")
    print(f"  🔍 {port_from.upper()} → {port_to.upper()}")
    print(f"{'='*55}")

    base = {"from": port_from, "to": port_to,
            "etd": "N/A", "tt": "N/A",
            "price_20": None, "price_40": None, "price_40hq": None,
            "remark": "", "error": None}
    try:
        # ── Nhập FROM ──
        print(f"   📍 Nhập FROM: {port_from}")
        select_port_emc('//input[@aria-label="input for From"]', port_from)

        # ── Nhập TO ──
        print(f"   📍 Nhập TO: {port_to}")
        select_port_emc('//input[@aria-label="input for To"]', port_to)

        # ── Chọn số lượng cont = 1 ──
        print(f"   📦 Kiểm tra số lượng cont...")
        for label in ["20' GP", "40' GP", "40' HQ"]:
            try:
                inp_cont = driver.find_element(By.XPATH, f'//input[@aria-label="input for {label}"]')
                if inp_cont.get_attribute("value") == "1":
                    print(f"      -> {label} đã là 1, bỏ qua ✅")
                    continue
                # Click vào div.searchbar bao ngoài input
                searchbar = inp_cont.find_element(By.XPATH, "ancestor::div[contains(@class,'searchbar')]")
                searchbar.click()
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "ul.open li[data-arrow-control]")))
                time.sleep(0.15)
                for li in driver.find_elements(By.CSS_SELECTOR, "ul.open li[data-arrow-control]"):
                    if li.find_element(By.CSS_SELECTOR, "span.item-text").text.strip() == "1":
                        li.click()
                        print(f"      -> {label} = 1 ✅")
                        break
            except Exception as e:
                print(f"      ⚠️ {label}: {e}")

        # ── Chọn ngày ETD = hôm nay + 7 ──
        print(f"   📅 Kiểm tra ngày ETD...")
        etd_date = datetime.now() + timedelta(days=7)
        etd_target = etd_date.strftime("%m/%d/%Y")
        try:
            date_input = driver.find_element(By.CSS_SELECTOR, "input.mx-input")
            current_date = date_input.get_attribute("value")
            if current_date == etd_target:
                print(f"      -> Ngày đã đúng: {etd_target}, bỏ qua ✅")
            else:
                ActionChains(driver).move_to_element(date_input).click().perform()
                time.sleep(0.6)
                for fmt in [f"{etd_date.month}/{etd_date.day}/{etd_date.year}", etd_target]:
                    cells = driver.find_elements(By.XPATH, f'//td[@title="{fmt}"]')
                    if cells:
                        ActionChains(driver).move_to_element(cells[0]).click().perform()
                        print(f"      -> Đã chọn: {fmt} ✅")
                        time.sleep(0.5)
                        break
        except Exception as e:
            print(f"      ⚠️ Lỗi chọn ngày: {e}")

        # ── Bấm SEARCH ──
        search_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,
            "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'search')]")))
        driver.execute_script("arguments[0].click();", search_btn)
        print(f"   🖱️ Đã bấm SEARCH, chờ kết quả...")

        # ── Chờ cards load ──
        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "li.quotes-search-result")))
        time.sleep(2)

    except Exception as e:
        print(f"   ❌ Lỗi: {e}")
        return {**base, "error": str(e)}

    # ── Đọc cards, lọc chỉ lấy card có nút "Book" ──
    print(f"   📋 Đang đọc danh sách kết quả...")
    list_chuyen = []
    idx = 0
    while True:
        cards = driver.find_elements(By.CSS_SELECTOR, "li.quotes-search-result")
        if idx >= len(cards):
            break
        card = cards[idx]
        idx += 1
        try:
            btns = card.find_elements(By.CSS_SELECTOR, "button.book-button")
            if not btns or btns[0].text.strip().lower() != "book":
                print(f"      ⏭️ Bỏ qua card (không có Book)")
                continue

            etd_str = card.find_element(By.CSS_SELECTOR, ".estimated-dates-left.content span.date").text.strip()
            eta_str = card.find_element(By.CSS_SELECTOR, ".estimated-dates-right.content span.date").text.strip()
            tt_text = card.find_element(By.CSS_SELECTOR, ".estimated-dates-elapsed span").text.strip()
            tt_days = int(re.search(r'\d+', tt_text).group()) if re.search(r'\d+', tt_text) else 999

            etd_dt = datetime.strptime(etd_str, "%m/%d/%Y")
            print(f"      ✅ Card hợp lệ: ETD={etd_str} ETA={eta_str} TT={tt_days}d")
            list_chuyen.append({"element": card, "etd_dt": etd_dt, "tt_days": tt_days})

        except Exception as e:
            print(f"      ⚠️ Lỗi đọc card: {e}")
            continue

    if not list_chuyen:
        print(f"   ⚠️ Không có card nào có nút Book (có thể SOLD OUT / JOIN WAITLIST hết)")
        return {**base, "error": "NO SERVICE / SOLD OUT"}

    # ── Áp dụng 9 quy tắc vàng ──
    etd_chuan, str_etd, str_tt = apply_9_golden_rules(list_chuyen)
    print(f"\n   🏆 ETD chọn: {str_etd} | T/T: {str_tt} days")

    # Re-fetch cards để tránh stale element
    time.sleep(1)
    fresh_cards = driver.find_elements(By.CSS_SELECTOR, "li.quotes-search-result")
    etd_chuan_dt = etd_chuan[0]["etd_dt"]

    dai_dien_el = None
    for fc in fresh_cards:
        try:
            etd_str_fc = fc.find_element(By.CSS_SELECTOR, ".estimated-dates-left.content span.date").text.strip()
            if datetime.strptime(etd_str_fc, "%m/%d/%Y") == etd_chuan_dt:
                btns = fc.find_elements(By.CSS_SELECTOR, "button.book-button")
                if btns and btns[0].text.strip().lower() == "book":
                    dai_dien_el = fc
                    break
        except:
            continue

    if not dai_dien_el:
        return {**base, "error": "Không tìm lại được card đại diện sau re-fetch"}

    # ── Bấm Price Details trên card đại diện ──
    try:
        li_el = dai_dien_el

        price_details_tab = dai_dien_el.find_element(By.XPATH,
            ".//div[contains(@class,'quotes-search-results-list-item-tabs-tab')]"
            "[.//span[text()='Price Details']]")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", price_details_tab)
        time.sleep(0.3)
        ActionChains(driver).move_to_element(price_details_tab).click().perform()
        print(f"   🖱️ Đã bấm Price Details, chờ bảng giá...")

        subrow = li_el.find_element(By.CSS_SELECTOR, ".quotes-search-results-list-item-subrow")
        WebDriverWait(driver, 10).until(
            lambda d: subrow.find_elements(By.CSS_SELECTOR,
                ".quotes-search-results-subrow-price-details-content")
        )
        time.sleep(1)

    except Exception as e:
        return {**base, "error": f"Không bấm được Price Details: {e}"}

    # ── Parse bảng giá từ subrow ──
    try:
        print(f"\n   💰 Đang parse bảng giá...")
        totals, remark = parse_price_details(subrow)
    except Exception as e:
        return {**base, "error": f"Lỗi parse giá: {e}"}

    print(f"\n   💰 KẾT QUẢ SAU LỌC:")
    print(f"      ETD    : {str_etd}")
    print(f"      T/T    : {str_tt} days")
    print(f"      20' GP : {'$' + str(totals['20'])   if totals['20']   else 'N/A'}")
    print(f"      40' GP : {'$' + str(totals['40'])   if totals['40']   else 'N/A'}")
    print(f"      40' HQ : {'$' + str(totals['40hq']) if totals['40hq'] else 'N/A'}")
    if remark:
        print(f"      📝 Remark: {remark}")

    return {**base,
            "etd": str_etd, "tt": str_tt,
            "price_20":   totals.get("20"),
            "price_40":   totals.get("40"),
            "price_40hq": totals.get("40hq"),
            "remark":     remark}


# ===================================================================================
# --- MAIN ---
# ===================================================================================
print("""
╔══════════════════════════════════════════════╗
║        EMC Price Checker  🚢                 ║
╚══════════════════════════════════════════════╝
""")

wb = openpyxl.load_workbook(excel_path)
ws = wb.active

print(f"\n🌐 Bắt đầu chạy...")

first_emc_row = None

# FIX: Luôn navigate về trang chủ khi bắt đầu script → đảm bảo clean state cho cả lần chạy thứ 2
print("[HỆ THỐNG] Đang khởi tạo: navigate về https://portal.greenxtrade.com/ ...")
try:
    driver.get("https://portal.greenxtrade.com/")
    time.sleep(2)
    # Chuyển sang trang quotes nếu chưa tự redirect
    if "portal.greenxtrade.com" in driver.current_url and "/quotes" not in driver.current_url:
        driver.get(BASE_URL)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//input[@aria-label="input for From"]'))
    )
    print("[HỆ THỐNG] ✅ Trang quotes đã load xong — bắt đầu vòng lặp.")
except Exception as _e:
    print(f"[HỆ THỐNG] ⚠️ Không load được trang chủ EMC: {_e}. Thử BASE_URL...")
    try:
        driver.get(BASE_URL)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//input[@aria-label="input for From"]'))
        )
    except Exception as _e2:
        print(f"[HỆ THỐNG] ❌ Không mở được trang EMC: {_e2}")

for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    pol_excel = str(row[2] or "").strip().title()   # cột C = POL
    pod       = str(row[3] or "").strip().title()   # cột D = POD
    carrier   = str(row[4] or "").strip().upper()   # cột E = Carrier

    if not pol_excel or not pod: continue
    if carrier not in EMC_GROUP: continue
    if FILTER_POL and pol_excel.upper() != FILTER_POL: continue
    if FILTER_POD and pod.upper() != FILTER_POD: continue
    if first_emc_row is None:
        first_emc_row = i

    print(f"\n==========================================")
    print(f"--- DÒNG {i}: {pol_excel} → {pod} | Hãng: {carrier} ---")

    # FIX: Chỉ navigate lần đầu (lần sau dùng lại form đang hiện)
    if i == first_emc_row:
        # Đã navigate ở trên, chỉ cần đảm bảo form đang hiện
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@aria-label="input for From"]'))
            )
        except:
            # Nếu form biến mất (bị redirect lạ), navigate lại
            driver.get(BASE_URL)
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//input[@aria-label="input for From"]'))
            )
        driver.switch_to.window(driver.current_window_handle)
        driver.execute_script("window.focus();")
        time.sleep(1)

    result = search_one(pol_excel, pod)

    if result.get("error"):
        ws.cell(row=i, column=6).value  = result["error"]
    else:
        ws.cell(row=i, column=6).value  = result["price_20"]
        ws.cell(row=i, column=7).value  = result["price_40"]
        ws.cell(row=i, column=8).value  = result["price_40hq"]
        ws.cell(row=i, column=9).value  = result["etd"]
        ws.cell(row=i, column=10).value = result["tt"]
        ws.cell(row=i, column=13).value = result["remark"]

    try:
        wb.save(excel_path)
        print(f"   💾 Đã lưu dòng {i}")
    except PermissionError:
        print(f"   ❌ LỖI GHI FILE: TẮT FILE EXCEL ĐI!")

    time.sleep(2)

print(f"\n✅ Hoàn tất!")
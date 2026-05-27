"""
HPL (Hapag-Lloyd) Price Checker — 3 Tab Pipeline (Optimized)
Flow:
  Bước 1: Tab1 search → Tab2 search → Tab3 search
  Bước 2 (vòng lặp):
    Tab1: lấy giá → back → search mới (chỉ nhập field thay đổi)
    Tab2: lấy giá → back → search mới
    Tab3: lấy giá → back → search mới
    lặp lại cho đến hết
Tối ưu:
  - Alt+← back thay vì hash change → giữ form state
  - Cache POL/POD per tab, chỉ nhập lại field khác biệt
  - Sort queue theo POL để tối đa tái sử dụng
  - Chỉ đọc Price Breakdown 1 lần (cùng QQ price = cùng breakdown)
  - Giảm sleep thừa, bỏ human_scroll() khi search
  - Security check nhẹ (check URL thay vì page_source)
Selenium + Edge port 9525
"""

import math
import re
import time
import random
import os
from collections import defaultdict
from datetime import datetime, timedelta

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

if not is_port_in_use(9525):
    print("[HỆ THỐNG] Edge HPL chưa mở. Đang tự động khởi động...")
    try:
        subprocess.Popen([
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            "--remote-debugging-port=9525",
            r"--user-data-dir=C:\edge_hpl",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-backgrounding-occluded-windows"
        ])
        time.sleep(3)
    except: pass
else:
    print("[HỆ THỐNG] Edge HPL đã mở sẵn. Bỏ qua lệnh khởi động trình duyệt.")

edge_options = Options()
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9525")
service = Service(executable_path=driver_path)
driver  = webdriver.Edge(service=service, options=edge_options)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(document, 'visibilityState', {get: () => 'visible'});
        Object.defineProperty(document, 'hidden', {get: () => false});
        document.addEventListener('visibilitychange', e => e.stopImmediatePropagation(), true);
    """
})

BASE_URL  = "https://www.hapag-lloyd.com/solutions/new-quote/#/simple"
HPL_GROUP = {"HPL", "HAPAG", "HAPAG-LLOYD", "HAPAG LLOYD"}

POL_MAP = {
    "HO CHI MINH": "VNSGN",
    "HAI PHONG":   "VNHPH",
    "DA NANG":     "VNDAD",
    "ĐÀ NẴNG":     "VNDAD",
}

EUR_TO_USD = 1.08

EUROPE_PORTS = {
    "ALGECIRAS", "ANTWERP", "BARCELONA", "FOS SUR MER", "GENOA",
    "HAMBURG", "LE HAVRE", "NAPOLI", "ROTTERDAM", "VALENCIA", "VENEZIA",
    "FELIXSTOWE", "SOUTHAMPTON", "LONDON GATEWAY", "TILBURY",
    "BREMERHAVEN", "BREMEN", "AMSTERDAM", "ZEEBRUGGE", "DUNKIRK",
    "MARSEILLE", "BORDEAUX", "NANTES", "MONTOIR",
    "BILBAO", "VIGO", "SINES", "LISBON", "LEIXOES", "SETUBAL",
    "GIOIA TAURO", "TARANTO", "LA SPEZIA", "LIVORNO", "TRIESTE",
    "VENICE", "RAVENNA", "SALERNO", "CIVITAVECCHIA",
    "GOTHENBURG", "OSLO", "AARHUS", "COPENHAGEN", "HELSINKI",
    "STOCKHOLM", "TALLINN", "RIGA", "KLAIPEDA", "GDANSK", "GDYNIA",
    "SZCZECIN", "ROSTOCK",
    "PIRAEUS", "THESSALONIKI", "CONSTANTA", "VARNA", "BURGAS",
    "SPLIT", "RIJEKA", "KOPER", "DUBROVNIK",
    "TANGER MED",
    "BELFAST", "DUBLIN", "CORK", "LIVERPOOL", "BRISTOL",
}

POD_ALIASES = {
    "FOS SUR MER":    ["Fos", "Fos Sur Mer", "Fos-Sur-Mer"],
    "LE HAVRE":       ["Le Havre", "LeHavre"],
    "MARSEILLE":      ["Marseille", "Marseilles"],
    "DUNKIRK":        ["Dunkirk", "Dunkerque"],
    "MONTOIR":        ["Montoir", "Montoir-de-Bretagne", "Saint-Nazaire"],
    "GENOA":          ["Genoa", "Genova", "Genes"],
    "NAPOLI":         ["Napoli", "Naples", "Naple"],
    "LA SPEZIA":      ["La Spezia", "La-Spezia", "Spezia"],
    "LIVORNO":        ["Livorno", "Leghorn"],
    "TRIESTE":        ["Trieste", "Triest"],
    "VENEZIA":        ["Venezia", "Venice", "Venise"],
    "RAVENNA":        ["Ravenna"],
    "GIOIA TAURO":    ["Gioia Tauro", "Gioia-Tauro", "Gioia"],
    "CIVITAVECCHIA":  ["Civitavecchia", "Rome"],
    "SALERNO":        ["Salerno"],
    "TARANTO":        ["Taranto"],
    "ALGECIRAS":      ["Algeciras"],
    "BARCELONA":      ["Barcelona"],
    "VALENCIA":       ["Valencia"],
    "BILBAO":         ["Bilbao"],
    "VIGO":           ["Vigo"],
    "SINES":          ["Sines"],
    "LISBON":         ["Lisbon", "Lisboa"],
    "LEIXOES":        ["Leixoes", "Porto", "Matosinhos"],
    "ROTTERDAM":      ["Rotterdam"],
    "AMSTERDAM":      ["Amsterdam"],
    "ANTWERP":        ["Antwerp", "Antwerpen", "Anvers"],
    "ZEEBRUGGE":      ["Zeebrugge", "Zeebruges"],
    "HAMBURG":        ["Hamburg"],
    "BREMERHAVEN":    ["Bremerhaven", "Bremen"],
    "FELIXSTOWE":     ["Felixstowe"],
    "SOUTHAMPTON":    ["Southampton"],
    "LONDON GATEWAY": ["London Gateway", "London"],
    "TILBURY":        ["Tilbury"],
    "LIVERPOOL":      ["Liverpool"],
    "GOTHENBURG":     ["Gothenburg", "Goteborg", "Göteborg"],
    "OSLO":           ["Oslo"],
    "AARHUS":         ["Aarhus", "Arhus"],
    "COPENHAGEN":     ["Copenhagen", "Kobenhavn"],
    "HELSINKI":       ["Helsinki"],
    "STOCKHOLM":      ["Stockholm"],
    "GDANSK":         ["Gdansk", "Danzig"],
    "GDYNIA":         ["Gdynia"],
    "KLAIPEDA":       ["Klaipeda"],
    "RIGA":           ["Riga"],
    "TALLINN":        ["Tallinn", "Tallin"],
    "PIRAEUS":        ["Piraeus", "Pireaus", "Pireas", "Athens"],
    "THESSALONIKI":   ["Thessaloniki", "Salonika", "Salonica"],
    "CONSTANTA":      ["Constanta", "Constanța"],
    "KOPER":          ["Koper", "Capodistria"],
    "RIJEKA":         ["Rijeka", "Fiume"],
    "TANGER MED":     ["Tanger Med", "Tangier Med", "Tanger", "Tangier"],
}

# Port code ưu tiên khi dropdown có nhiều kết quả cùng tên (vd Rotterdam NL vs Rotterdam US)
PORT_PREFERRED_CODE = {
    "Rotterdam":  "NLRTM",
    "Hamburg":    "DEHAM",
    "Antwerp":    "BEANR",
    "Barcelona":  "ESBCN",
    "Valencia":   "ESVLC",
    "Genoa":      "ITGOA",
    "Genova":     "ITGOA",
    "Naples":     "ITNAP",
    "Venice":     "ITVCE",
    "Le Havre":   "FRLEH",
    "Algeciras":  "ESALG",
    "Piraeus":    "GRPIR",
    "Felixstowe": "GBFXT",
    "Southampton":"GBSOU",
    "Gothenburg": "SEGOT",
}

NUM_TABS = 3

# Track last POL/POD per tab để skip nhập lại field không đổi
tab_last_pol = [None] * NUM_TABS  # last POL entered per tab
tab_last_pod = [None] * NUM_TABS  # last POD entered per tab
cookie_dismissed = False          # chỉ dismiss cookie 1 lần

# ===================================================================================
# --- HELPERS ---
# ===================================================================================
def rand_sleep(a=0.3, b=0.8):
    time.sleep(random.uniform(a, b))

def activate_tab(handle):
    try:
        driver.execute_script("""
            window.dispatchEvent(new Event('focus'));
            document.dispatchEvent(new Event('focus'));
        """)
        rand_sleep(0.1, 0.2)
    except:
        pass

def human_move_and_click(element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    rand_sleep(0.1, 0.2)
    ActionChains(driver).move_to_element_with_offset(
        element, random.randint(-4, 4), random.randint(-3, 3)
    ).pause(random.uniform(0.05, 0.15)).click().perform()
    rand_sleep(0.1, 0.2)

def go_back():
    """Alt+← (browser back) — nhanh hơn hash change, giữ nguyên form state."""
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ALT + Keys.ARROW_LEFT)
    except:
        driver.execute_script("window.history.back();")
    try:
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
        )
    except:
        driver.execute_script("window.location.hash = '/simple';")
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
        )
    rand_sleep(0.3, 0.5)

def reload_tab():
    for attempt in range(2):
        try:
            driver.get(BASE_URL)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
            )
            rand_sleep(1.5, 2.5)
            return
        except Exception as e:
            if attempt == 0:
                print(f"   ⚠️ reload_tab lỗi lần 1: {type(e).__name__}, thử lại...")
                time.sleep(3)
            else:
                raise

def fmt_date(raw):
    for fmt in ("%Y-%m-%d", "%d %b %Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(raw.strip(), fmt).strftime("%d-%b")
        except:
            continue
    return raw.strip()

# ===================================================================================
# --- CHỌN CẢNG ---
# ===================================================================================
def select_port_hpl(selector, port_name, aliases=None):
    names_to_try = [port_name]
    if aliases:
        names_to_try += [a for a in aliases if a != port_name]

    for name in names_to_try:
        for attempt in range(2):
            try:
                inp = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )

                # Chỉ giữ nguyên nếu field đã chứa đúng port_name gốc
                current_val = driver.execute_script("return arguments[0].value;", inp)
                if current_val and "(" in current_val and port_name.upper() in current_val.upper():
                    print(f"        -> Giữ nguyên: {current_val.strip()}")
                    return

                # Luôn fetch inp mới (tránh stale sau bất kỳ DOM re-render nào)
                inp = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
                rand_sleep(0.1, 0.2)
                driver.execute_script("arguments[0].click(); arguments[0].focus();", inp)
                rand_sleep(0.2, 0.3)

                # Fetch lại inp ngay trước send_keys (Quasar re-render sau click)
                inp = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                inp.send_keys(Keys.CONTROL + "a")
                inp.send_keys(Keys.DELETE)
                rand_sleep(0.15, 0.25)
                inp.send_keys(name)
                rand_sleep(0.3, 0.5)

                # Chờ dropdown HOẶC No results
                WebDriverWait(driver, 6).until(lambda d:
                    d.find_elements(By.CSS_SELECTOR, ".q-menu .q-item") or
                    d.find_elements(By.CSS_SELECTOR, ".q-item__section--no-results")
                )

                if driver.find_elements(By.CSS_SELECTOR, ".q-item__section--no-results"):
                    print(f"        ⚠️ [{name}] No results → thử tên tiếp...")
                    break

                suggestions = driver.find_elements(By.CSS_SELECTOR, ".q-menu .q-item")
                rand_sleep(0.2, 0.3)  # chờ animation dropdown xong
                # Dùng JS check thay vì is_displayed() vì hay fail khi tab nền
                visible = [s for s in suggestions if driver.execute_script(
                    "return arguments[0].offsetParent !== null && arguments[0].offsetHeight > 0;", s
                )]
                if not visible:
                    visible = suggestions  # fallback: thử tất cả nếu JS cũng trống
                if visible:
                    # Ưu tiên chọn item có preferred port code (tránh chọn nhầm cùng tên ở nước khác)
                    preferred_code = PORT_PREFERRED_CODE.get(name)
                    chosen = None
                    if preferred_code:
                        for s in visible:
                            if preferred_code in s.text:
                                chosen = s
                                break
                    if not chosen:
                        chosen = visible[0]
                    human_move_and_click(chosen)
                    print(f"        -> Chốt [{name}]: {chosen.text.strip()[:60]}")
                    rand_sleep(0.2, 0.3)
                    return
                else:
                    raise Exception("Dropdown trống")

            except Exception as e:
                print(f"        ⚠️ [{name}] lần {attempt+1}: {e}")
                rand_sleep(0.3, 0.5)

    raise Exception(f"Thất bại tất cả aliases: {names_to_try}")

# ===================================================================================
# --- ĐẶT NGÀY ---
# ===================================================================================
def set_date_hpl():
    target = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
    try:
        date_inp = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="validity-input"]'))
        )
        if date_inp.get_attribute("value") == target:
            return
        driver.execute_script("""
            var inp = arguments[0]; var val = arguments[1];
            inp.focus();
            var setter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value').set;
            setter.call(inp, val);
            inp.dispatchEvent(new Event('input', { bubbles: true }));
            inp.dispatchEvent(new Event('change', { bubbles: true }));
        """, date_inp, target)
        rand_sleep(0.2, 0.3)
    except Exception as e:
        print(f"      ⚠️ Lỗi đặt ngày: {e}")

# ===================================================================================
# --- KIỂM TRA SECURITY CHECK ---
# ===================================================================================
def check_security_block(tab_idx):
    try:
        url = driver.current_url or ""
        if "security" in url.lower() or "challenge" in url.lower():
            print(f"   [Tab{tab_idx}] ⚠️ Phát hiện chặn Security Check!")
            if not os.environ.get("EXCEL_PATH"):
                input(f"   >>> Hãy tự tick xác thực trên Tab {tab_idx}, chờ web load xong rồi nhấn [ENTER] tại đây để tiếp tục... <<<")
            else:
                print(f"   [Tab{tab_idx}] ⏳ Chạy qua main.py — chờ 90s để user giải captcha...")
                time.sleep(90)
            # FIX: Sau khi giải captcha, reload từ BASE_URL để reset form → tránh bị detect lại
            print(f"   [Tab{tab_idx}] 🔄 Reload BASE_URL sau khi giải captcha...")
            try:
                driver.get(BASE_URL)
                WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
                )
                print(f"   [Tab{tab_idx}] ✅ Form sẵn sàng sau captcha.")
            except Exception as _reload_err:
                print(f"   [Tab{tab_idx}] ⚠️ Reload sau captcha gặp lỗi: {_reload_err}")
            return
        title = driver.title or ""
        if "Security" in title and "Check" in title:
            print(f"   [Tab{tab_idx}] ⚠️ Phát hiện chặn Security Check!")
            if not os.environ.get("EXCEL_PATH"):
                input(f"   >>> Hãy tự tick xác thực trên Tab {tab_idx}, chờ web load xong rồi nhấn [ENTER] tại đây để tiếp tục... <<<")
            else:
                print(f"   [Tab{tab_idx}] ⏳ Chạy qua main.py — chờ 90s để user giải captcha...")
                time.sleep(90)
            # FIX: Reload sau captcha
            print(f"   [Tab{tab_idx}] 🔄 Reload BASE_URL sau khi giải captcha...")
            try:
                driver.get(BASE_URL)
                WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
                )
                print(f"   [Tab{tab_idx}] ✅ Form sẵn sàng sau captcha.")
            except Exception as _reload_err:
                print(f"   [Tab{tab_idx}] ⚠️ Reload sau captcha gặp lỗi: {_reload_err}")
    except:
        pass

# ===================================================================================
# --- NHẬP & SEARCH ---
# ===================================================================================
def do_search(pol, pod, tab_idx):
    global cookie_dismissed
    t = tab_idx - 1  # 0-based index

    print(f"   [Tab{tab_idx}] 📍 Nhập {pol} → {pod}")
    check_security_block(tab_idx)

    # Dismiss cookie/notification banner — chỉ 1 lần
    if not cookie_dismissed:
        try:
            driver.execute_script("""
                ['#onetrust-accept-btn-handler', '.onetrust-close-btn-handler',
                 '.hl-notification__close', '.hl-cookie-banner__close'
                ].forEach(function(s){
                    var el = document.querySelector(s); if (el) el.click();
                });
            """)
            rand_sleep(0.2, 0.3)
            cookie_dismissed = True
        except: pass

    # --- SMART POL: chỉ nhập lại nếu khác row trước ---
    if tab_last_pol[t] == pol:
        # Check xem field có còn giữ giá trị cũ không
        try:
            inp = driver.find_element(By.CSS_SELECTOR, 'input[data-testid="start-input"]')
            cur = driver.execute_script("return arguments[0].value;", inp) or ""
            if pol.upper() in cur.upper():
                print(f"        -> POL giữ nguyên: {cur.strip()}")
            else:
                select_port_hpl('input[data-testid="start-input"]', pol)
                rand_sleep(0.3, 0.5)
        except:
            select_port_hpl('input[data-testid="start-input"]', pol)
            rand_sleep(0.3, 0.5)
    else:
        select_port_hpl('input[data-testid="start-input"]', pol)
        rand_sleep(0.3, 0.5)
    tab_last_pol[t] = pol

    # --- SMART POD: chỉ nhập lại nếu khác row trước ---
    pod_upper   = pod.upper()
    pod_aliases = POD_ALIASES.get(pod_upper, [pod])
    if tab_last_pod[t] == pod:
        try:
            inp = driver.find_element(By.CSS_SELECTOR, 'input[data-testid="end-input"]')
            cur = driver.execute_script("return arguments[0].value;", inp) or ""
            first_alias = pod_aliases[0]
            if first_alias.upper() in cur.upper() or pod_upper in cur.upper():
                print(f"        -> POD giữ nguyên: {cur.strip()}")
            else:
                select_port_hpl('input[data-testid="end-input"]', pod_aliases[0], aliases=pod_aliases)
        except:
            select_port_hpl('input[data-testid="end-input"]', pod_aliases[0], aliases=pod_aliases)
    else:
        select_port_hpl('input[data-testid="end-input"]', pod_aliases[0], aliases=pod_aliases)
    tab_last_pod[t] = pod

    set_date_hpl()

    search_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
        (By.XPATH, "//button[.//span[text()='Search']]")))
    human_move_and_click(search_btn)
    print(f"   [Tab{tab_idx}] 🖱️ Đã Search")

# ===================================================================================
# --- ĐỌC VALID DATE TỪ SIDEBAR ---
# ===================================================================================
def read_valid_from_sidebar(tab_idx):
    try:
        el = WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.XPATH,
                '//*[@id="hlMain"]/div[2]/div[2]/div/div/div/div[3]/div[2]/div[2]'
                '/div[1]/div[2]/div/div[1]/span/span[2]'
            ))
        )
        raw = el.text.strip()
        if raw:
            return fmt_date(raw)
    except:
        pass
    try:
        spans = driver.find_elements(By.CSS_SELECTOR, ".sidebar-schedule span.text-button-s")
        for sp in spans:
            raw = sp.text.strip()
            if re.match(r'\d{4}-\d{2}-\d{2}', raw) or re.match(r'\d{2} \w+ \d{4}', raw):
                return fmt_date(raw)
    except:
        pass
    return ""

# ===================================================================================
# --- ĐỌC CARDS ---
# FIX: track (etd_str, tt_days) combo để bỏ card trùng qua Next
# ===================================================================================
def read_all_cards(tab_idx):
    print(f"   [Tab{tab_idx}] 📋 Đọc cards...")
    check_security_block(tab_idx)
    result     = []
    seen_combo = set()  # track (etd_str, tt_days) để loại trùng

    try:
        WebDriverWait(driver, 30).until(lambda d:
            d.find_elements(By.CSS_SELECTOR, "button.carousel__item") or
            d.find_elements(By.CSS_SELECTOR, ".q-notification--danger")
        )
    except:
        raise Exception("NO SERVICE / SOLD OUT")

    if driver.find_elements(By.CSS_SELECTOR, ".q-notification--danger"):
        try:
            dismiss = driver.find_element(By.XPATH, "//button[.//span[text()='Dismiss']]")
            driver.execute_script("arguments[0].click();", dismiss)
        except: pass
        raise Exception("NO SERVICE / SOLD OUT")

    rand_sleep(0.5, 1.0)
    max_pages   = 6
    page_count  = 0

    while page_count < max_pages:
        cards = driver.find_elements(By.CSS_SELECTOR, "button.carousel__item")

        for idx in range(len(cards)):
            try:
                cards = driver.find_elements(By.CSS_SELECTOR, "button.carousel__item")
                card  = cards[idx]

                etd_span = card.find_element(By.CSS_SELECTOR, "div.carousel__date span")
                etd_str  = etd_span.text.strip()
                if not etd_str:
                    continue

                aria = card.get_attribute("aria-label") or ""
                qq_match = re.search(r"Quick Quotes: USD([\d ,]+)", aria)
                if not qq_match:
                    print(f"   [Tab{tab_idx}] ⏭️ Skip {etd_str} (no price)")
                    continue

                base_price = int(qq_match.group(1).strip().replace(" ", "").replace(",", ""))
                etd_dt     = datetime.strptime(etd_str, "%Y-%m-%d")

                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
                rand_sleep(0.05, 0.1)
                driver.execute_script("arguments[0].click();", card)
                rand_sleep(0.05, 0.1)

                try:
                    transit_el = WebDriverWait(driver, 4).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".sidebar-schedule__days"))
                    )
                    tt_days = int(re.search(r'\d+', transit_el.text.strip()).group())
                except:
                    tt_days = 999

                # Bỏ combo trùng (etd + tt) — tránh đọc lại card cũ sau Next
                combo = (etd_str, tt_days)
                if combo in seen_combo:
                    continue
                seen_combo.add(combo)

                valid_to = read_valid_from_sidebar(tab_idx)

                print(f"   [Tab{tab_idx}] ✅ ETD={etd_str} TT={tt_days}d Price={base_price} Valid={valid_to}")
                result.append({
                    "etd_dt":     etd_dt,
                    "tt_days":    tt_days,
                    "etd_str":    etd_str,
                    "base_price": base_price,
                    "valid_to":   valid_to,
                })

            except Exception as e:
                print(f"   [Tab{tab_idx}] ⚠️ Lỗi card idx={idx}: {e}")
                continue

        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Next Departures']")
            if next_btn.get_attribute("disabled") is None:
                human_move_and_click(next_btn)
                page_count += 1
                print(f"   [Tab{tab_idx}] ➡️ Next ({page_count}/{max_pages})...")
                rand_sleep(0.5, 0.8)
            else:
                break
        except:
            break

    return result

# ===================================================================================
# --- RESET CAROUSEL ---
# ===================================================================================
def reset_carousel(tab_idx):
    for _ in range(20):
        try:
            prev_btn = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Previous Departures']")
            if prev_btn.get_attribute("disabled") is not None:
                break
            driver.execute_script("arguments[0].click();", prev_btn)
            rand_sleep(0.15, 0.25)
        except:
            break

# ===================================================================================
# --- ĐỌC LỊCH TÀU TỪ POPUP "View Departure Details" ---
# ===================================================================================
def read_departure_details(tab_idx):
    """
    Click 'View Departure Details' → đọc tất cả card trong popup.
    Trả về list dict: etd_dt, etd_str, base_price, tt_days,
           vessel, service, transshipment, arrival, valid_to
    Trả về None nếu popup không mở được (caller sẽ fallback sang read_all_cards).
    """
    print(f"   [Tab{tab_idx}] 📋 Mở popup Departure Details...")
    check_security_block(tab_idx)

    # Chờ carousel load xong trước
    try:
        WebDriverWait(driver, 30).until(lambda d:
            d.find_elements(By.CSS_SELECTOR, "button.carousel__item") or
            d.find_elements(By.CSS_SELECTOR, ".q-notification--danger")
        )
    except:
        raise Exception("NO SERVICE / SOLD OUT")

    if driver.find_elements(By.CSS_SELECTOR, ".q-notification--danger"):
        try:
            dismiss = driver.find_element(By.XPATH, "//button[.//span[text()='Dismiss']]")
            driver.execute_script("arguments[0].click();", dismiss)
        except: pass
        raise Exception("NO SERVICE / SOLD OUT")

    rand_sleep(0.5, 1.0)

    # Click "View Departure Details"
    try:
        detail_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[.//span[text()='View Departure Details']]")))
        human_move_and_click(detail_btn)
        rand_sleep(0.8, 1.2)
    except Exception as e:
        print(f"   [Tab{tab_idx}] ⚠️ Không click được 'View Departure Details': {e}")
        return None

    # Chờ dialog hiện
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".q-dialog .q-dialog__content"))
        )
        rand_sleep(0.5, 0.8)
    except:
        print(f"   [Tab{tab_idx}] ⚠️ Popup không hiện")
        return None

    try:
        dialog = driver.find_element(By.CSS_SELECTOR, ".q-dialog")
    except:
        print(f"   [Tab{tab_idx}] ⚠️ Không tìm thấy dialog")
        return None

    # Đảm bảo chip "Routing Details" được chọn
    try:
        chips = dialog.find_elements(By.CSS_SELECTOR, ".q-chip")
        for chip in chips:
            if "Routing" in chip.text and "q-chip--selected" not in (chip.get_attribute("class") or ""):
                chip.click()
                rand_sleep(0.3, 0.5)
                break
    except:
        pass

    result     = []
    seen_combo = set()
    max_pages  = 6
    page_count = 0

    while page_count < max_pages:
        try:
            cards = dialog.find_elements(By.CSS_SELECTOR, "button.carousel__item")
        except:
            break

        for card in cards:
            try:
                # ETD từ div[1]/span
                etd_el  = card.find_element(By.XPATH, "./div[1]/span")
                etd_str = etd_el.text.strip()
                if not etd_str:
                    continue

                # QQ Price từ aria-label
                aria     = card.get_attribute("aria-label") or ""
                qq_match = re.search(r"Quick Quotes:\s*USD\s*([\d.,\s]+)", aria)
                if not qq_match:
                    print(f"   [Tab{tab_idx}] ⏭️ Skip {etd_str} (no price)")
                    continue
                price_str  = qq_match.group(1).strip().replace(" ", "").replace(",", "")
                base_price = int(float(price_str))

                # Transit time từ div[5]
                try:
                    tt_div   = card.find_element(By.XPATH, "./div[5]")
                    tt_match = re.search(r'(\d+)', tt_div.text.strip())
                    tt_days  = int(tt_match.group(1)) if tt_match else 999
                except:
                    tt_days = 999

                # Arrival date từ div[4]
                try:
                    arr_div     = card.find_element(By.XPATH, "./div[4]")
                    arrival_str = arr_div.text.strip()
                except:
                    arrival_str = ""

                # Vessel name từ div[6]
                try:
                    vessel_el   = card.find_element(By.XPATH, "./div[6]//div[contains(@class,'ellipsis')]")
                    vessel_name = vessel_el.text.strip()
                except:
                    vessel_name = ""

                # Service từ div[7]
                try:
                    service_el   = card.find_element(By.XPATH, "./div[7]//div[contains(@class,'ellipsis')]")
                    service_name = service_el.text.strip()
                except:
                    service_name = ""

                # Via/Transshipment từ div[11]
                try:
                    via_el   = card.find_element(By.XPATH, "./div[11]")
                    via_subs = via_el.find_elements(By.CSS_SELECTOR, "div.ellipsis")
                    if via_subs:
                        via_text = via_subs[0].text.strip()
                    else:
                        via_text = via_el.text.strip()
                    if via_text == "-" or not via_text:
                        transshipment = "DIRECT"
                    else:
                        transshipment = via_text.upper()
                except:
                    transshipment = ""

                # Dedup theo (etd, tt, vessel) — tránh đọc trùng sau Next
                combo = (etd_str, tt_days, vessel_name)
                if combo in seen_combo:
                    continue
                seen_combo.add(combo)

                etd_dt = datetime.strptime(etd_str, "%Y-%m-%d")
                result.append({
                    "etd_dt":        etd_dt,
                    "etd_str":       etd_str,
                    "tt_days":       tt_days,
                    "base_price":    base_price,
                    "arrival":       arrival_str,
                    "vessel":        vessel_name,
                    "service":       service_name,
                    "transshipment": transshipment,
                    "valid_to":      "",
                })
                print(f"   [Tab{tab_idx}] 🚢 ETD={etd_str} Price={base_price} TT={tt_days}d "
                      f"Vessel={vessel_name} T/S={transshipment}")

            except Exception as e:
                print(f"   [Tab{tab_idx}] ⚠️ Lỗi đọc popup card: {e}")
                continue

        # Next page trong dialog
        try:
            next_btns = dialog.find_elements(By.CSS_SELECTOR, "button[aria-label='Next Departures']")
            if next_btns and next_btns[0].get_attribute("disabled") is None:
                human_move_and_click(next_btns[0])
                page_count += 1
                print(f"   [Tab{tab_idx}] ➡️ Popup Next ({page_count}/{max_pages})...")
                rand_sleep(0.5, 0.8)
            else:
                break
        except:
            break

    # Đóng dialog
    try:
        close_btn = dialog.find_element(By.CSS_SELECTOR, "button.q-dialog__x")
        driver.execute_script("arguments[0].click();", close_btn)
        rand_sleep(0.3, 0.5)
    except:
        try:
            dialog.send_keys(Keys.ESCAPE)
            rand_sleep(0.3, 0.5)
        except:
            pass

    print(f"   [Tab{tab_idx}] 📋 Popup: tìm thấy {len(result)} departures")
    return result if result else None

# ===================================================================================
# --- 9 QUY TẮC VÀNG ---
# ===================================================================================
def apply_9_golden_rules(danh_sach_chuyen):
    co_gia = [c for c in danh_sach_chuyen if c.get("base_price") is not None]
    if not co_gia:
        return [], "N/A", "N/A"

    # --- ĐIỀU KIỆN MỚI: Ưu tiên tháng hiện tại ---
    # Sắp xếp để tìm ngày ETD sớm nhất (tháng hiện tại của view)
    co_gia.sort(key=lambda x: x["etd_dt"])
    first_year, first_month = co_gia[0]["etd_dt"].year, co_gia[0]["etd_dt"].month
    
    # Lấy toàn bộ card thuộc tháng đầu tiên này
    thang_nay = [c for c in co_gia if c["etd_dt"].year == first_year and c["etd_dt"].month == first_month]
    
    # Đếm số lượng ETD khác nhau trong tháng này
    unique_etds_thang_nay = set([c["etd_dt"] for c in thang_nay])
    
    # Nếu tháng này có từ 3 ngày ETD trở lên -> Chốt chỉ chơi với tháng này, mặc kệ giá tháng sau
    if len(unique_etds_thang_nay) >= 3:
        co_gia = thang_nay
    # ---------------------------------------------

    min_price        = min(c["base_price"] for c in co_gia)
    same_price_group = [c for c in co_gia if c["base_price"] == min_price]
    same_price_group.sort(key=lambda x: (x["etd_dt"], x["tt_days"]))

    # Lọc trùng ETD — mỗi ETD chỉ giữ TT ngắn nhất
    seen, list_loc_trung = set(), []
    for c in same_price_group:
        if c["etd_dt"] not in seen:
            list_loc_trung.append(c)
            seen.add(c["etd_dt"])

    etd_dat_chuan = []
    if list_loc_trung:
        ngan_nhat  = min(c["tt_days"] for c in list_loc_trung)
        first_date = list_loc_trung[0]["etd_dt"]
        for c in list_loc_trung:
            if len(etd_dat_chuan) >= 3: break
            if len(etd_dat_chuan) > 0 and (c["etd_dt"] - etd_dat_chuan[-1]["etd_dt"]).days < 2: continue
            if (c["etd_dt"] - first_date).days <= 9 and c["tt_days"] <= ngan_nhat + 10:
                etd_dat_chuan.append(c)

    num = len(etd_dat_chuan)
    if num == 0:   str_etd = "N/A"
    elif num == 1: str_etd = etd_dat_chuan[0]["etd_dt"].strftime("%d-%b")
    elif num == 2: str_etd = (f"{etd_dat_chuan[0]['etd_dt'].strftime('%d-%b')} & "
                              f"{etd_dat_chuan[1]['etd_dt'].strftime('%d-%b')}")
    else:
        d1 = etd_dat_chuan[0]["etd_dt"].strftime("%d")
        d2 = etd_dat_chuan[1]["etd_dt"].strftime("%d")
        d3 = etd_dat_chuan[2]["etd_dt"].strftime("%d-%b")
        str_etd = f"{d1}, {d2}, {d3}"

    all_tt = [c["tt_days"] for c in etd_dat_chuan]
    str_tt = f"{min(all_tt)}" if min(all_tt) == max(all_tt) else f"{min(all_tt)}-{max(all_tt)}"
    return etd_dat_chuan, str_etd, str_tt

# ===================================================================================
# --- PARSE BẢNG GIÁ ---
# ===================================================================================
def parse_hpl_price(tab_idx, pod=""):
    charges      = {}
    has_thc_orig = False
    ows_entries  = []   # list of (weight_tons, price_20, price_40)
    in_ows       = False
    avail_cols   = set()

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".offer-charges"))
        )
        # Wait: đếm tổng row 2 lần liên tiếp cách nhau 0.5s phải bằng nhau → DOM ổn định
        def count_all_rows(d):
            tbls = d.find_elements(By.CSS_SELECTOR, ".offer-charges table.q-table")
            return sum(len(t.find_elements(By.CSS_SELECTOR, "tbody tr")) for t in tbls)
        try:
            WebDriverWait(driver, 15).until(lambda d: count_all_rows(d) > 0)
            import time as _t
            _t.sleep(0.2) 
        except:
            rand_sleep(3.0, 4.0)

        all_tables = driver.find_elements(By.CSS_SELECTOR, ".offer-charges table.q-table")
        print(f"   [Tab{tab_idx}] [Tables] {len(all_tables)} tables, {count_all_rows(driver)} rows")

        # PASS 1: Scan TẤT CẢ tables để detect THC (kể cả Export Surcharges)
        for tbl in all_tables:
            for row in tbl.find_elements(By.CSS_SELECTOR, "tbody tr"):
                tds = row.find_elements(By.CSS_SELECTOR, "td")
                if tds and "Terminal Handling Charge Orig" in tds[0].text:
                    has_thc_orig = True
                    print(f"   [Tab{tab_idx}] [THC] Tìm thấy")
                    break
            if has_thc_orig:
                break

        # PASS 2: Cộng giá từ Freight Charges + Freight Surcharges
        for table in all_tables:
            try:
                all_ths  = table.find_elements(By.CSS_SELECTOR, "thead th")
                th_texts = [th.text.strip() for th in all_ths]
                section  = th_texts[0] if th_texts else ""
            except:
                continue

            if "Freight" not in section:
                continue

            col_map = {}
            for i, txt in enumerate(th_texts):
                if txt in ("20STD", "40STD", "40HC"):
                    col_map[i] = txt
                    avail_cols.add(txt)

            for row in table.find_elements(By.CSS_SELECTOR, "tbody tr"):
                try:
                    tds = row.find_elements(By.CSS_SELECTOR, "td")
                    if len(tds) < 2: continue
                    charge_name = tds[0].text.strip().split("\n")[0]
                    currency    = tds[2].text.strip()

                    # --- OWS: Phát hiện Heavy Lift / Weight Tier → skip cộng giá, thu thập chi tiết ---
                    full_text = tds[0].text.strip()
                    is_ows_row = False
                    if "Heavy Lift" in charge_name or "Weight Tier" in charge_name:
                        in_ows = True
                        is_ows_row = True
                    elif in_ows and "Between" in full_text:
                        is_ows_row = True
                    else:
                        if charge_name and "Between" not in full_text:
                            in_ows = False

                    if is_ows_row:
                        # Parse weight threshold từ description
                        wt_match = re.search(r'[Bb]etween\s+([\d.]+)\s+and', full_text)
                        wt_tons = 0
                        if wt_match:
                            wt_tons = int(float(wt_match.group(1)))
                        # Parse giá cho từng container
                        ows_20 = 0
                        ows_40 = 0
                        for col_idx, cont_key in col_map.items():
                            if col_idx < len(tds):
                                val_str = re.sub(r'[\s,\u00a0\u202f]', '', tds[col_idx].text.strip())
                                if val_str and val_str != '-':
                                    try:
                                        v = float(val_str)
                                        if cont_key == '20STD':
                                            ows_20 = int(v)
                                        elif cont_key in ('40STD', '40HC') and ows_40 == 0:
                                            ows_40 = int(v)
                                    except: pass
                        if wt_tons > 0 and (ows_20 > 0 or ows_40 > 0):
                            ows_entries.append((wt_tons, ows_20, ows_40))
                        print(f"   [Tab{tab_idx}] [OWS] {full_text[:60]} → wt={wt_tons} 20'=${ows_20} 40'=${ows_40}")
                        continue

                    if col_map:
                        if currency == "USD":
                            rate = 1.0
                        elif currency == "EUR":
                            rate = EUR_TO_USD
                            print(f"   [Tab{tab_idx}] [EUR] {charge_name} → x{EUR_TO_USD}")
                        else:
                            continue

                        for col_idx, cont_key in col_map.items():
                            if col_idx < len(tds):
                                val_str = re.sub(r'[\s,\u00a0\u202f]', '', tds[col_idx].text.strip())
                                if val_str and val_str != "-":
                                    try:
                                        charges[cont_key] = charges.get(cont_key, 0.0) + float(val_str) * rate
                                    except: pass
                        log = [(k, tds[i].text.strip()) for i, k in col_map.items() if i < len(tds)]
                        print(f"   [Tab{tab_idx}] [+] {section} | {charge_name} ({currency}): {log}")

                except: continue

        if not has_thc_orig:
            print(f"   [Tab{tab_idx}] ⚠️ Không có THC Orig → Trừ $140/$210")
            if "20STD" in avail_cols: charges["20STD"] = charges.get("20STD", 0.0) - 140.0
            if "40STD" in avail_cols: charges["40STD"] = charges.get("40STD", 0.0) - 210.0
            if "40HC"  in avail_cols: charges["40HC"]  = charges.get("40HC",  0.0) - 210.0

        pod_upper   = pod.upper()
        is_europe   = any(ep in pod_upper for ep in EUROPE_PORTS)
        base_remark = "SUBJECT TO THC, BILL, SEAL, TLX" + (", ENS" if is_europe else "")

        # --- GHI REMARK ĐÚNG CHUẨN (bao gồm chi tiết OWS) ---
        if ows_entries:
            ows_parts = []
            for wt, p20, p40 in ows_entries:
                if p20 > 0:
                    ows_parts.append(f"OWS ${p20}/20' (>{wt} TONS)")
                if p40 > 0:
                    ows_parts.append(f"OWS ${p40}/40' (>{wt} TONS)")
            if ows_parts:
                remark = base_remark + ", " + ", ".join(ows_parts)
            else:
                remark = base_remark + ", OWS"
        else:
            remark = base_remark

    except Exception as e:
        print(f"   [Tab{tab_idx}] ⚠️ Lỗi parse giá: {e}")
        remark = "SUBJECT TO THC, BILL, SEAL, TLX"

    totals = {}
    for key in ("20STD", "40STD", "40HC"):
        totals[key] = math.ceil(charges[key]) if (key in avail_cols and charges.get(key, 0) > 0) else "-"

    return totals, remark

# ===================================================================================
# --- FORMAT ETD/TT ---
# ===================================================================================
def format_etd_tt(etd_list):
    """Rebuild str_etd, str_tt từ danh sách ETD đã lọc."""
    num = len(etd_list)
    if num == 0: return "N/A", "N/A"
    elif num == 1: s = etd_list[0]["etd_dt"].strftime("%d-%b")
    elif num == 2:
        s = (f"{etd_list[0]['etd_dt'].strftime('%d-%b')} & "
             f"{etd_list[1]['etd_dt'].strftime('%d-%b')}")
    else:
        s = (f"{etd_list[0]['etd_dt'].strftime('%d')}, "
             f"{etd_list[1]['etd_dt'].strftime('%d')}, "
             f"{etd_list[2]['etd_dt'].strftime('%d-%b')}")
    all_tt = [c["tt_days"] for c in etd_list]
    tt = str(min(all_tt)) if min(all_tt)==max(all_tt) else f"{min(all_tt)}-{max(all_tt)}"
    return s, tt

# ===================================================================================
# --- LẤY GIÁ & GHI EXCEL ---
# ===================================================================================
def get_price_and_save(row_i, tab_idx, wb, ws, job_pod=""):
    try:
        # Ưu tiên đọc từ popup Departure Details (nhanh hơn + có lịch tàu)
        all_cards = read_departure_details(tab_idx)
        if all_cards is None:
            print(f"   [Tab{tab_idx}] ⚠️ Popup thất bại, dùng carousel cũ")
            all_cards = read_all_cards(tab_idx)
        if not all_cards:
            ws.cell(row=row_i, column=6).value = "NO SERVICE"
            wb.save(excel_path); return

        etd_chuan, _, _ = apply_9_golden_rules(all_cards)
        if not etd_chuan:
            ws.cell(row=row_i, column=6).value = "Không có card đạt chuẩn"
            wb.save(excel_path); return

        # --- Click từng card ETD để đọc valid_to, chỉ mở Price Breakdown 1 lần ---
        def find_and_click_card(etd_target):
            """Tìm card theo ETD string, click vào, return True/False."""
            reset_carousel(tab_idx)
            rand_sleep(0.15, 0.25)
            for _ in range(6):
                for card in driver.find_elements(By.CSS_SELECTOR, "button.carousel__item"):
                    try:
                        if card.find_element(By.CSS_SELECTOR, "div.carousel__date span").text.strip() == etd_target:
                            human_move_and_click(card)
                            rand_sleep(0.2, 0.3)
                            return True
                    except: continue
                try:
                    next_btn = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Next Departures']")
                    if next_btn.get_attribute("disabled") is None:
                        driver.execute_script("arguments[0].click();", next_btn)
                        rand_sleep(0.3, 0.5)
                    else:
                        break
                except:
                    break
            return False

        totals = None
        remark = ""
        valid_to_list = []
        best_etd_list = []

        for idx_etd, etd_info in enumerate(etd_chuan):
            etd_target = etd_info["etd_str"]

            if not find_and_click_card(etd_target):
                print(f"   [Tab{tab_idx}] ⚠️ Không tìm lại card ETD={etd_target}")
                continue

            # Đọc valid_to từ sidebar cho MỖI card
            vt = read_valid_from_sidebar(tab_idx) or etd_info.get("valid_to", "")
            valid_to_list.append(vt)
            best_etd_list.append(etd_info)
            print(f"   [Tab{tab_idx}] 📅 ETD={etd_target} Valid={vt}")

            # Chỉ mở Price Breakdown ở card ĐẦU TIÊN (cùng QQ price = cùng breakdown)
            if idx_etd == 0:
                try:
                    price_btn = WebDriverWait(driver, 8).until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[.//span[text()='Price Breakdown']]")))
                    human_move_and_click(price_btn)
                    rand_sleep(0.2, 0.3)
                except Exception as e:
                    print(f"   [Tab{tab_idx}] ⚠️ Không click Price Breakdown: {e}")
                    ws.cell(row=row_i, column=6).value = "Không mở Price Breakdown"
                    wb.save(excel_path); return

                totals, remark = parse_hpl_price(tab_idx, pod=job_pod)
                print(f"   [Tab{tab_idx}] 💰 ETD={etd_target}: 20'={totals.get('20STD')} 40'={totals.get('40STD')} 40HC={totals.get('40HC')}")

                # Đóng Price Breakdown
                try:
                    close_btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.q-dialog__x"))
                    )
                    driver.execute_script("arguments[0].click();", close_btn)
                    rand_sleep(0.2, 0.3)
                except:
                    try:
                        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                        rand_sleep(0.2, 0.3)
                    except: pass

        if not best_etd_list or totals is None:
            ws.cell(row=row_i, column=6).value = "Không có giá hợp lệ"
            wb.save(excel_path); return

        # Valid lấy từ card ETD cuối cùng
        valid_to = valid_to_list[-1] if valid_to_list else ""
        str_etd, str_tt = format_etd_tt(best_etd_list)

        print(f"   [Tab{tab_idx}] 🏆 ETD: {str_etd} | T/T: {str_tt}")
        ws.cell(row=row_i, column=6).value  = totals.get("20STD")
        ws.cell(row=row_i, column=7).value  = totals.get("40STD")
        ws.cell(row=row_i, column=8).value  = totals.get("40HC")
        ws.cell(row=row_i, column=9).value  = str_etd
        ws.cell(row=row_i, column=10).value = str_tt
        ws.cell(row=row_i, column=11).value = valid_to
        ws.cell(row=row_i, column=13).value = remark

        # Cột O (15): Tất cả tàu đạt chuẩn — nhiều dòng trong 1 ô
        vessel_lines = []
        for e_info in best_etd_list:
            v_name = e_info.get("vessel", "")
            e_date = e_info["etd_dt"].strftime("%d-%b").lstrip("0")
            e_tt   = e_info.get("tt_days", "")
            e_ts   = e_info.get("transshipment", "")
            vessel_lines.append(
                f"{v_name} / ETD: {e_date} / Transit time: {e_tt} Days / Transshipment: {e_ts}"
            )
        ws.cell(row=row_i, column=15).value = "\n".join(vessel_lines)

        # Cột P (16): Transshipment ports
        ts_per_card = []
        for e_info in best_etd_list:
            ts = e_info.get("transshipment", "")
            if ts and ts != "DIRECT":
                ts_per_card.append(ts)
        if not ts_per_card:
            ts_col = "DIRECT"
        else:
            unique_ts = list(dict.fromkeys(ts_per_card))  # giữ thứ tự
            ts_col = " or ".join(unique_ts)
        ws.cell(row=row_i, column=16).value = ts_col

        print(f"   [Tab{tab_idx}] 📅 Valid={valid_to} | 📝 {remark}")
        print(f"   [Tab{tab_idx}] 🚢 Vessels: {len(vessel_lines)} | T/S: {ts_col}")
        try:
            wb.save(excel_path)
            print(f"   [Tab{tab_idx}] 💾 Saved dòng {row_i}")
        except PermissionError:
            print(f"   [Tab{tab_idx}] ❌ Tắt Excel đi!")

    except Exception as e:
        print(f"   [Tab{tab_idx}] ❌ Lỗi: {e}")
        ws.cell(row=row_i, column=6).value = f"LỖI: {e}"
        try: wb.save(excel_path)
        except: pass

# ===================================================================================
# --- MAIN ---
# ===================================================================================

# ===================================================================================
# --- ĐĂNG NHẬP ---
# ===================================================================================
def handle_login():
    print("   [HỆ THỐNG] Kiểm tra trạng thái đăng nhập...")
    driver.get(BASE_URL)
    time.sleep(5) # Đợi trang load và chuyển hướng nếu chưa login
    
    if "identity.hapag-lloyd.com" in driver.current_url:
        print("   [HỆ THỐNG] Đang ở trang đăng nhập. Tự động click nút 'Log in'...")
        try:
            login_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "next"))
            )
            login_btn.click()
            print("   [HỆ THỐNG] Đã click xong. Đang chờ tải trang...")
        except Exception as e:
            print("   [HỆ THỐNG] Không tìm thấy nút Log in, vui lòng tự thao tác trên web.")
        
        print("   [HỆ THỐNG] ⚠️ Nếu có xác thực con người (Captcha), hãy tự giải quyết trên màn hình Edge.")
        if not os.environ.get("EXCEL_PATH"):
            input("   >>> SAU KHI VÀO ĐƯỢC TRANG TÌM GIÁ, HÃY NHẤN [ENTER] TẠI MÀN HÌNH CONSOLE NÀY ĐỂ CHẠY TIẾP... <<<")
        else:
            print("   [HỆ THỐNG] ⏳ Chạy qua main.py — chờ tối đa 120s để login hoàn tất...")
            for _ in range(24):
                time.sleep(5)
                try:
                    cur = driver.current_url or ""
                    if "hapag-lloyd.com/solutions" in cur or "new-quote" in cur:
                        print("   [HỆ THỐNG] ✅ Đã vào được trang tìm giá!")
                        break
                except:
                    pass
            else:
                print("   [HỆ THỐNG] ⚠️ Hết 120s chờ login, tiếp tục chạy...")
    else:
        print("   [HỆ THỐNG] Đã đăng nhập sẵn.")

print("""
╔══════════════════════════════════════════════╗
║   HPL Price Checker — 3 Tab Pipeline  🚢     ║
╚══════════════════════════════════════════════╝
""")

wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# --- Đọc Free Time từ freetime_hpl.xlsx ---
freetime_map = {}
try:
    ft_path = os.path.join(current_folder, "freetime_hpl.xlsx")
    if os.path.exists(ft_path):
        ft_wb = openpyxl.load_workbook(ft_path, read_only=True)
        ft_ws = ft_wb.active
        for ft_row in ft_ws.iter_rows(min_row=2, values_only=True):
            country = str(ft_row[0] or "").strip().upper()
            ft_val  = str(ft_row[1] or "").strip()
            if country and ft_val:
                freetime_map[country] = ft_val
        ft_wb.close()
        print(f"📋 Đọc được {len(freetime_map)} quốc gia từ freetime_hpl.xlsx")
    else:
        print("⚠️ Không tìm thấy freetime_hpl.xlsx — bỏ qua free time")
except Exception as e:
    print(f"⚠️ Lỗi đọc freetime_hpl.xlsx: {e}")

# --- Điền Free Time vào cột N (14) cho các dòng HAPAG LLOYD ---
if freetime_map:
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        carrier = str(row[4] or "").strip().upper()
        if carrier not in HPL_GROUP:
            continue
        country = str(row[1] or "").strip().upper()   # Cột B = country
        ft = freetime_map.get(country, "")
        if ft:
            ws.cell(row=i, column=14).value = ft
    try:
        wb.save(excel_path)
        print(f"✅ Đã điền free time cho các dòng HPL")
    except:
        pass

row_queue = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    pol_excel = str(row[2] or "").strip().upper()
    pod       = str(row[3] or "").strip().title()
    carrier   = str(row[4] or "").strip().upper()
    if not pol_excel or not pod: continue
    if carrier not in HPL_GROUP: continue
    if FILTER_POL and pol_excel != FILTER_POL: continue
    if FILTER_POD and pod.upper() != FILTER_POD: continue
    pol_search = POL_MAP.get(pol_excel, pol_excel.title())
    row_queue.append((i, pol_search, pod))

# Sort theo POL → cùng POL liên tiếp nhau → skip nhập lại POL, tiết kiệm thời gian
row_queue.sort(key=lambda x: x[1])
total = len(row_queue)
print(f"📋 Tổng cộng {total} dòng HPL (đã sort theo POL)")

# Gọi hàm xử lý đăng nhập ở đây
handle_login()

# ── CHECK TAB ĐÃ MỞ SẴN WEB HPL ──
STEALTH_SCRIPT = """
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
    Object.defineProperty(document, 'visibilityState', {get: () => 'visible'});
    Object.defineProperty(document, 'hidden', {get: () => false});
    const _origAdd = EventTarget.prototype.addEventListener;
    EventTarget.prototype.addEventListener = function(type, listener, options) {
        if (type === 'visibilitychange' || type === 'blur' || type === 'pagehide') return;
        return _origAdd.call(this, type, listener, options);
    };
"""

tabs = []
print("🔍 Đang kiểm tra các tab đã mở sẵn...")
all_handles = driver.window_handles
for h in all_handles:
    if len(tabs) >= NUM_TABS:
        break
    try:
        driver.switch_to.window(h)
        url = driver.current_url or ""
        if "hapag-lloyd.com" in url and "new-quote" in url:
            # Tab đã mở sẵn trang tìm giá HPL
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
                )
                tabs.append(h)
                print(f"   ✅ Tái sử dụng tab có sẵn: {h[:8]}... (URL: {url[:60]})")
            except:
                # Tab mở HPL nhưng chưa sẵn sàng → reload thử
                print(f"   🔄 Tab HPL chưa sẵn sàng, đang reload: {h[:8]}...")
                driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": STEALTH_SCRIPT})
                driver.get(BASE_URL)
                try:
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
                    )
                    tabs.append(h)
                    print(f"   ✅ Tab reload thành công: {h[:8]}...")
                except:
                    print(f"   ❌ Tab không load được, bỏ qua: {h[:8]}...")
    except:
        pass

if tabs:
    print(f"♻️  Tìm thấy {len(tabs)} tab HPL có sẵn, tái sử dụng!")

# Chỉ mở thêm tab mới nếu chưa đủ NUM_TABS
need_more = NUM_TABS - len(tabs)
if need_more > 0:
    print(f"🌐 Cần mở thêm {need_more} tab mới...")
    for i in range(need_more):
        driver.switch_to.new_window('tab')
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": STEALTH_SCRIPT})
        rand_sleep(0.5, 0.8)

        tab_num = len(tabs) + 1
        print(f"   🌐 Tab{tab_num} load HPL...")
        for attempt in range(2):
            try:
                driver.get(BASE_URL)
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="start-input"]'))
                )
                break
            except Exception as e:
                if attempt == 0:
                    print(f"   ⚠️ Tab{tab_num} load chậm hoặc bị kẹt, đang thử reload...")
                else:
                    raise Exception(f"Tab{tab_num} không thể load được trang tìm giá.")

        tabs.append(driver.current_window_handle)
        rand_sleep(0.8, 1.2)
else:
    print(f"♻️  Đã đủ {NUM_TABS} tab HPL có sẵn, không cần mở thêm. Tiết kiệm RAM!")

print(f"🌐 {NUM_TABS} tabs sẵn sàng: {[h[:8] for h in tabs]}\n")

# ===================================================================================
# PIPELINE
# ===================================================================================
queue_idx = 0
pending   = [None] * NUM_TABS

def search_job(t, job):
    driver.switch_to.window(tabs[t])
    activate_tab(tabs[t])
    print(f"\n{'='*50}")
    print(f"[Tab{t+1}] Dòng {job[0]}: {job[1]} → {job[2]}")
    try:
        do_search(job[1], job[2], t+1)
        return True
    except Exception as e:
        msg = str(e)
        val = "NO SERVICE" if "Thất bại tất cả aliases" in msg else f"LỖI: {msg}"
        print(f"[Tab{t+1}] ❌ Search lỗi: {msg}")
        ws.cell(row=job[0], column=6).value = val
        try: wb.save(excel_path)
        except: pass
        try:
            print(f"[Tab{t+1}] 🔄 Reload tab sau lỗi...")
            reload_tab()
            # Reset cache sau reload vì form đã reset
            tab_last_pol[t] = None
            tab_last_pod[t] = None
        except: pass
        return False

# --- Bước 1: Search 3 job đầu ---
import time as _timer
_pipeline_start = _timer.time()
print("🚀 Bước 1: Search 3 tab đầu tiên...")
for t in range(NUM_TABS):
    if queue_idx >= total:
        break
    job = row_queue[queue_idx]; queue_idx += 1
    ok  = search_job(t, job)
    pending[t] = job if ok else None
    while not ok and queue_idx < total:
        job = row_queue[queue_idx]; queue_idx += 1
        ok  = search_job(t, job)
        pending[t] = job if ok else None

# --- Bước 2: Pipeline ---
print(f"\n🔄 Bước 2: Pipeline...")
while any(p is not None for p in pending):
    for t in range(NUM_TABS):
        job = pending[t]
        if job is None:
            continue

        driver.switch_to.window(tabs[t])
        activate_tab(tabs[t])
        print(f"\n[Tab{t+1}] 💰 Lấy giá dòng {job[0]}...")
        get_price_and_save(job[0], t+1, wb, ws, job_pod=job[2])

        print(f"[Tab{t+1}] ⬅️ Back...")
        go_back()

        if queue_idx < total:
            new_job = row_queue[queue_idx]; queue_idx += 1
            ok = search_job(t, new_job)
            pending[t] = new_job if ok else None
            while not ok and queue_idx < total:
                new_job = row_queue[queue_idx]; queue_idx += 1
                ok = search_job(t, new_job)
                pending[t] = new_job if ok else None
        else:
            pending[t] = None
            print(f"[Tab{t+1}] ✅ Hết việc")

elapsed = _timer.time() - _pipeline_start
avg = elapsed / total if total > 0 else 0
print(f"\n✅ Hoàn tất {total} dòng trong {elapsed:.0f}s (trung bình {avg:.1f}s/dòng)")
"""
MAIN ORCHESTRATOR — Điều khiển 8 bot check giá hãng tàu
=========================================================
Cách dùng:
  python main.py                            → chạy FULL Excel, 8 bot song song
  python main.py "HO CHI MINH" "CHENNAI"   → 8 bot cùng check 1 tuyến
  python main.py 145                        → chỉ chạy bot của carrier ở dòng 145

Cách hoạt động:
  - Mỗi bot chạy trong subprocess riêng
  - Mỗi bot ghi vào BẢN COPY Excel riêng (input_gia_CMA.xlsx, input_gia_HPL.xlsx, ...)
  - Sau khi TẤT CẢ bot xong → main.py gộp kết quả từ các bản copy vào file gốc
  - Tránh corrupt do nhiều process ghi cùng 1 file xlsx
"""

import sys
import os
import subprocess
import time
import shutil
import threading
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("[ERROR] Cần cài openpyxl: pip install openpyxl")
    sys.exit(1)

# ===================================================================================
# CONFIG
# ===================================================================================
current_folder = os.getcwd()
excel_path     = os.path.join(current_folder, "input_gia.xlsx")
PYTHON_EXE     = sys.executable

# Carrier → bot file mapping
CARRIER_TO_BOT = {
    "CMA":          "bot_cma.py",
    "ANL":          "bot_cma.py",
    "CNC":          "bot_cma.py",
    "APL":          "bot_cma.py",
    "COSCO":        "bot_COSCO.py",
    "EMC":          "bot_EMC.py",
    "EVERGREEN":    "bot_EMC.py",
    "HPL":          "bot_HPL.py",
    "HAPAG":        "bot_HPL.py",
    "HAPAG-LLOYD":  "bot_HPL.py",
    "HAPAG LLOYD":  "bot_HPL.py",
    "KMTC":         "bot_KMTC.py",
    "MSC":          "bot_msc.py",
    "ONE":          "bot_one.py",
    "OOCL":         "bot_oocl.py",
}

# Bot short names (for display & temp file naming)
BOT_DISPLAY = {
    "bot_cma.py":   "CMA",
    "bot_COSCO.py": "COSCO",
    "bot_EMC.py":   "EMC",
    "bot_HPL.py":   "HAPAG LLOYD",
    "bot_KMTC.py":  "KMTC",
    "bot_msc.py":   "MSC",
    "bot_one.py":   "ONE",
    "bot_oocl.py":  "OOCL",
}

# Carrier groups — để biết bot nào xử lý carrier nào
BOT_CARRIERS = {
    "bot_cma.py":   {"CMA", "ANL", "CNC", "APL"},
    "bot_COSCO.py": {"COSCO"},
    "bot_EMC.py":   {"EMC", "EVERGREEN"},
    "bot_HPL.py":   {"HPL", "HAPAG", "HAPAG-LLOYD", "HAPAG LLOYD"},
    "bot_KMTC.py":  {"KMTC"},
    "bot_msc.py":   {"MSC"},
    "bot_one.py":   {"ONE"},
    "bot_oocl.py":  {"OOCL"},
}

# Columns kết quả mà bot ghi: F(6) G(7) H(8) I(9) J(10) K(11) M(13) N(14) O(15) P(16)
RESULT_COLS = [6, 7, 8, 9, 10, 11, 13, 14, 15, 16]

# 2 phase chạy — tiết kiệm RAM, mỗi phase 4 bot
PHASE_1 = {"bot_cma.py", "bot_COSCO.py","bot_one.py", "bot_oocl.py"}
PHASE_2 = {"bot_EMC.py", "bot_msc.py","bot_KMTC.py" , "bot_HPL.py"}


# ===================================================================================
# HELPERS
# ===================================================================================
def ts():
    return datetime.now().strftime("%H:%M:%S")

def print_header():
    print(f"""
╔═══════════════════════════════════════════════════╗
║   MAIN ORCHESTRATOR — 8 Bot Price Checker  🚀     ║
╚═══════════════════════════════════════════════════╝
[{ts()}] Thư mục: {current_folder}
[{ts()}] Excel:   {excel_path}
""")

def check_excel():
    if not os.path.exists(excel_path):
        print(f"[{ts()}] ❌ Không tìm thấy {excel_path}")
        sys.exit(1)

def get_all_carriers():
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb.active
    carriers = set()
    for row in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        c = str(row[4] or "").strip().upper()
        if c:
            carriers.add(c)
    wb.close()
    return carriers

def get_bots_for_carriers(carriers):
    bots = set()
    for c in carriers:
        bot = CARRIER_TO_BOT.get(c)
        if bot:
            bots.add(bot)
        else:
            print(f"[{ts()}] ⚠️ Carrier '{c}' không có bot tương ứng")
    return bots

def make_bot_copy(bot_file):
    """
    Tạo bản copy Excel riêng cho mỗi bot.
    Ví dụ: input_gia_HPL.xlsx
    Bot sẽ đọc/ghi vào bản copy này thay vì file gốc.
    """
    name = BOT_DISPLAY.get(bot_file, bot_file.replace(".py", ""))
    copy_path = excel_path.replace(".xlsx", f"_{name}.xlsx")
    shutil.copy2(excel_path, copy_path)
    return copy_path

def run_bot(bot_file, excel_override=None, extra_env=None):
    """
    Chạy bot trong subprocess.
    Nếu excel_override, set env EXCEL_PATH để bot dùng file khác.
    """
    bot_path = os.path.join(current_folder, bot_file)
    if not os.path.exists(bot_path):
        print(f"[{ts()}] ❌ Không tìm thấy: {bot_file}")
        return None

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    if excel_override:
        env["EXCEL_PATH"] = excel_override
    if extra_env:
        env.update(extra_env)

    proc = subprocess.Popen(
        [PYTHON_EXE, bot_path],
        env=env,
        cwd=current_folder,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return proc

def stream_output(name, proc):
    """Đọc stdout của process và in với prefix [BOT_NAME]"""
    for line in proc.stdout:
        line = line.rstrip("\n\r")
        if line:
            print(f"[{ts()}] [{name}] {line}")
    # Chờ process exit tối đa 10s sau khi stdout đóng
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        print(f"[{ts()}] [{name}] ⚠️ Process không exit sau 10s — force kill")
        proc.kill()
        proc.wait()
    rc = proc.returncode
    status = "✅ OK" if rc == 0 else f"❌ exit={rc}"
    print(f"[{ts()}] [{name}] {status}")

# Không giới hạn timeout — chờ bot tự xong

def run_bots_parallel(bots_to_run, use_copies=True, extra_env=None):
    """
    Chạy nhiều bot song song. Mỗi bot ghi vào bản copy xlsx riêng.
    Trả về dict {bot_file: copy_path}
    extra_env: dict env vars bổ sung (VD: FILTER_POL, FILTER_POD)
    """
    procs = {}   # name → (proc, thread)
    copies = {}  # bot_file → copy_path

    for bot_file in sorted(bots_to_run):
        name = BOT_DISPLAY.get(bot_file, bot_file)

        if use_copies:
            copy_path = make_bot_copy(bot_file)
            copies[bot_file] = copy_path
            print(f"[{ts()}] 🚀 {name} → {os.path.basename(copy_path)}")
            proc = run_bot(bot_file, excel_override=copy_path, extra_env=extra_env)
        else:
            print(f"[{ts()}] 🚀 {name}")
            proc = run_bot(bot_file, extra_env=extra_env)

        if proc:
            t = threading.Thread(target=stream_output, args=(name, proc), daemon=True)
            t.start()
            procs[name] = (proc, t)
            time.sleep(2)  # stagger 2s giữa các bot

    # Chờ tất cả kết thúc — không giới hạn thời gian
    print(f"\n[{ts()}] ⏳ Đang chờ {len(procs)} bot hoàn tất...\n")
    for name, (proc, thread) in procs.items():
        thread.join()  # chờ đến khi bot tự xong

    return copies

def merge_copies_to_main(copies, filter_pol=None, filter_pod=None, filter_row=None):
    """
    Đọc kết quả từ các bản copy, ghi vào file gốc.
    Chỉ ghi các cột kết quả (F-P) của dòng mà bot xử lý (dựa vào carrier).

    filter_pol/filter_pod: chỉ gộp dòng có POL/POD khớp (cho Mode 2)
    filter_row: chỉ gộp đúng 1 dòng (cho Mode 3)
    """
    if not copies:
        return

    if filter_row:
        print(f"[{ts()}] 📊 Gộp kết quả dòng {filter_row} vào file gốc...")
    elif filter_pol and filter_pod:
        print(f"[{ts()}] 📊 Gộp kết quả tuyến {filter_pol} → {filter_pod} vào file gốc...")
    else:
        print(f"[{ts()}] 📊 Gộp kết quả từ {len(copies)} bản copy vào file gốc...")

    wb_main = openpyxl.load_workbook(excel_path)
    ws_main = wb_main.active

    total_merged = 0
    for bot_file, copy_path in copies.items():
        if not os.path.exists(copy_path):
            continue

        name = BOT_DISPLAY.get(bot_file, bot_file)
        carrier_set = BOT_CARRIERS.get(bot_file, set())

        try:
            wb_copy = openpyxl.load_workbook(copy_path, read_only=True)
            ws_copy = wb_copy.active

            count = 0
            for r in range(2, ws_copy.max_row + 1):
                # Filter theo row number (Mode 3)
                if filter_row and r != filter_row:
                    continue

                carrier = str(ws_copy.cell(row=r, column=5).value or "").strip().upper()
                if carrier not in carrier_set:
                    continue

                # Filter theo POL/POD (Mode 2)
                if filter_pol and filter_pod:
                    r_pol = str(ws_copy.cell(row=r, column=3).value or "").strip().upper()
                    r_pod = str(ws_copy.cell(row=r, column=4).value or "").strip().upper()
                    if r_pol != filter_pol or r_pod != filter_pod:
                        continue

                # Check xem có kết quả mới không (cột F không trống)
                val_f = ws_copy.cell(row=r, column=6).value
                if val_f is None:
                    continue

                # Copy các cột kết quả
                for col in RESULT_COLS:
                    val = ws_copy.cell(row=r, column=col).value
                    if val is not None:
                        ws_main.cell(row=r, column=col).value = val

                count += 1

            wb_copy.close()
            total_merged += count
            print(f"[{ts()}]    [{name}] Gộp {count} dòng")

        except Exception as e:
            print(f"[{ts()}]    [{name}] ❌ Lỗi đọc copy: {e}")

        # Xóa bản copy
        try:
            os.remove(copy_path)
        except:
            pass

    try:
        wb_main.save(excel_path)
        print(f"[{ts()}] ✅ Đã gộp {total_merged} dòng vào {os.path.basename(excel_path)}")
    except PermissionError:
        print(f"[{ts()}] ❌ Không ghi được — đóng file Excel rồi thử lại!")
        # Lưu tạm
        fallback = excel_path.replace(".xlsx", "_merged.xlsx")
        wb_main.save(fallback)
        print(f"[{ts()}] 💾 Đã lưu tạm: {os.path.basename(fallback)}")
    wb_main.close()

# ===================================================================================
# MODE 1: python main.py → chạy FULL Excel
# ===================================================================================
def mode_full():
    print(f"[{ts()}] 📋 Chế độ: FULL — chạy tất cả dòng trong Excel (2 phase)")
    check_excel()

    carriers = get_all_carriers()
    bots = get_bots_for_carriers(carriers)

    if not bots:
        print(f"[{ts()}] ⚠️ Không có carrier nào cần chạy")
        return

    print(f"[{ts()}] 📊 Carriers: {', '.join(sorted(carriers))}")
    print(f"[{ts()}] 🤖 Bots: {', '.join(BOT_DISPLAY.get(b, b) for b in sorted(bots))}")

    bots_p1 = bots & PHASE_1
    bots_p2 = bots & PHASE_2

    all_copies = {}

    # ── PHASE 1: CMA, COSCO, HPL, KMTC ──
    if bots_p1:
        p1_names = ', '.join(BOT_DISPLAY.get(b, b) for b in sorted(bots_p1))
        print(f"\n{'='*55}")
        print(f"[{ts()}] 🔵 PHASE 1/2: {p1_names}")
        print(f"{'='*55}")
        copies1 = run_bots_parallel(bots_p1)
        all_copies.update(copies1)
        print(f"\n[{ts()}] ✅ Phase 1 xong!")
    else:
        print(f"\n[{ts()}] Phase 1: không có bot nào cần chạy")

    # ── PHASE 2: ONE, OOCL, EMC, MSC ──
    if bots_p2:
        p2_names = ', '.join(BOT_DISPLAY.get(b, b) for b in sorted(bots_p2))
        print(f"\n{'='*55}")
        print(f"[{ts()}] 🟢 PHASE 2/2: {p2_names}")
        print(f"{'='*55}")
        copies2 = run_bots_parallel(bots_p2)
        all_copies.update(copies2)
        print(f"\n[{ts()}] ✅ Phase 2 xong!")
    else:
        print(f"\n[{ts()}] Phase 2: không có bot nào cần chạy")

    # ── GỘP KẾT QUẢ ──
    print()
    merge_copies_to_main(all_copies)
    print(f"\n[{ts()}] 🎉 HOÀN TẤT!")

# ===================================================================================
# MODE 2: python main.py "HO CHI MINH" "CHENNAI" ["ĐẤT NƯỚC"] → 8 bot check 1 tuyến
# ===================================================================================
def mode_route(pol, pod, country=None):
    pol_upper = pol.strip().upper()
    pod_title = pod.strip().title()
    pod_upper = pod.strip().upper()
    country_title = country.strip().title() if country else ""
    country_upper = country.strip().upper() if country else ""

    if country_title:
        print(f"[{ts()}] 🗺️  Chế độ: ROUTE — {pol_upper} → {pod_title} ({country_title})")
    else:
        print(f"[{ts()}] 🗺️  Chế độ: ROUTE — {pol_upper} → {pod_title}")
    check_excel()

    # Đảm bảo mỗi carrier chính có 1 row cho tuyến này
    main_carriers = ["CMA", "COSCO", "EMC", "HPL", "KMTC", "MSC", "ONE", "OOCL"]

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Scan existing rows + cập nhật country nếu cần
    existing_carriers = set()
    country_updated = 0
    for r in range(2, ws.max_row + 1):
        r_pol     = str(ws.cell(row=r, column=3).value or "").strip().upper()
        r_pod     = str(ws.cell(row=r, column=4).value or "").strip().upper()
        r_carrier = str(ws.cell(row=r, column=5).value or "").strip().upper()
        if r_pol == pol_upper and r_pod == pod_upper and r_carrier:
            existing_carriers.add(r_carrier)
            # Nếu có country mà cột B trống → điền country
            if country_title:
                existing_country = str(ws.cell(row=r, column=2).value or "").strip()
                if not existing_country:
                    ws.cell(row=r, column=2).value = country_title
                    country_updated += 1
    if country_updated > 0:
        print(f"[{ts()}] 📝 Đã cập nhật country '{country_title}' cho {country_updated} dòng có sẵn")

    # Thêm row mới cho carrier chưa có
    next_row = ws.max_row + 1
    added = 0
    for carrier in main_carriers:
        if carrier not in existing_carriers:
            if country_title:
                ws.cell(row=next_row, column=2).value = country_title
            ws.cell(row=next_row, column=3).value = pol_upper
            ws.cell(row=next_row, column=4).value = pod_title
            ws.cell(row=next_row, column=5).value = carrier
            print(f"[{ts()}] ➕ Thêm dòng {next_row}: {pol_upper} → {pod_title} [{carrier}]")
            next_row += 1
            added += 1

    if added > 0 or country_updated > 0:
        try:
            wb.save(excel_path)
            if added > 0:
                print(f"[{ts()}] 💾 Đã thêm {added} dòng mới")
            if added == 0:
                print(f"[{ts()}] ✅ Tất cả carrier đã có row cho tuyến này")
        except PermissionError:
            print(f"[{ts()}] ❌ Đóng file Excel trước khi chạy!")
            sys.exit(1)
    else:
        print(f"[{ts()}] ✅ Tất cả carrier đã có row cho tuyến này")
    wb.close()

    # Truyền filter POL/POD/COUNTRY cho các bot → bot chỉ check tuyến này
    route_env = {"FILTER_POL": pol_upper, "FILTER_POD": pod_upper}
    if country_upper:
        route_env["FILTER_COUNTRY"] = country_upper

    # Chạy tất cả bot theo 2 phase
    bots = set(CARRIER_TO_BOT[c] for c in main_carriers)
    bots_p1 = bots & PHASE_1
    bots_p2 = bots & PHASE_2
    all_copies = {}

    if bots_p1:
        p1_names = ', '.join(BOT_DISPLAY.get(b, b) for b in sorted(bots_p1))
        print(f"\n[{ts()}] 🔵 PHASE 1/2: {p1_names}")
        copies1 = run_bots_parallel(bots_p1, extra_env=route_env)
        all_copies.update(copies1)

    if bots_p2:
        p2_names = ', '.join(BOT_DISPLAY.get(b, b) for b in sorted(bots_p2))
        print(f"\n[{ts()}] 🟢 PHASE 2/2: {p2_names}")
        copies2 = run_bots_parallel(bots_p2, extra_env=route_env)
        all_copies.update(copies2)

    print()
    merge_copies_to_main(all_copies, filter_pol=pol_upper, filter_pod=pod_upper)
    print(f"\n[{ts()}] 🎉 HOÀN TẤT tuyến {pol_upper} → {pod_title}!")

# ===================================================================================
# MODE 3: python main.py 145 → chỉ chạy bot của carrier ở dòng 145
# ===================================================================================
def mode_single_row(row_num):
    print(f"[{ts()}] 🎯 Chế độ: SINGLE ROW — dòng {row_num}")

    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb.active
    max_row = ws.max_row

    if row_num < 2 or row_num > max_row:
        print(f"[{ts()}] ❌ Dòng {row_num} ngoài phạm vi (2-{max_row})")
        wb.close()
        return

    pol     = str(ws.cell(row=row_num, column=3).value or "").strip()
    pod     = str(ws.cell(row=row_num, column=4).value or "").strip()
    carrier = str(ws.cell(row=row_num, column=5).value or "").strip().upper()
    wb.close()

    if not carrier:
        print(f"[{ts()}] ❌ Dòng {row_num} không có carrier (cột E trống)")
        return
    if not pol or not pod:
        print(f"[{ts()}] ❌ Dòng {row_num} thiếu POL hoặc POD")
        return

    bot_file = CARRIER_TO_BOT.get(carrier)
    if not bot_file:
        print(f"[{ts()}] ❌ Carrier '{carrier}' không có bot")
        return

    name = BOT_DISPLAY.get(bot_file, bot_file)
    print(f"[{ts()}] 📋 Dòng {row_num}: {pol} → {pod} [{carrier}] → {name}")

    # Dùng bản copy để an toàn, sau đó chỉ gộp đúng dòng row_num
    copy_path = make_bot_copy(bot_file)
    print(f"[{ts()}] 🚀 {name} → {os.path.basename(copy_path)}")
    proc = run_bot(bot_file, excel_override=copy_path, extra_env={"SINGLE_ROW": str(row_num)})
    if not proc:
        return

    print(f"\n[{ts()}] ⏳ Đang chờ {name}...\n")

    # Stream output
    t = threading.Thread(target=stream_output, args=(name, proc), daemon=True)
    t.start()
    t.join()  # chờ đến khi bot tự xong

    # Gộp chỉ dòng row_num
    print()
    merge_copies_to_main({bot_file: copy_path}, filter_row=row_num)
    print(f"\n[{ts()}] 🎉 HOÀN TẤT dòng {row_num}!")

# ===================================================================================
# ENTRY POINT
# ===================================================================================
if __name__ == "__main__":
    print_header()

    args = sys.argv[1:]

    if len(args) == 0:
        mode_full()

    elif len(args) == 1:
        try:
            row_num = int(args[0])
            mode_single_row(row_num)
        except ValueError:
            print(f"[{ts()}] ❌ Tham số không hợp lệ: '{args[0]}'")
            print(f"  Dùng: python main.py 145           (chạy 1 dòng)")
            print(f"  Hoặc: python main.py \"POL\" \"POD\"   (chạy 1 tuyến)")
            sys.exit(1)

    elif len(args) == 2:
        pol, pod = args[0], args[1]
        mode_route(pol, pod)

    elif len(args) == 3:
        pol, pod, country = args[0], args[1], args[2]
        mode_route(pol, pod, country=country)

    else:
        print(f"[{ts()}] ❌ Sai cú pháp!")
        print(f"  python main.py                                    → chạy full Excel")
        print(f"  python main.py \"HO CHI MINH\" \"CHENNAI\"            → check 1 tuyến")
        print(f"  python main.py \"HO CHI MINH\" \"HAMBURG\" \"GERMANY\"  → check 1 tuyến + quốc gia")
        print(f"  python main.py 145                                → check 1 dòng")
        sys.exit(1)

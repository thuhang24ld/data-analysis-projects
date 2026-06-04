"""
iPos Fabi - Auto Crawler (Sale Detail)
=======================================
- Tự động login lấy token mới mỗi ngày
- Crawl sale-detail của ngày hôm qua
- Lưu ra file Excel theo ngày

Cài: pip install requests pandas openpyxl
Chạy: python crawl_sale_detail.py
"""

import requests
import pandas as pd
import json, os, sys, base64, logging, time, socket
from datetime import datetime, timezone, timedelta, date

# ============================================================
# CẤU HÌNH
# ============================================================

IPOS_EMAIL    = "username"
IPOS_PASSWORD = "password"   # <-- điền password


COMPANY_UID    = "COMPANY_UID"
BRAND_UID      = "BRAND_UID"
STORE_UID      = "STORE_UID"
LIST_STORE_UID = "LIST_STORE_UID"
ACCESS_TOKEN   = "ACCESS_TOKEN"

OUTPUT_DIR       = r"D:\Rooftop_IPos_data\sale_detail"
TOKEN_CACHE_FILE = r"D:\Rooftop_IPos_data\.token_cache.json"  # dùng chung cache với crawler chính

LOGIN_URL  = "https://posapi.ipos.vn/api/accounts/v1/user/login"
DETAIL_URL = "https://posapi.ipos.vn/api/v1/accounting/report/sale-detail"
TZ_VN      = timezone(timedelta(hours=7))

# ============================================================
# LOGGING
# ============================================================
 
os.makedirs(OUTPUT_DIR, exist_ok=True)
 
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(OUTPUT_DIR, "crawler_detail.log"), encoding="utf-8"),
    ]
)
log = logging.getLogger(__name__)
 
# ============================================================
# TOKEN
# ============================================================
 
def load_cached_token():
    try:
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache = json.load(f)
        if cache.get("exp", 0) - datetime.now().timestamp() > 7200:
            log.info("✅ Dùng token cache (còn hạn)")
            return cache["token"]
    except Exception:
        pass
    return None
 
def save_token(token, exp):
    os.makedirs(os.path.dirname(TOKEN_CACHE_FILE), exist_ok=True)
    with open(TOKEN_CACHE_FILE, "w") as f:
        json.dump({"token": token, "exp": exp}, f)
 
def login():
    log.info("🔐 Đang login iPos...")
 
    resp = requests.post(
        LOGIN_URL,
        json={"email": IPOS_EMAIL, "password": IPOS_PASSWORD},
        headers={
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "vi",
            "Access_token": ACCESS_TOKEN,
            "Authorization": "",
            "Content-Type": "application/json;charset=UTF-8",
            "Fabi_type": "pos-cms",
            "Origin": "https://fabi.ipos.vn",
            "Referer": "https://fabi.ipos.vn/",
            "X-Client-Timezone": "25200000",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        },
        timeout=15
    )
 
    if not resp.ok:
        raise Exception(f"Login thất bại: HTTP {resp.status_code} - {resp.text[:300]}")
 
    data = resp.json()
    token = (
        data.get("token") or
        data.get("access_token") or
        (data.get("data") or {}).get("token") or
        (data.get("data") or {}).get("access_token")
    )
    if not token:
        raise Exception(f"Không tìm thấy token: {str(data)[:300]}")
 
    try:
        pad = token.split(".")[1]
        pad += "=" * (4 - len(pad) % 4)
        exp = json.loads(base64.b64decode(pad)).get("exp", int(datetime.now().timestamp()) + 86400)
    except Exception:
        exp = int(datetime.now().timestamp()) + 86400
 
    save_token(token, exp)
    log.info(f"✅ Login OK! Token hết hạn: {datetime.fromtimestamp(exp, tz=TZ_VN)}")
    return token
 
def get_token():
    return load_cached_token() or login()
 
def make_headers(token):
    return {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "vi",
        "Access_token": ACCESS_TOKEN,
        "Authorization": token,
        "Fabi_type": "pos-cms",
        "Origin": "https://fabi.ipos.vn",
        "Referer": "https://fabi.ipos.vn/",
        "X-Client-Timezone": "25200000",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    }
 
# ============================================================
# CRAWLER
# ============================================================
 
def day_range_ms(d):
    start = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=TZ_VN)
    end   = datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=TZ_VN)
    return int(start.timestamp() * 1000), int(end.timestamp() * 1000)
 
def crawl_one_day(d, token):
    records   = []
    cursor    = None
    batch_num = 1
    start_ms, end_ms = day_range_ms(d)
 
    while True:
        params = {
            "company_uid":      COMPANY_UID,
            "brand_uid":        BRAND_UID,
            "list_store_uid":   LIST_STORE_UID,
            "start_date":       start_ms,
            "end_date":         end_ms,
            "store_open_at":    2,
            "results_per_page": 1000,  # lấy nhiều nhất có thể mỗi lần
        }
        if cursor:
            params["next_cursor"] = cursor
 
        resp = requests.get(
            DETAIL_URL,
            headers=make_headers(token),
            params=params,
            timeout=120
        )
 
        if resp.status_code == 401:
            log.warning("Token hết hạn, login lại...")
            token = login()
            continue
 
        if not resp.ok:
            log.warning(f"HTTP {resp.status_code}: {resp.text[:100]}")
            break
 
        data  = resp.json()
        batch = data.get("data", [])
 
        if not batch:
            break
 
        for r in batch:
            r["_date"] = str(d)
        records.extend(batch)
 
        new_cursor = data.get("next_cursor")
        log.info(f"   Batch {batch_num}: +{len(batch)} records (tổng: {len(records)}, cursor: {new_cursor})")
 
        # Dừng nếu không có cursor mới hoặc cursor không đổi
        if not new_cursor or new_cursor == cursor:
            break
 
        cursor    = new_cursor
        batch_num += 1
        time.sleep(0.3)
 
    return records
 
def export_excel(records, d):
    filename = os.path.join(OUTPUT_DIR, f"sale_detail_{d}.xlsx")
    if not records:
        log.warning(f"Không có data cho ngày {d}")
        return
 
    df = pd.DataFrame(records)
    for col in df.columns:
        if any(k in col.lower() for k in ["date", "time", "created", "updated"]):
            try:
                df[col] = pd.to_datetime(df[col], unit="s", errors="coerce")
            except Exception:
                pass
 
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Data", index=False)
 
    log.info(f"💾 Saved: {filename} ({len(df)} rows)")
    print(json.dumps({"status": "success", "date": str(d), "rows": len(df), "file": filename}))
 
# ============================================================
# MAIN
# ============================================================
 
def wait_for_network(max_wait=60):
    log.info("⏳ Chờ kết nối mạng...")
    for i in range(max_wait):
        try:
            socket.setdefaulttimeout(3)
            socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("8.8.8.8", 53))
            log.info(f"✅ Mạng đã kết nối (sau {i}s)")
            return True
        except Exception:
            time.sleep(1)
    log.error("❌ Không có mạng sau 60 giây!")
    return False
 
if __name__ == "__main__":
    target = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 \
             else (datetime.now(tz=TZ_VN) - timedelta(days=1)).date()
 
    log.info(f"🚀 Crawl sale detail ngày {target}")
    try:
        if not wait_for_network():
            sys.exit(1)
        token   = get_token()
        records = crawl_one_day(target, token)
        log.info(f"📊 {len(records)} records")
        export_excel(records, target)
    except Exception as e:
        log.error(f"❌ {e}")
        print(json.dumps({"status": "error", "message": str(e)}))
        sys.exit(1)
 
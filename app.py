"""
客人集運預報系統
FWT JAPAN 雲端集運
"""

from flask import Flask, request, jsonify, render_template, make_response, send_file, session
from datetime import datetime, timedelta
import requests
import json
import os


def _load_dotenv(path=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")):
    """本機開發用：讀同目錄的 .env（KEY=VALUE）。已存在的環境變數優先，部署平台設定不會被覆蓋。"""
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
    except FileNotFoundError:
        pass


_load_dotenv()
import sqlite3
import csv
import math
import tw_zip
import io
import time
import re
import secrets
import threading

# 廠商出貨檔案範本（Nigel / JpD…）
import vendors as vendor_templates

# PWA（manifest / service worker）
from pwa import register_pwa
from brand import BRAND, MEMBER_METAFIELD_KEY, ENABLE_AGENTS

# SQLite 每日自動備份 → Google Drive（快照/驗證/上傳；租約與排程在下方「自動備份」區）
import gdrive_backup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
# Session 設定（環境變數 SESSION_SECRET 沒設就用隨機值，每次重啟會失效但不會暴露 fallback）
app.secret_key = os.environ.get("SESSION_SECRET") or secrets.token_hex(32)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)


@app.before_request
def _agents_disabled_guard():
    """代理商模組關閉（ENABLE_AGENTS 未開）時，代理相關 API 一律 404。"""
    if not ENABLE_AGENTS and (request.path.startswith("/api/admin/agents")
                              or request.path.startswith("/api/agent/")):
        return jsonify({"success": False, "error": "代理商功能未啟用"}), 404


@app.context_processor
def _inject_brand():
    """模板裡可用 {{ brand.name }} 等（見 brand.py）"""
    return {"brand": BRAND, "enable_agents": ENABLE_AGENTS}

# 回應壓縮：admin.html / index.html 是 341 kB / 152 kB 的純靜態 HTML，
# gzip 後約剩 1/6，直接砍掉客戶端的下載時間（TTFB 本來就只有 15ms）。
# 缺套件時只是不壓縮，不會讓主程式起不來（沿用 recon 的容錯寫法）。
app.config["COMPRESS_MIMETYPES"] = [
    "text/html", "text/css", "text/plain", "text/javascript",
    "application/javascript", "application/json", "image/svg+xml",
]
app.config["COMPRESS_LEVEL"] = 6       # 1~9，6 是壓縮率與 CPU 的平衡點
app.config["COMPRESS_MIN_SIZE"] = 500  # 小於 500 bytes 不壓（壓了反而變大）
try:
    from flask_compress import Compress
    Compress(app)
    print("[App] ✅ 回應壓縮已啟用", flush=True)
except Exception as _compress_err:
    print(f"[App] ⚠️ 回應壓縮未啟用，將以未壓縮傳送: {_compress_err}", flush=True)

# PWA：註冊 /sw.js + /manifest.webmanifest + /admin-manifest.webmanifest 三條路由
register_pwa(app)


@app.after_request
def _no_store_html(resp):
    """HTML 頁面一律 no-store，避免瀏覽器用啟發式快取留住舊版前端。

    2026-09 曾有一台電腦的瀏覽器一直跑舊版 admin.html（8/31 上線的最低計費重量 2kg
    規則到 9 月中仍以 1kg 存檔），根因就是頁面沒有任何快取指示。
    只針對 text/html：/static/ 底下（含 offline.html）維持可快取，交給 SW cache-first；
    /api/ 是 JSON，SW 本來就 network-only，不動。"""
    if request.path.startswith("/static/"):
        return resp
    if (resp.mimetype or "").lower() == "text/html":
        resp.headers["Cache-Control"] = "no-store, must-revalidate"
        resp.headers["Pragma"] = "no-cache"
    return resp

# 對帳模組：/recon 上傳銷帳檔比對帳單（限老闆）。DB_PATH 讓 recon 沿用同一個資料庫
app.config["DB_PATH"] = os.environ.get("DB_PATH", "packages.db")
try:
    from recon.routes import bp as recon_bp
    app.register_blueprint(recon_bp)
except Exception as _recon_err:
    print(f"[recon] 對帳模組載入失敗（不影響主程式）: {_recon_err}", flush=True)

# ============ 設定區（從環境變數讀取）============
JPD_BASE_URL = "https://biz.cloudwh.jp"
JPD_EMAIL = os.environ.get("JPD_EMAIL", "")
JPD_PASSWORD = os.environ.get("JPD_PASSWORD", "")
JPD_WAREHOUSE_ID = int(os.environ.get("JPD_WAREHOUSE_ID", "1"))
JPD_DELIV_ID = int(os.environ.get("JPD_DELIV_ID", "40"))  # 台灣空運線

SHOPIFY_STORE = os.environ.get("SHOPIFY_STORE", "")
SHOPIFY_ACCESS_TOKEN = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")

# 預設運費（台幣/kg），0 表示未設定
DEFAULT_SHIPPING_RATE = int(os.environ.get("DEFAULT_SHIPPING_RATE") or "180")

# 台幣 → 日圓匯率（可透過環境變數調整）
TWD_TO_JPY_RATE = float(os.environ.get("TWD_TO_JPY_RATE", "5.0"))

DB_PATH = os.environ.get("DB_PATH", "packages.db")
# ================================


# ============ SQLite 初始化 ============

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    # 等鎖最多 5 秒（預設 0 秒）→ 大幅減少「database is locked」錯誤、
    # 多 worker / LINE Bot / 客戶端同時操作時不互卡
    conn.execute("PRAGMA busy_timeout=5000")
    return conn


# ============ 加值服務預設目錄 ============
# 存進 admin_settings（key='extra_service_catalog'）後即可後台管理；此處僅為首次 seed。
# sel=True 才會出現在客戶端出貨申請的可勾選清單（變動價/特殊計費項目 sel=False，由管理員請款時確認）。
# FWT JAPAN 加值服務：只有易碎品包裝（日幣定價，依 TWD_TO_JPY_RATE 換成台幣出帳）
DEFAULT_EXTRA_SERVICES = [
    {"id": "es01", "name": "易碎品包裝", "cat": "加固",
     "desc": f"每個 ¥{BRAND['fragile_jpy']}（約 NT${round(BRAND['fragile_jpy'] / TWD_TO_JPY_RATE)}）",
     "price": round(BRAND["fragile_jpy"] / TWD_TO_JPY_RATE), "sel": True},
]


def get_extra_service_catalog(conn=None):
    """讀取加值服務目錄（admin_settings.extra_service_catalog）。缺則回預設。"""
    own = False
    if conn is None:
        conn = get_db(); own = True
    try:
        row = conn.execute("SELECT value FROM admin_settings WHERE key='extra_service_catalog'").fetchone()
    finally:
        if own:
            conn.close()
    if not row:
        return list(DEFAULT_EXTRA_SERVICES)
    try:
        data = json.loads(row["value"])
        return data if isinstance(data, list) else list(DEFAULT_EXTRA_SERVICES)
    except (ValueError, TypeError):
        return list(DEFAULT_EXTRA_SERVICES)


# ============ 台灣配送貨況（Google Sheet 同步）============
# 貨運行把台灣端派件資料填在這張試算表；系統定時抓 CSV 匯出、用「客戶編號」比對出貨單。
# 預設不帶任何試算表；在後台設定或用環境變數 TRACKING_SHEET_URL（CSV 匯出網址）
DEFAULT_TRACKING_SHEET_URL = os.environ.get("TRACKING_SHEET_URL", "")
_sync_lock = threading.Lock()


def _get_setting(key, default=""):
    conn = get_db()
    try:
        row = conn.execute("SELECT value FROM admin_settings WHERE key=?", (key,)).fetchone()
    finally:
        conn.close()
    return row["value"] if row else default


def _set_setting(key, value):
    conn = get_db()
    conn.execute(
        "INSERT INTO admin_settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value)
    )
    conn.commit()
    conn.close()


def delivery_tracking_url(carrier, tracking):
    """依物流商組查詢網址。"""
    t = (tracking or "").strip()
    c = (carrier or "").strip()
    if not t:
        return ""
    if "新竹" in c or "HCT" in c.upper():
        return f"https://www.aftership.com/zh-hant/track/hct-logistics/{t}"
    # 預設黑貓
    return f"https://www.t-cat.com.tw/Inquire/TraceDetail.aspx?BillID={t}"


def sync_delivery_tracking():
    """抓取貨運行試算表 CSV，解析後 upsert 進 delivery_tracking。回傳寫入筆數。"""
    url = _get_setting("tracking_sheet_url", DEFAULT_TRACKING_SHEET_URL)
    if not url:
        return 0  # 尚未設定貨況試算表
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    text = resp.content.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        return 0

    # 依標題找欄位（容忍欄位順序變動）；找不到就用固定位置 C=2 / F=5 / G=6
    header = [h.strip() for h in rows[0]]
    def _col(names, fallback):
        for i, h in enumerate(header):
            if any(n in h for n in names):
                return i
        return fallback
    ci_code = _col(["客戶編號", "客編"], 2)
    ci_track = _col(["派件轉單號", "轉單號", "追蹤"], 5)
    ci_carrier = _col(["貨態查詢", "物流", "物流商"], 6)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    count = 0
    for r in rows[1:]:
        if len(r) <= max(ci_code, ci_track, ci_carrier):
            continue
        code = (r[ci_code] or "").strip()
        tracking = (r[ci_track] or "").strip()
        carrier = (r[ci_carrier] or "").strip()
        if not code or not tracking:
            continue
        conn.execute(
            "INSERT INTO delivery_tracking (customer_code, carrier, tracking_num, synced_at) "
            "VALUES (?, ?, ?, ?) "
            "ON CONFLICT(customer_code) DO UPDATE SET "
            "carrier=excluded.carrier, tracking_num=excluded.tracking_num, synced_at=excluded.synced_at",
            (code, carrier, tracking, now)
        )
        count += 1
    conn.commit()
    conn.close()
    _set_setting("tracking_last_sync", now)
    return count


def maybe_auto_sync():
    """後台有人活動時，若距上次同步 > 24 小時就背景同步一次（不阻塞請求）。"""
    try:
        last = _get_setting("tracking_last_sync", "")
        if last:
            try:
                if (datetime.now() - datetime.strptime(last[:19], "%Y-%m-%d %H:%M:%S")).total_seconds() < 86400:
                    return
            except ValueError:
                pass
        if not _sync_lock.acquire(blocking=False):
            return
        # 先佔位，避免其他 worker/請求重複觸發
        _set_setting("tracking_last_sync", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        def _run():
            try:
                n = sync_delivery_tracking()
                print(f"[tracking] 自動同步完成，{n} 筆", flush=True)
            except Exception as e:
                print(f"[tracking] 自動同步失敗: {e}", flush=True)
            finally:
                _sync_lock.release()
        threading.Thread(target=_run, daemon=True).start()
    except Exception as e:
        print(f"[tracking] maybe_auto_sync 例外: {e}", flush=True)


def init_db():
    conn = get_db()
    # ===== 啟用 WAL 模式（一次性設定，會持久化在 DB 檔案內）=====
    # WAL：讀寫不互鎖、併發效能大幅提升（讀者不擋寫者、寫者不擋讀者）
    # synchronous=NORMAL：搭配 WAL 安全，速度比 FULL 快很多
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        mode_row = conn.execute("PRAGMA journal_mode").fetchone()
        print(f"[DB] ✅ SQLite journal_mode = {mode_row[0]}", flush=True)
    except Exception as e:
        print(f"[DB] ⚠️ 啟用 WAL 失敗（將使用預設模式）: {e}", flush=True)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS packages (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            logis_num   TEXT,
            product_name TEXT   DEFAULT '',
            weight      TEXT    DEFAULT '',
            status      TEXT    DEFAULT '已到貨',
            note        TEXT    DEFAULT '',
            in_date     TEXT,
            created_at  TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS shipment_requests (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            customer_name TEXT  DEFAULT '',
            package_ids TEXT    NOT NULL,
            package_summary TEXT DEFAULT '',
            status      TEXT    DEFAULT '待處理',
            note        TEXT    DEFAULT '',
            admin_note  TEXT    DEFAULT '',
            created_at  TEXT    NOT NULL,
            updated_at  TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS forecasts (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            customer_name TEXT  DEFAULT '',
            items_json  TEXT    NOT NULL,
            status      TEXT    DEFAULT '待處理',
            note        TEXT    DEFAULT '',
            created_at  TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS admin_settings (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS delivery_tracking (
            customer_code TEXT PRIMARY KEY,
            carrier       TEXT DEFAULT '',
            tracking_num  TEXT DEFAULT '',
            synced_at     TEXT DEFAULT ''
        )
    """)
    # ── 操作紀錄（誰做了什麼：到貨建立/出貨/帳單確認/認領）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS operation_logs (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            operator   TEXT DEFAULT '',
            role       TEXT DEFAULT '',
            action     TEXT DEFAULT '',
            target     TEXT DEFAULT '',
            detail     TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)
    # ── 無主包裹認領牆（沒客編/羅馬拼音的包裹，先登記保留到倉日，認領後轉入 packages）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS unclaimed_packages (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            recipient_name  TEXT DEFAULT '',
            logis_num       TEXT DEFAULT '',
            product_name    TEXT DEFAULT '',
            weight          TEXT DEFAULT '',
            note            TEXT DEFAULT '',
            registered_date TEXT,
            created_at      TEXT NOT NULL
        )
    """)
    # ── 無主包裹認領申請（會員在認領牆按「這是我的」→ 留紀錄等管理員確認，不直接轉入）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS unclaimed_claims (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            unclaimed_id INTEGER NOT NULL,
            g_code       TEXT NOT NULL,
            member_name  TEXT DEFAULT '',
            note         TEXT DEFAULT '',
            created_at   TEXT NOT NULL,
            UNIQUE(unclaimed_id, g_code)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_unclaimed_claims_uid ON unclaimed_claims(unclaimed_id)")
    # ── 停用會員名單（集運系統層級，不動 Shopify；g_code 為鍵）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS disabled_members (
            g_code       TEXT PRIMARY KEY,
            reason       TEXT DEFAULT '',
            disabled_at  TEXT DEFAULT ''
        )
    """)
    # ── 代理每週分潤撥款記錄（agent_id + period_key 唯一）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS agent_payouts (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            agent_id      INTEGER NOT NULL,
            period_key    TEXT NOT NULL,
            amount        REAL DEFAULT 0,
            payment_last5 TEXT DEFAULT '',
            paid_at       TEXT DEFAULT '',
            note          TEXT DEFAULT '',
            created_at    TEXT DEFAULT '',
            UNIQUE(agent_id, period_key)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS addresses (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            label       TEXT    DEFAULT '',
            recipient   TEXT    NOT NULL,
            phone       TEXT    NOT NULL,
            zipcode     TEXT    DEFAULT '',
            address     TEXT    NOT NULL,
            is_default  INTEGER DEFAULT 0,
            created_at  TEXT    NOT NULL
        )
    """)
    # ── 申報人（報單收貨人／納稅義務人）：會員層級 1..N，與地址簿無關 ──
    #    粒度是「箱」：台灣快遞進口一箱一份簡易申報單，一箱一位申報人。
    #    phone 必須是該人 EZ WAY 實名認證綁定的門號；依規定不收身分證字號。
    # ⚠️ 完整設計理由與法規依據見 docs/helpshipping-declarant-spec.md（§0 背景、§2 Schema）。
    #    「獨立一張表而非塞進 addresses」「存快照字串不存 declarant_id」「不收身分證字號」
    #    都是規格書寫死的約束，不可自行簡化。
    conn.execute("""
        CREATE TABLE IF NOT EXISTS declarants (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            name        TEXT    NOT NULL,
            phone       TEXT    NOT NULL,
            address     TEXT    NOT NULL,
            is_default  INTEGER DEFAULT 0,
            created_at  TEXT    NOT NULL
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_declarants_gcode ON declarants(g_code)")
    # ── 登入嘗試紀錄（/api/verify_customer 速率限制用）──
    #    密碼＝手機號碼、客編連號，沒有次數限制就能離線列舉。
    #    計數放 SQLite 而非記憶體：Procfile 是 gunicorn --workers 2，
    #    兩個 worker 不共享記憶體，用全域變數門檻會變兩倍且重啟即清空。
    #    created_ts 一律存 unix epoch 秒（容器時區 UTC，只做時間差，避開時區問題）。
    conn.execute("""
        CREATE TABLE IF NOT EXISTS login_attempts (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code     TEXT    DEFAULT '',
            ip         TEXT    DEFAULT '',
            success    INTEGER DEFAULT 0,
            created_ts INTEGER NOT NULL
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_login_attempts_ts ON login_attempts(created_ts)")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS announcements (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            title       TEXT    NOT NULL,
            content     TEXT    NOT NULL,
            is_active   INTEGER DEFAULT 1,
            created_at  TEXT    NOT NULL
        )
    """)
    # ── 內部公告（只給後台老闆/員工看，客戶端讀不到）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS internal_announcements (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            title       TEXT    NOT NULL,
            content     TEXT    NOT NULL,
            is_active   INTEGER DEFAULT 1,
            created_at  TEXT    NOT NULL
        )
    """)
    # ── 內部公告已讀記錄（誰按過「我知道了」）──
    conn.execute("""
        CREATE TABLE IF NOT EXISTS internal_ann_reads (
            username    TEXT NOT NULL,
            ann_id      INTEGER NOT NULL,
            read_at     TEXT NOT NULL,
            UNIQUE(username, ann_id)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS admin_users (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            username    TEXT    UNIQUE NOT NULL,
            password    TEXT    NOT NULL,
            role        TEXT    DEFAULT 'admin',
            created_at  TEXT    NOT NULL
        )
    """)
    # ===== 代理帳號表（Phase 1）=====
    conn.execute("""
        CREATE TABLE IF NOT EXISTS agents (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            username        TEXT UNIQUE NOT NULL,
            password        TEXT NOT NULL,
            prefix          TEXT UNIQUE NOT NULL,
            name            TEXT NOT NULL,
            min_rate        REAL DEFAULT 180,
            contact_phone   TEXT DEFAULT '',
            contact_email   TEXT DEFAULT '',
            status          TEXT DEFAULT 'active',
            note            TEXT DEFAULT '',
            created_at      TEXT NOT NULL
        )
    """)
    # ===== 會員表（代理建的客戶，存本地；你自己的客戶仍走 Shopify）=====
    conn.execute("""
        CREATE TABLE IF NOT EXISTS members (
            g_code          TEXT PRIMARY KEY,
            agent_id        INTEGER NOT NULL,
            name            TEXT NOT NULL,
            password        TEXT DEFAULT '',
            phone           TEXT DEFAULT '',
            address         TEXT DEFAULT '',
            line_id         TEXT DEFAULT '',
            email           TEXT DEFAULT '',
            note            TEXT DEFAULT '',
            status          TEXT DEFAULT 'active',
            created_at      TEXT NOT NULL
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_members_agent ON members(agent_id)")
    # 帳單欄位遷移（已存在的表加欄位）
    for col, col_type, default in [
        ("admin_note", "TEXT", "''"),
        ("updated_at", "TEXT", "NULL"),
        ("billed_weight", "REAL", "0"),
        ("rate_per_kg", "REAL", "0"),
        ("shipping_fee", "REAL", "0"),
        ("handling_fee", "REAL", "0"),
        ("total_fee", "REAL", "0"),
        ("payment_last5", "TEXT", "''"),
        ("payment_at", "TEXT", "''"),
        ("tracking_num", "TEXT", "''"),
        ("extra_services", "TEXT", "''"),
        ("ship_recipient", "TEXT", "''"),
        ("ship_phone", "TEXT", "''"),
        ("ship_address", "TEXT", "''"),
        ("consolidation_fee", "REAL", "0"),
        ("letter_fee", "REAL", "0"),
        ("boxes_json", "TEXT", "''"),   # 多箱明細 [{actual_weight,length,width,height,tracking_num,billed_weight,declarant_name,declarant_phone,declarant_address}]
        # 申報人：前三欄＝主申報人快照（後台每箱未指定時的預設值）
        ("declarant_name", "TEXT", "''"),
        ("declarant_phone", "TEXT", "''"),
        ("declarant_address", "TEXT", "''"),
        ("declarants_json", "TEXT", "''"),   # 本次授權可用的申報人快照 [{name,phone,address}]
        # 申報人同意聲明存證：申報人是報單上的納稅義務人，發生冒名報關爭議或本人否認授權時，
        # 必須舉證「客戶在什麼時間、對哪幾位申報人做過聲明」。對象即同一筆的 declarants_json。
        ("declarant_consent", "INTEGER", "0"),
        ("declarant_consent_at", "TEXT", "''"),
        ("declarant_consent_ip", "TEXT", "''"),
        # 出檔案給廠商（Nigel / JpD）追蹤欄位
        ("exported_at", "TEXT", "''"),
        ("exported_vendor", "TEXT", "''"),
        ("exported_batch_id", "TEXT", "''"),
    ]:
        try:
            conn.execute(f"ALTER TABLE shipment_requests ADD COLUMN {col} {col_type} DEFAULT {default}")
        except:
            pass

    # ===== Phase 2: 加 agent_id 欄位（既有資料預設 0 = 主管理員的）=====
    for table in ["packages", "forecasts", "shipment_requests", "announcements"]:
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN agent_id INTEGER DEFAULT 0")
            print(f"[migrate] 已加 {table}.agent_id 欄位", flush=True)
        except:
            pass
    # 索引加速 agent 過濾
    for table in ["packages", "forecasts", "shipment_requests"]:
        try:
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_agent ON {table}(agent_id)")
        except:
            pass

    # ===== Phase 3+: 加 members.shipping_rate 欄位（代理為每個客戶設定獨立費率）=====
    # 預設 0 = 沿用該代理的 min_rate；>0 = 該會員的專屬費率
    try:
        conn.execute("ALTER TABLE members ADD COLUMN shipping_rate REAL DEFAULT 0")
        print("[migrate] 已加 members.shipping_rate 欄位", flush=True)
    except:
        pass

    # ===== M1 會員主檔同步：Shopify 會員的本地副本，獨立一張表 =====
    # 為什麼不寫進 members：登入、會員列表、搜尋、出檔案 fallback 都直接讀 members，
    # 且都假設 members 只有代理/本地會員。Shopify 的 G 會員若混進去，讀取行為就變了。
    # 獨立表 = 物理上不可能影響既有讀取路徑，不用靠「每個查詢都記得加過濾條件」。
    # M1 只寫不讀；M2 切讀取時再決定怎麼合併。
    # 這張表是純 Shopify 副本（沒有本地獨有欄位），舊版欄位不同時直接重建（下次同步會補齊）。
    try:
        ms_cols = [r["name"] for r in conn.execute("PRAGMA table_info(members_shopify)").fetchall()]
        if ms_cols and ("shopify_created_at" not in ms_cols or "password" in ms_cols):
            conn.execute("DROP TABLE members_shopify")
            print("[migrate] members_shopify 舊結構已移除，將重建（純副本，同步會補齊）", flush=True)
    except Exception as e:
        print(f"[migrate] ⚠️ 檢查 members_shopify 結構失敗: {e}", flush=True)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS members_shopify (
            g_code              TEXT PRIMARY KEY,
            name                TEXT DEFAULT '',
            phone               TEXT DEFAULT '',
            email               TEXT DEFAULT '',
            address             TEXT DEFAULT '',
            shipping_rate       REAL DEFAULT 0,
            shopify_customer_id TEXT DEFAULT '',
            shopify_created_at  TEXT DEFAULT '',
            status              TEXT DEFAULT 'active',
            synced_at           TEXT DEFAULT '',
            phone_raw           TEXT DEFAULT ''
        )
    """)
    # M2：讀取切到本地後，phone_raw（Shopify 原始電話字串，客戶端顯示用）也要有，才能與舊路徑逐鍵相同
    try:
        conn.execute("ALTER TABLE members_shopify ADD COLUMN phone_raw TEXT DEFAULT ''")
        print("[migrate] 已加 members_shopify.phone_raw 欄位", flush=True)
    except:
        pass

    # ===== 包裹類型欄位：區分「包裹」與「信件」（信件計費一件 +NT$20，見帳單邏輯）=====
    try:
        conn.execute("ALTER TABLE packages ADD COLUMN pkg_type TEXT DEFAULT '包裹'")
        print("[migrate] 已加 packages.pkg_type 欄位", flush=True)
    except:
        pass

    # ===== 出檔案客戶編號（{g_code}-{MMDD}）：存起來供台灣配送貨況比對 =====
    try:
        conn.execute("ALTER TABLE shipment_requests ADD COLUMN export_code TEXT DEFAULT ''")
        print("[migrate] 已加 shipment_requests.export_code 欄位", flush=True)
    except:
        pass

    # ===== 客戶 × 廠商編號對照（出檔案給 Nigel / JpD 等廠商時用） =====
    # 對 Shopify 主帳號客戶 + 代理客戶都通用
    conn.execute("""
        CREATE TABLE IF NOT EXISTS customer_vendor_codes (
            g_code TEXT NOT NULL,
            vendor TEXT NOT NULL,
            code TEXT NOT NULL,
            updated_at TEXT,
            PRIMARY KEY (g_code, vendor)
        )
    """)

    # ===== 每團主每月結算快照（階梯運費的依據 + 貢獻報表的基礎）=====
    # 一列 = 一個 g_code 在一個年月的結算，但存了「語意不同的兩種東西」：
    #
    #   ① 衍生統計（total_kg / paid_kg / shipping_fee / …）
    #      純粹是 shipment_requests 的彙總，每次 _refresh_monthly_snapshot() 覆寫。
    #      出貨單被 revert、改重量、補收款，這些數字都會跟著變 —— 本來就該變。
    #
    #   ② applied_rate（本月帳單實際套用的每公斤費率）
    #      第一次被用到時寫入，之後永不覆寫（見 _effective_rate_for）。
    #      為什麼不能每次算：費率來自「上月 total_kg」，而上月快照是①、會變動。
    #      若每次重算，同一團主同月的兩張單會拿到不同費率，
    #      直接違反「當月帳單一出即定、不追溯回算」。凍結點就在這一欄。
    #
    # total_kg vs paid_kg 是刻意的兩套口徑，不要統一：
    #   total_kg = 已出貨即算（費率用這個，因為費率不該等客人匯款才決定）
    #   paid_kg  = 已出貨且已收款（月報統計用這個，與 /api/admin/stats/monthly 同口徑）
    conn.execute("""
        CREATE TABLE IF NOT EXISTS monthly_kg_snapshots (
            g_code          TEXT NOT NULL,
            ym              TEXT NOT NULL,          -- 'YYYY-MM'，依 updated_at 歸月
            total_kg        REAL DEFAULT 0,         -- 已出貨（不論收款）計費重量合計
            paid_kg         REAL DEFAULT 0,         -- 已出貨且已收款
            shipment_count  INTEGER DEFAULT 0,
            shipping_fee    REAL DEFAULT 0,         -- 運費收入（不含理貨/合箱/信件/加值）
            total_fee       REAL DEFAULT 0,         -- 帳單總額
            gross_margin    REAL DEFAULT 0,         -- 運費 − cost_per_kg × total_kg
            cost_per_kg     REAL DEFAULT 0,         -- 算此列毛利時用的成本（存下來，日後調整不會讓舊列失真）
            applied_rate    REAL DEFAULT 0,         -- 本月套用費率；0 = 尚未凍結
            rate_locked_at  TEXT DEFAULT '',        -- applied_rate 寫入時間
            rate_basis_kg   REAL DEFAULT 0,         -- 凍結當下讀到的「上月 total_kg」，存證用
            updated_at      TEXT DEFAULT '',
            PRIMARY KEY (g_code, ym)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_mks_ym ON monthly_kg_snapshots(ym)")

    # ===== GoyouLink 團主對照表 =====
    # 把「GoyouLink 的租戶」對應到「helpshipping 的一個出貨帳號」。
    # 對應關係放這邊而不是 GoyouLink：那邊是 N 個獨立 DB，放那邊會有 N 份、
    # 查不了「所有團主」、新增租戶要改 N 個地方。這裡是單一 DB，是唯一能做跨團主報表的地方。
    #
    # 這張表同時是階梯運費的「適用對象」名單（見 _is_goyoulink_tenant）：
    #   在表裡且 status='active' → 團主 → 套階梯
    #   不在表裡 / 已停用        → 散客 → 維持原費率
    # 所以表是空的時候，就算全站切到階梯制也不會有任何人受影響。
    #
    # created_at 不只是紀錄時間，它是「何時開始算團主」的依據：
    # 登錄當月一律第一階，次月起只採計 updated_at >= created_at 的出貨量。
    # 不吃登錄前的散客量 —— 那不是團主貢獻，且會被「先衝量再登錄套最低價」鑽。
    conn.execute("""
        CREATE TABLE IF NOT EXISTS goyoulink_tenants (
            tenant_key   TEXT PRIMARY KEY,      -- 自己發的穩定 ID，例：'ninita'
            display_name TEXT DEFAULT '',
            g_code       TEXT NOT NULL,         -- 對應的集運出貨帳號
            service_url  TEXT DEFAULT '',       -- Railway service 網址（做法B 才用得到）
            status       TEXT DEFAULT 'active', -- active / disabled
            note         TEXT DEFAULT '',
            created_at   TEXT NOT NULL          -- ＝生效起算時間，見上面說明
        )
    """)
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_gl_tenants_gcode ON goyoulink_tenants(g_code)")

    # ===== 代理品牌欄位（用於 referral URL + 登入後客製內容）=====
    for col, col_type, default in [
        ("contact_line", "TEXT", "''"),
        ("insurance_url", "TEXT", "''"),
        ("insurance_label", "TEXT", "''"),
        ("insurance_desc", "TEXT", "''"),
        ("signup_guide", "TEXT", "''"),
        ("promo_text", "TEXT", "''"),
        ("promo_price", "TEXT", "''"),
        ("owner_name", "TEXT", "''"),
        ("owner_address", "TEXT", "''"),
        ("bank_code", "TEXT", "''"),
        ("bank_name", "TEXT", "''"),
        ("bank_branch", "TEXT", "''"),
        ("bank_account", "TEXT", "''"),
        ("bank_account_name", "TEXT", "''"),
    ]:
        try:
            conn.execute(f"ALTER TABLE agents ADD COLUMN {col} {col_type} DEFAULT {default}")
            print(f"[migrate] 已加 agents.{col} 欄位", flush=True)
        except:
            pass

    # ===== 首次 seed 加值服務目錄（之後由後台管理，不覆寫既有值）=====
    try:
        has_cat = conn.execute("SELECT 1 FROM admin_settings WHERE key='extra_service_catalog'").fetchone()
        if not has_cat:
            conn.execute(
                "INSERT INTO admin_settings (key, value) VALUES ('extra_service_catalog', ?)",
                (json.dumps(DEFAULT_EXTRA_SERVICES, ensure_ascii=False),)
            )
            print("[migrate] 已 seed 加值服務目錄（26 項）", flush=True)
    except Exception as e:
        print(f"[migrate] ⚠️ seed 加值服務目錄失敗: {e}", flush=True)

    conn.commit()
    conn.close()


init_db()

# ============ 工具函數 ============

def normalize_phone(phone_raw):
    phone = phone_raw.replace(" ", "").replace("-", "")
    if phone.startswith("+886"):
        phone = "0" + phone[4:]
    elif phone.startswith("+81"):
        phone = "0" + phone[3:]
    elif phone.startswith("886") and len(phone) == 12 and phone[3:].isdigit() and phone[3] == "9":
        # 沒有 '+' 的 886 格式：REAL affinity 曾把 '+' 吃掉，殘留這種值。
        # 電話一律以 0 開頭儲存，這裡把前綴換回 0。
        # （'+81' 的規則維持原樣，日本號碼規則不同，不一起改。）
        phone = "0" + phone[3:]
    return phone


def twd_to_jpy(twd_rate):
    """台幣運費 → 日圓運費（四捨五入至整數）"""
    return round(twd_rate * TWD_TO_JPY_RATE)


def jpd_request(operation, data):
    url = f"{JPD_BASE_URL}/api/json.php?Service=SDC&Operation={operation}"
    payload = {
        "login_email": JPD_EMAIL,
        "login_password": JPD_PASSWORD,
        "data": data
    }
    print(f"\n{'='*50}")
    print(f"📤 JPD API 請求: {operation}")
    try:
        response = requests.post(url, json=payload, timeout=30)
        result = response.json()
        return result
    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return {"error": str(e)}


def shopify_graphql(query, variables=None):
    graphql_url = f"https://{SHOPIFY_STORE}/admin/api/2026-01/graphql.json"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    payload = {"query": query}
    if variables:
        payload["variables"] = variables
    try:
        response = requests.post(graphql_url, headers=headers, json=payload, timeout=15)
        return response.json()
    except Exception as e:
        print(f"❌ GraphQL 錯誤: {e}")
        return {"error": str(e)}


def shopify_request(endpoint, method="GET", data=None):
    url = f"https://{SHOPIFY_STORE}/admin/api/2026-01/{endpoint}"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    try:
        if method == "GET":
            response = requests.get(url, headers=headers, timeout=30)
        elif method == "POST":
            response = requests.post(url, headers=headers, json=data, timeout=30)
        return response.json()
    except Exception as e:
        return {"error": str(e)}


# ============ Shopify 會員快取 ============
# 持久化到磁碟（容器重啟、Zeabur 重新部署都不用重抓）+ stale-while-revalidate

# 快取檔案位置：跟 DB 放同個目錄（Zeabur Volume 持久化）
_db_dir = os.path.dirname(os.path.abspath(DB_PATH))
# 舊的 Shopify 記憶體/磁碟快取層已在 M2 移除（讀取改走本地 members_shopify，見下方）。
# 這個檔案保留不刪、也不再寫入：只在「members_shopify 為空且 Shopify 也抓不到」的災難情境唯讀使用。
SHOPIFY_CACHE_FILE = os.environ.get(
    "SHOPIFY_CACHE_FILE",
    os.path.join(_db_dir or ".", "shopify_cache.json")
)

# ============ 會員讀取（M2：本地 members_shopify）============
# get_all_goyoutati_customers() 的簽章與回傳格式完全不變（7 個呼叫點一行未改），
# 只把資料來源從「Shopify 快取」改成 SELECT … FROM members_shopify。
# ★ 回傳 dict 必須與 _fetch_customers_from_shopify() 產出的逐鍵相同：
#   g_code / customer_id / gid / name / email / address / phone / phone_raw / shipping_rate / created_at
#   shipping_rate 沿用舊格式：字串（"180"；沒設定 = ""），登入端用 int(c["shipping_rate"]) 解析。
#   phone_raw：Shopify 原始電話字串（客戶端顯示用），同步時一併存進 members_shopify.phone_raw。
CUSTOMER_DICT_KEYS = ("g_code", "customer_id", "gid", "name", "email", "address",
                      "phone", "phone_raw", "shipping_rate", "created_at")


def _rate_float_to_str(v):
    """members_shopify.shipping_rate（REAL）→ 舊路徑的字串格式：180.0 → "180"、180.5 → "180.5"、0/None → ""。"""
    try:
        f = float(v or 0)
    except (ValueError, TypeError):
        return ""
    if f <= 0:
        return ""
    return str(int(f)) if f == int(f) else str(f)


def _customer_dict_from_local_row(r):
    cid = r["shopify_customer_id"] or ""
    phone = r["phone"] or ""
    return {
        "g_code": r["g_code"] or "",
        "customer_id": cid,
        "gid": f"gid://shopify/Customer/{cid}" if cid else "",
        "name": r["name"] or "",
        "email": r["email"] or "",
        "address": r["address"] or "",
        "phone": phone,
        "phone_raw": r["phone_raw"] or phone,   # 舊資料還沒同步到 phone_raw 前先用正規化值
        "shipping_rate": _rate_float_to_str(r["shipping_rate"]),  # 台幣，字串
        "created_at": r["shopify_created_at"] or "",
    }


def _load_customers_from_local():
    """SELECT members_shopify → 與 Shopify 路徑同格式的 list。
    status='shopify_missing'（Shopify 上已消失）的不回：舊路徑本來就抓不到他們。"""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT g_code, name, phone, phone_raw, email, address, shipping_rate, shopify_customer_id, shopify_created_at "
            "FROM members_shopify WHERE status != 'shopify_missing' ORDER BY g_code"
        ).fetchall()
    finally:
        conn.close()
    return [_customer_dict_from_local_row(r) for r in rows]


def _load_customers_from_disk_backup():
    """災難備援：members_shopify 為空、Shopify 也抓不到 → 讀舊的 shopify_cache.json（唯讀）。
    內容可能很舊，但比整站沒有會員好。"""
    try:
        if not os.path.exists(SHOPIFY_CACHE_FILE):
            return []
        with open(SHOPIFY_CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        out = []
        for c in (data.get("data") or []):
            if not isinstance(c, dict) or not c.get("g_code"):
                continue
            d = {k: c.get(k, "") for k in CUSTOMER_DICT_KEYS}
            d = {k: ("" if v is None else v) for k, v in d.items()}
            out.append(d)
        return out
    except Exception as e:
        print(f"[members] ⚠️ 讀取 shopify_cache.json 備援失敗: {e}", flush=True)
        return []


def _recent_sync_attempt(within_sec):
    """最近 within_sec 秒內是否已嘗試過同步（成功失敗都算；看 last_report.synced_at）。
    用途：members_shopify 為空時避免每個請求都去打 Shopify。"""
    try:
        rep = json.loads(_get_setting(MEMBER_SYNC_REPORT_KEY, "") or "{}")
        ts = datetime.strptime((rep.get("synced_at") or "")[:19], "%Y-%m-%d %H:%M:%S")
        return (datetime.now() - ts).total_seconds() < within_sec
    except Exception:
        return False


def get_all_goyoutati_customers(force_refresh=False):
    """
    取得 Shopify 會員清單 —— M2 起讀本地 members_shopify（由 M1 同步維護），不再打 Shopify。
      • force_refresh=True（admin 按「整理」）：先跑一次同步再讀本地；同步失敗仍回本地現有資料
      • 一般：直接讀本地
      • 降級：本地為空 → 嘗試同步一次（60 秒內只試一次）→ 仍為空就讀 shopify_cache.json 備援
    回傳的是每次新建的 list / dict，呼叫端 append / sort 不會影響任何共用狀態。
    """
    if force_refresh:
        try:
            rep = run_member_sync(trigger="force_refresh")   # 永不拋例外；租約搶不到會回 skipped
            if not rep.get("success"):
                print(f"[members] force_refresh 同步未完成（{rep.get('error')}），改回本地現有資料", flush=True)
        except Exception as e:
            print(f"[members] force_refresh 同步例外（改回本地現有資料）: {e}", flush=True)

    try:
        customers = _load_customers_from_local()
    except Exception as e:
        print(f"[members] ❌ 讀取 members_shopify 失敗: {e}", flush=True)
        customers = []
    if customers:
        return customers

    # ── 降級：本地表為空 ──
    if not force_refresh and not _recent_sync_attempt(60):
        print("[members] ⚠️ members_shopify 為空 → 嘗試同步一次", flush=True)
        try:
            run_member_sync(trigger="empty_table")
            customers = _load_customers_from_local()
        except Exception as e:
            print(f"[members] 空表同步例外: {e}", flush=True)
    if customers:
        return customers
    backup = _load_customers_from_disk_backup()
    print(f"[members] ⚠️ members_shopify 仍為空，改用 shopify_cache.json 備援：{len(backup)} 位", flush=True)
    return backup


def _fetch_customers_from_shopify():
    customers = []
    cursor = None
    has_next = True
    page = 0

    while has_next and page < 10:  # 最多 10 頁 = 1000 會員
        page += 1
        after_arg = f', after: "{cursor}"' if cursor else ''
        graphql_query = '{metafieldDefinitions(first:1,ownerType:CUSTOMER,namespace:"custom",key:"' + MEMBER_METAFIELD_KEY + '"){edges{node{id metafields(first:100' + after_arg + '){edges{node{value owner{...on Customer{id firstName lastName email phone defaultAddress{phone province city zip address1 address2} createdAt shippingRate:metafield(namespace:"custom",key:"shipping_rate"){value}}}} cursor} pageInfo{hasNextPage}}}}}}'

        page_t0 = time.time()
        result = shopify_graphql(graphql_query)
        page_ms = int((time.time() - page_t0) * 1000)
        has_next = False

        if "data" not in result:
            print(f"[Shopify] page {page} error ({page_ms}ms): {result}", flush=True)
            break

        definitions = result["data"].get("metafieldDefinitions", {}).get("edges", [])
        if not definitions:
            print("[Shopify] No metafieldDefinitions found", flush=True)
            break

        metafields_data = definitions[0]["node"].get("metafields", {})
        edges = metafields_data.get("edges", [])
        page_info = metafields_data.get("pageInfo", {})
        has_next = page_info.get("hasNextPage", False)
        print(f"[Shopify] page {page} ({page_ms}ms): got {len(edges)} metafields, hasNextPage={has_next}", flush=True)

        for mf in edges:
            node = mf["node"]
            cursor = mf.get("cursor")
            g_code = node.get("value", "")
            owner = node.get("owner", {})
            if not g_code or not owner:
                continue
            customers.append(_customer_dict_from_shopify_owner(g_code, owner))
    return customers


def _customer_dict_from_shopify_owner(g_code, owner):
    """Shopify Customer 節點 → 會員 dict（欄位順序/型別是全站的基準格式，
    本地讀取 _customer_dict_from_local_row 與登入 fallback 都必須產出相同的鍵）。"""
    gid = owner.get("id", "")
    customer_id = gid.split("/")[-1] if "/" in gid else gid
    customer_name = f"{owner.get('lastName', '')}{owner.get('firstName', '')}".strip()
    if not customer_name:
        customer_name = owner.get("email", "")
    default_address = owner.get("defaultAddress") or {}
    phone_raw = default_address.get("phone") or owner.get("phone") or ""
    phone = normalize_phone(phone_raw)
    # 用郵遞區號反查補齊缺的縣市/區（Shopify 拆欄常漏縣市區 → 黑貓無法投遞）
    address, _addr_fixed = tw_zip.compose_full_address(
        default_address.get("province", ""),
        default_address.get("city", ""),
        default_address.get("address1", ""),
        default_address.get("address2", ""),
        default_address.get("zip", ""),
    )
    rate_mf = owner.get("shippingRate")
    # shipping_rate 現在儲存台幣值
    shipping_rate_twd = rate_mf["value"] if rate_mf and rate_mf.get("value") else ""
    return {
        "g_code": g_code,
        "customer_id": customer_id,
        "gid": gid,
        "name": customer_name,
        "email": owner.get("email", ""),
        "address": address,
        "phone": phone,
        "phone_raw": phone_raw,
        "shipping_rate": shipping_rate_twd,  # 台幣
        "created_at": owner.get("createdAt", "")
    }


# ============ 登入 fallback：本地查無 → Shopify 單筆補撈（T2）============
# 新客人剛在 Shopify 貼完 會員編號 metafield、定時同步還沒輪到的空窗期保險。
# Shopify 的 customers 搜尋語法不支援 metafield 過濾，所以抓「最近更新的 100 位」
# （貼 metafield 會更新客戶的 updatedAt）並在其中找該編號；一個請求、5 秒逾時、永不拋例外。
# 每次觸發都 print [member_fallback] g_code=… hit=…，次數 = 手貼 metafield 時間差的量化指標。
#
# 節流（T-fix）：抓回來的 100 筆在記憶體快取 MEMBER_FALLBACK_CACHE_SEC 秒。
# 打錯客編、掃客編的每一次失敗登入都會走到 fallback，若每次都打 Shopify，一支換編號的迴圈
# 就能放大成 N 次請求；快取後 60 秒內不論試幾個編號都只打 1 次。
# ★ 只存這一批的「編號 → 客戶節點」查找表，不落盤、不與其他快取共用；
#   多 worker 各持一份可接受（最多放大 worker 數倍，不是 N 倍），不引入共用儲存。
# ★ 只快取「成功抓到」的結果：逾時／例外不快取，下一次再重抓（登入本身有 5 秒逾時，不會卡住）。
MEMBER_FALLBACK_TIMEOUT_SEC = 5
MEMBER_FALLBACK_SCAN = 100
MEMBER_FALLBACK_CACHE_SEC = 60

_fallback_lock = threading.Lock()
# {"at": 抓取時間 time.time(), "by_code": {g_code: Shopify customer node}}；at=0 表示沒有快取
_fallback_cache = {"at": 0.0, "by_code": {}}
# 自程序啟動起算、不持久化。attempts 高但 hits≈0 → 多半是打錯客編而非空窗期，fallback 價值要重評。
_fallback_stats = {"attempts": 0, "hits": 0}


def _fallback_cache_get(g_code, now):
    """快取有效期內回傳 (True, node_or_None)；過期或沒有快取回傳 (False, None)。"""
    with _fallback_lock:
        if _fallback_cache["at"] and now - _fallback_cache["at"] < MEMBER_FALLBACK_CACHE_SEC:
            return True, _fallback_cache["by_code"].get(g_code)
    return False, None


def _fallback_cache_put(by_code, now):
    with _fallback_lock:
        _fallback_cache["at"] = now
        _fallback_cache["by_code"] = by_code


def _fallback_stats_snapshot():
    with _fallback_lock:
        return dict(_fallback_stats)


def _upsert_member_shopify(c):
    """把一筆會員 dict 寫進 members_shopify（fallback 補撈用）。已存在則更新五欄。
    若編號已是代理/本地會員（members 表）則不寫（與同步的衝突規則一致）。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    code = (c.get("g_code") or "").strip().upper()
    try:
        rate = float(c.get("shipping_rate") or 0)
    except (ValueError, TypeError):
        rate = 0.0
    conn = get_db()
    try:
        if conn.execute("SELECT 1 FROM members WHERE g_code=?", (code,)).fetchone():
            print(f"[member_fallback] {code} 已存在於 members（代理/本地會員），不寫入 members_shopify", flush=True)
            return False
        conn.execute(
            "INSERT INTO members_shopify (g_code, name, phone, phone_raw, email, address, shipping_rate, "
            "shopify_customer_id, shopify_created_at, status, synced_at) VALUES (?,?,?,?,?,?,?,?,?,'active',?) "
            "ON CONFLICT(g_code) DO UPDATE SET name=excluded.name, phone=excluded.phone, phone_raw=excluded.phone_raw, "
            "email=excluded.email, address=excluded.address, shipping_rate=excluded.shipping_rate, "
            "shopify_customer_id=excluded.shopify_customer_id, status='active', synced_at=excluded.synced_at",
            (code, (c.get("name") or "").strip(), normalize_phone(c.get("phone") or ""), c.get("phone_raw") or "",
             (c.get("email") or "").strip(), (c.get("address") or "").strip(), rate,
             str(c.get("customer_id") or ""), str(c.get("created_at") or ""), now)
        )
        conn.commit()
        return True
    finally:
        conn.close()


# 已知盲區：
# 本 fallback 取最近更新的 100 位會員。若貼上 metafield 之後
# 有超過 100 位客戶的 Shopify 資料被更新，該編號會掉出視窗而撈不到。
# Shopify 不支援以 metafield 過濾，這是變通做法。
# 撈不到時客人會收到『找不到此會員編號』，等下一輪同步後即可登入。
# 另外，快取期間（60 秒）內新貼編號的客人也會撈不到——快取裡沒有他，且不會重抓；
# 等快取過期或下一輪同步即可，這是節流的預期行為。
def _member_login_fallback(g_code):
    """本地 members / members_shopify 都查無此編號時呼叫。
    回傳與 get_all_goyoutati_customers() 同格式的 dict，或 None。永不拋例外、有逾時上限。"""
    t0 = time.time()
    hit = False
    cached = False
    err = ""
    with _fallback_lock:
        _fallback_stats["attempts"] += 1
    try:
        if not SHOPIFY_STORE or not SHOPIFY_ACCESS_TOKEN:
            err = "未設定 Shopify"
            return None
        cached, owner = _fallback_cache_get(g_code, t0)
        if cached:
            if owner is None:
                return None
            hit = True
            return _fallback_hit(g_code, owner)
        query = (
            '{customers(first:' + str(MEMBER_FALLBACK_SCAN) + ',sortKey:UPDATED_AT,reverse:true){edges{node{'
            'id firstName lastName email phone createdAt defaultAddress{phone province city zip address1 address2} '
            'gcode:metafield(namespace:"custom",key:"' + MEMBER_METAFIELD_KEY + '"){value} '
            'shippingRate:metafield(namespace:"custom",key:"shipping_rate"){value}}}}}'
        )
        resp = requests.post(
            f"https://{SHOPIFY_STORE}/admin/api/2026-01/graphql.json",
            headers={"X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN, "Content-Type": "application/json"},
            json={"query": query}, timeout=MEMBER_FALLBACK_TIMEOUT_SEC
        )
        data = resp.json()
        if "data" not in data:
            err = f"Shopify 回應無 data: {str(data)[:120]}"
            return None
        by_code = {}
        for edge in (data["data"].get("customers") or {}).get("edges", []):
            owner = edge.get("node") or {}
            code = ((owner.get("gcode") or {}).get("value") or "").strip().upper()
            if code and code not in by_code:   # 同編號多筆時保留最近更新的那筆（列表已依 UPDATED_AT 降冪）
                by_code[code] = owner
        _fallback_cache_put(by_code, t0)
        owner = by_code.get(g_code)
        if owner is None:
            return None
        hit = True
        return _fallback_hit(g_code, owner)
    except Exception as e:
        err = f"{type(e).__name__}: {e}"
        return None
    finally:
        if hit:
            with _fallback_lock:
                _fallback_stats["hits"] += 1
        print(f"[member_fallback] g_code={g_code} hit={hit} cached={cached} {time.time() - t0:.2f}s"
              + (f" error={err}" if err else ""), flush=True)


def _fallback_hit(g_code, owner):
    """命中後的共同處理：轉成登入用 dict 並順手寫進 members_shopify（寫入失敗不影響登入）。"""
    c = _customer_dict_from_shopify_owner(g_code, owner)
    try:
        _upsert_member_shopify(c)
    except Exception as e:
        print(f"[member_fallback] 寫入 members_shopify 失敗（登入照常繼續）: {e}", flush=True)
    return c


# ============ 會員主檔同步（M1：只寫不讀）============
# 目的：讓集運系統有自己的會員主檔，為日後搬離 Shopify 做準備。
# 本階段只建立「Shopify → members_shopify」的單向同步，★ 不改任何讀取路徑：
# 登入、會員列表、統計全部照舊走 Shopify 快取（get_all_goyoutati_customers 與 7 個呼叫點一字未改），
# members 表（代理/本地會員）一個欄位、一筆資料都不動。同步出錯時對營運零影響。
#
# 安全規則（硬規則，改動前先讀）：
#   • 只寫 members_shopify。members 表只讀來偵測衝突，永遠不寫。
#   • Shopify 編號若已存在於 members（代理/本地會員，例如貼錯的 B0049）→ 不寫入，記入 conflicts。
#   • 更新覆蓋 name / phone / phone_raw / email / address / shipping_rate + synced_at（這張表全部都是 Shopify 來源）；
#     六欄都沒變就不寫（冪等）。status 只有同步自己會動（shopify_missing ↔ active）。
#   • 不刪會員。Shopify 上消失的只標 status='shopify_missing'，資料留著（歷史出貨憑證）。
#   • 同一個 會員編號 metafield 掛在兩個以上 Shopify 客戶身上 → 該編號整個跳過，記入 duplicates。
#   • Shopify 回空或拋例外 → 一筆都不寫，回報失敗。
#   • 多 worker 防重用 admin_settings 的租約（條件式 UPDATE），不用記憶體變數。
MEMBER_SYNC_LEASE_KEY = "member_sync_lease"          # 值：unix epoch（0 = 沒人在跑）
MEMBER_SYNC_LEASE_SEC = 300                          # 租約有效期；超過視為前一個 worker 掛了，可搶
MEMBER_SYNC_INTERVAL_KEY = "member_sync_interval_min"   # 同步間隔（分鐘），admin_settings 可調
MEMBER_SYNC_INTERVAL_DEFAULT_MIN = 30                # 預設 30。M2 起這個值 = 本地資料（含計費用的 shipping_rate）的新鮮度上限，別隨手調長
MEMBER_SYNC_INTERVAL_MIN_MIN, MEMBER_SYNC_INTERVAL_MAX_MIN = 5, 1440
MEMBER_SYNC_RETRY_BACKOFF_SEC = 300                  # 上次嘗試失敗後至少隔 5 分鐘再試（Shopify 掛掉時不要每分鐘打）
MEMBER_SYNC_TICK_SEC = 60                            # 定時執行緒每分鐘檢查一次「到期了沒」（間隔改了不用重啟、啟動時到期就立刻跑）
MEMBER_SYNC_LAST_KEY = "member_sync_last_at"         # 最後一次成功同步時間（'%Y-%m-%d %H:%M:%S'）
MEMBER_SYNC_REPORT_KEY = "member_sync_last_report"   # 最後一次同步回報（JSON，成功失敗都存）
MEMBER_SYNC_FIELDS = ("name", "phone", "phone_raw", "email", "address", "shipping_rate")  # 更新時唯一允許覆蓋的欄位


def _acquire_member_sync_lease():
    """搶同步租約。用 admin_settings 做條件式 UPDATE：舊值 < now - 300 才搶得到。
    跨 worker 有效（gunicorn 2 個 worker 各自跑定時器，只能有一個實際執行）。
    回 True 表示搶到；搶到的人做完要呼叫 _release_member_sync_lease()。"""
    now_ts = int(time.time())
    conn = get_db()
    try:
        conn.execute(
            "INSERT OR IGNORE INTO admin_settings (key, value) VALUES (?, '0')",
            (MEMBER_SYNC_LEASE_KEY,)
        )
        cur = conn.execute(
            "UPDATE admin_settings SET value=? WHERE key=? AND CAST(value AS INTEGER) < ?",
            (str(now_ts), MEMBER_SYNC_LEASE_KEY, now_ts - MEMBER_SYNC_LEASE_SEC)
        )
        conn.commit()
        return cur.rowcount == 1
    finally:
        conn.close()


def _release_member_sync_lease():
    try:
        _set_setting(MEMBER_SYNC_LEASE_KEY, "0")
    except Exception as e:
        print(f"[member_sync] 釋放租約失敗（{MEMBER_SYNC_LEASE_SEC}s 後自動失效）: {e}", flush=True)


def _member_sync_is_running():
    """租約值非 0 且未過期 → 有人正在跑。"""
    try:
        v = int(_get_setting(MEMBER_SYNC_LEASE_KEY, "0") or "0")
    except (ValueError, TypeError):
        v = 0
    return v > 0 and (time.time() - v) < MEMBER_SYNC_LEASE_SEC


def _shopify_rate_to_float(v):
    try:
        return float(v) if v not in (None, "", 0, "0") else 0.0
    except (ValueError, TypeError):
        return 0.0


def sync_members_from_shopify(customers=None):
    """Shopify 會員 → 本地 members_shopify 主檔（只寫不讀）。

    customers=None 時用既有的 _fetch_customers_from_shopify() 抓；測試可直接注入清單。
    回傳 dict（success / inserted / updated / unchanged / shopify_missing /
    conflicts[] / duplicates[] / shopify_total / local_total / elapsed_sec）。
    ★ 本函式不搶租約；租約由 run_member_sync() 負責。"""
    t0 = time.time()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report = {
        "success": False, "inserted": 0, "updated": 0, "unchanged": 0,
        "shopify_missing": 0, "shopify_missing_total": 0,
        "conflicts": [], "duplicates": [], "skipped_invalid": 0,
        "shopify_total": 0, "local_total": 0, "elapsed_sec": 0.0,
        "error": "", "synced_at": now,
    }

    # ── 1. 抓 Shopify（回空或例外 → 一筆都不寫）──
    try:
        if customers is None:
            customers = _fetch_customers_from_shopify()
    except Exception as e:
        report["error"] = f"Shopify 抓取失敗：{e}"
        report["elapsed_sec"] = round(time.time() - t0, 2)
        return report
    if not customers:
        report["error"] = "Shopify 回空（不寫入，保留本地資料）"
        report["elapsed_sec"] = round(time.time() - t0, 2)
        return report
    report["shopify_total"] = len(customers)

    # ── 2. 重複偵測：同一 會員編號 metafield 掛在兩個以上客戶 → 整個編號跳過 ──
    by_code = {}
    for c in customers:
        code = (c.get("g_code") or "").strip().upper()
        if not code:
            report["skipped_invalid"] += 1
            continue
        by_code.setdefault(code, []).append(c)
    duplicates = sorted(code for code, lst in by_code.items() if len(lst) > 1)
    for code in duplicates:
        ids = [c.get("customer_id", "") for c in by_code[code]]
        print(f"[member_sync] ⚠️ 重複 會員編號 {code}：Shopify 客戶 {ids}，整個跳過不寫", flush=True)
    report["duplicates"] = duplicates
    shopify_codes = set(by_code.keys())   # 含重複的（Shopify 上確實存在，不能標 missing）

    # ── 3. 逐筆比對寫入（單一交易，失敗全部回滾）──
    conn = get_db()
    try:
        conn.execute("BEGIN IMMEDIATE")
        # members 表（代理/本地會員）只讀：Shopify 上若出現同編號就是衝突，一律不寫
        agent_codes = {r["g_code"] for r in conn.execute("SELECT g_code FROM members").fetchall()}
        local_rows = {
            r["g_code"]: r for r in conn.execute(
                "SELECT g_code, name, phone, phone_raw, email, address, shipping_rate, status, shopify_customer_id "
                "FROM members_shopify"
            ).fetchall()
        }
        for code in sorted(by_code.keys()):
            if code in duplicates:
                continue
            if code in agent_codes:
                # 【衝突】編號已是代理/本地會員 → 完全不寫入（members 與 members_shopify 都不動）
                report["conflicts"].append(code)
                print(f"[member_sync] ⚠️ 衝突 {code}：已存在於 members（代理/本地會員），不寫入", flush=True)
                continue
            c = by_code[code][0]
            vals = {
                "name": (c.get("name") or "").strip(),
                "phone": normalize_phone(c.get("phone") or c.get("phone_raw") or ""),
                "phone_raw": c.get("phone_raw") or "",
                "email": (c.get("email") or "").strip(),
                "address": (c.get("address") or "").strip(),
                "shipping_rate": _shopify_rate_to_float(c.get("shipping_rate")),
            }
            customer_id = str(c.get("customer_id") or "")
            row = local_rows.get(code)

            if row is None:
                # 【新增】本地沒有 → INSERT（shopify_created_at 存 Shopify 原始 ISO 字串）
                conn.execute(
                    "INSERT INTO members_shopify (g_code, name, phone, phone_raw, email, address, shipping_rate, "
                    "shopify_customer_id, shopify_created_at, status, synced_at) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', ?)",
                    (code, vals["name"], vals["phone"], vals["phone_raw"], vals["email"], vals["address"], vals["shipping_rate"],
                     customer_id, str(c.get("created_at") or ""), now)
                )
                report["inserted"] += 1
                continue

            # 【更新】比對五個欄位 + customer_id，有變才覆蓋（連同 synced_at）
            changed = False
            for f in MEMBER_SYNC_FIELDS:
                cur = row[f]
                if f == "shipping_rate":
                    if _shopify_rate_to_float(cur) != vals[f]:
                        changed = True
                elif (cur or "") != vals[f]:
                    changed = True
            if (row["shopify_customer_id"] or "") != customer_id:
                changed = True
            # 之前被標 shopify_missing、現在又出現 → 還原 active（這個標記是同步自己打的，才允許改回）
            restore_status = (row["status"] == "shopify_missing")
            if changed or restore_status:
                conn.execute(
                    "UPDATE members_shopify SET name=?, phone=?, phone_raw=?, email=?, address=?, shipping_rate=?, "
                    "shopify_customer_id=?, synced_at=?"
                    + (", status='active'" if restore_status else "")
                    + " WHERE g_code=?",
                    (vals["name"], vals["phone"], vals["phone_raw"], vals["email"], vals["address"], vals["shipping_rate"],
                     customer_id, now, code)
                )
                report["updated"] += 1
            else:
                report["unchanged"] += 1

        # 【消失】本地有、Shopify 找不到 → 只標 status，不刪
        for code, row in local_rows.items():
            if code in shopify_codes:
                continue
            if row["status"] != "shopify_missing":
                conn.execute(
                    "UPDATE members_shopify SET status='shopify_missing', synced_at=? WHERE g_code=?",
                    (now, code)
                )
                report["shopify_missing"] += 1
                print(f"[member_sync] {code} 在 Shopify 上找不到 → status='shopify_missing'（資料保留）", flush=True)

        conn.commit()
        report["shopify_missing_total"] = conn.execute(
            "SELECT COUNT(*) AS c FROM members_shopify WHERE status='shopify_missing'"
        ).fetchone()["c"]
        report["local_total"] = conn.execute("SELECT COUNT(*) AS c FROM members_shopify").fetchone()["c"]
        report["success"] = True
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        report["error"] = f"寫入失敗（已回滾）：{e}"
    finally:
        conn.close()

    report["elapsed_sec"] = round(time.time() - t0, 2)
    return report


def run_member_sync(trigger="manual", customers=None):
    """搶租約 → 同步 → 存回報 → 釋放租約。永遠不拋例外（同步失敗不能影響任何請求）。
    回傳 report；搶不到租約時回 {"success": False, "skipped": True, "error": "同步進行中"}。"""
    try:
        if not _acquire_member_sync_lease():
            return {"success": False, "skipped": True, "error": "同步進行中（另一個 worker 正在執行）"}
    except Exception as e:
        print(f"[member_sync] 搶租約失敗: {e}", flush=True)
        return {"success": False, "skipped": True, "error": f"租約取得失敗：{e}"}
    try:
        report = sync_members_from_shopify(customers)
        report["trigger"] = trigger
        try:
            if report.get("success"):
                _set_setting(MEMBER_SYNC_LAST_KEY, report["synced_at"])
            _set_setting(MEMBER_SYNC_REPORT_KEY, json.dumps(report, ensure_ascii=False))
        except Exception as e:
            print(f"[member_sync] 存回報失敗: {e}", flush=True)
        print(f"[member_sync] ({trigger}) {json.dumps(report, ensure_ascii=False)}", flush=True)
        return report
    except Exception as e:
        print(f"[member_sync] ❌ 未預期例外: {e}", flush=True)
        return {"success": False, "error": f"未預期例外：{e}"}
    finally:
        _release_member_sync_lease()


def _get_sync_interval_min():
    """同步間隔（分鐘）。設定值壞掉或超出範圍就回預設 30。"""
    try:
        v = int(str(_get_setting(MEMBER_SYNC_INTERVAL_KEY, "") or "").strip())
        if MEMBER_SYNC_INTERVAL_MIN_MIN <= v <= MEMBER_SYNC_INTERVAL_MAX_MIN:
            return v
    except (ValueError, TypeError):
        pass
    return MEMBER_SYNC_INTERVAL_DEFAULT_MIN


def _seconds_since_setting_ts(key, from_report=False):
    """距 admin_settings 裡某個時間戳幾秒；沒有或壞掉回 None。"""
    try:
        raw = _get_setting(key, "")
        if from_report:
            raw = (json.loads(raw or "{}").get("synced_at") or "")
        ts = datetime.strptime(raw[:19], "%Y-%m-%d %H:%M:%S")
        return (datetime.now() - ts).total_seconds()
    except Exception:
        return None


def _member_sync_due(now_interval_min=None):
    """到期判斷：距上次「成功」已 ≥ 間隔（或從未成功），且距上次「嘗試」≥ 5 分鐘（失敗退避）。"""
    interval_sec = (now_interval_min or _get_sync_interval_min()) * 60
    since_ok = _seconds_since_setting_ts(MEMBER_SYNC_LAST_KEY)
    if since_ok is not None and since_ok < interval_sec:
        return False
    since_try = _seconds_since_setting_ts(MEMBER_SYNC_REPORT_KEY, from_report=True)
    if since_try is not None and since_try < MEMBER_SYNC_RETRY_BACKOFF_SEC:
        return False
    return True


def _member_sync_tick():
    """定時執行緒的一次迴圈：每次都重讀間隔設定（改了不用重啟）；到期才跑，租約保證跨 worker 只跑一個。
    容器啟動後第一次 tick 就會檢查，所以「距上次同步已超過間隔」的話會立刻補跑，不會因為重啟重置計時器。"""
    try:
        if _member_sync_due():
            run_member_sync(trigger="scheduled")
            return True
    except Exception as e:
        print(f"[member_sync] 定時任務例外: {e}", flush=True)
    return False


def _member_sync_loop():
    time.sleep(5)   # 讓 worker 完成啟動
    while True:
        _member_sync_tick()
        time.sleep(MEMBER_SYNC_TICK_SEC)


def _start_member_sync_thread():
    """MEMBER_SYNC_AUTO=0 可關閉（測試用）。"""
    if os.environ.get("MEMBER_SYNC_AUTO", "1") != "1":
        return
    try:
        threading.Thread(target=_member_sync_loop, daemon=True, name="MemberSync").start()
        print(f"[member_sync] 背景定時同步已啟動（間隔 {_get_sync_interval_min()} 分鐘，每 {MEMBER_SYNC_TICK_SEC}s 檢查到期）", flush=True)
    except Exception as e:
        print(f"[member_sync] 背景執行緒啟動失敗: {e}", flush=True)


_start_member_sync_thread()


# ============ SQLite 每日自動備份 → Google Drive ============
# 2026-09-18 Zeabur volume 被誤刪、/data/packages.db 全部遺失，前一晚的手動備份沒留住。
# 教訓：備份必須自動、必須離開這台機器。同一顆磁碟上的備份已證明沒用。
# 實作細節（快照 / 驗證 / 上傳 / 清理）在 gdrive_backup.py；這裡只做租約、排程、狀態、端點。
#   • 三個環境變數任一為空 → 備份停用，只在啟動時印一次警告，其餘功能完全不受影響。
#   • 每日一次，時間 admin_settings.backup_hour（預設 3 ＝ 容器時區凌晨三點；容器是 UTC 就是台灣 11:00）。
#   • 多 worker 防重：與會員同步相同的 admin_settings 租約（條件式 UPDATE），不用記憶體變數。
#   • 失敗當天每 30 分鐘重試；成功後當天不再跑。
BACKUP_LEASE_KEY = "backup_lease"                 # unix epoch（0 = 沒人在跑）
BACKUP_LEASE_SEC = 900                            # 租約有效期（含上傳）；超過視為前一個 worker 掛了
BACKUP_HOUR_KEY = "backup_hour"                   # 每日幾點跑（0~23，容器本地時間）
BACKUP_HOUR_DEFAULT = 3
BACKUP_LAST_OK_KEY = "backup_last_ok_at"          # 最後一次成功（'%Y-%m-%d %H:%M:%S'）
BACKUP_LAST_FAIL_KEY = "backup_last_fail_at"      # 最後一次失敗
BACKUP_LAST_FAIL_REASON_KEY = "backup_last_fail_reason"
BACKUP_LAST_ATTEMPT_KEY = "backup_last_attempt_at"
BACKUP_REPORT_KEY = "backup_last_report"          # 最後一次 report（JSON，成功失敗都存）
BACKUP_FOLDER_KEY = "gdrive_backup_folder_id"     # 程式自建的 Drive 資料夾 id（drive.file scope 只看得到自己建的）
BACKUP_RETRY_BACKOFF_SEC = 1800                   # 失敗後至少隔 30 分鐘再試
BACKUP_STALE_HOURS = 48                           # 超過 48 小時沒成功 → 後台標紅警告
BACKUP_TICK_SEC = 60
BACKUP_LOCAL_DIR = os.environ.get("BACKUP_LOCAL_DIR", os.path.join(_db_dir or ".", "backups"))
_backup_enabled = gdrive_backup.is_configured()

if not _backup_enabled:
    print("[backup] ⚠️ 自動備份停用：GDRIVE_CLIENT_ID / GDRIVE_CLIENT_SECRET / GDRIVE_REFRESH_TOKEN 未設齊。"
          "其餘功能不受影響；請儘快設定（tools/get_gdrive_token.py 取 refresh token）。", flush=True)


def _acquire_backup_lease():
    """搶備份租約（跨 worker）。舊值 < now - BACKUP_LEASE_SEC 才搶得到。"""
    now_ts = int(time.time())
    conn = get_db()
    try:
        conn.execute("INSERT OR IGNORE INTO admin_settings (key, value) VALUES (?, '0')", (BACKUP_LEASE_KEY,))
        cur = conn.execute(
            "UPDATE admin_settings SET value=? WHERE key=? AND CAST(value AS INTEGER) < ?",
            (str(now_ts), BACKUP_LEASE_KEY, now_ts - BACKUP_LEASE_SEC)
        )
        conn.commit()
        return cur.rowcount == 1
    finally:
        conn.close()


def _release_backup_lease():
    try:
        _set_setting(BACKUP_LEASE_KEY, "0")
    except Exception as e:
        print(f"[backup] 釋放租約失敗（{BACKUP_LEASE_SEC}s 後自動失效）: {e}", flush=True)


def _backup_is_running():
    try:
        v = int(_get_setting(BACKUP_LEASE_KEY, "0") or "0")
    except (ValueError, TypeError):
        v = 0
    return v > 0 and (time.time() - v) < BACKUP_LEASE_SEC


def _get_backup_hour():
    try:
        v = int(str(_get_setting(BACKUP_HOUR_KEY, "") or "").strip())
        if 0 <= v <= 23:
            return v
    except (ValueError, TypeError):
        pass
    return BACKUP_HOUR_DEFAULT


def _make_gdrive_client():
    cid, csec, rtok = gdrive_backup.get_env_credentials()
    return gdrive_backup.GDriveClient(cid, csec, rtok)


def run_backup_job(trigger="manual"):
    """搶租約 → 備份 → 存狀態 → 釋放租約。永遠不拋例外（備份失敗不能影響任何請求）。
    回 report；未設定回 {"success": False, "disabled": True}；搶不到租約回 {"success": False, "skipped": True}。"""
    if not gdrive_backup.is_configured():
        return {"success": False, "disabled": True, "error": "備份未設定（GDRIVE_* 環境變數未設齊）"}
    try:
        if not _acquire_backup_lease():
            return {"success": False, "skipped": True, "error": "備份進行中（另一個 worker 正在執行）"}
    except Exception as e:
        print(f"[backup] 搶租約失敗: {e}", flush=True)
        return {"success": False, "skipped": True, "error": f"租約取得失敗：{e}"}
    try:
        now = datetime.now()
        try:
            _set_setting(BACKUP_LAST_ATTEMPT_KEY, now.strftime("%Y-%m-%d %H:%M:%S"))
        except Exception:
            pass
        report = gdrive_backup.run_backup(
            DB_PATH, BACKUP_LOCAL_DIR, _make_gdrive_client(),
            stored_folder_id=_get_setting(BACKUP_FOLDER_KEY, "") or None, now=now,
        )
        report["trigger"] = trigger
        try:
            if report.get("folder_id") and report["folder_id"] != _get_setting(BACKUP_FOLDER_KEY, ""):
                _set_setting(BACKUP_FOLDER_KEY, report["folder_id"])
            if report.get("success"):
                _set_setting(BACKUP_LAST_OK_KEY, report["at"])
            else:
                _set_setting(BACKUP_LAST_FAIL_KEY, report["at"])
                _set_setting(BACKUP_LAST_FAIL_REASON_KEY, (report.get("error") or "")[:500])
            _set_setting(BACKUP_REPORT_KEY, json.dumps(report, ensure_ascii=False))
        except Exception as e:
            print(f"[backup] 存狀態失敗: {e}", flush=True)
        tag = "✅" if report.get("success") else "❌"
        print(f"[backup] {tag} ({trigger}) {json.dumps(report, ensure_ascii=False)}", flush=True)
        return report
    except Exception as e:
        print(f"[backup] ❌ 未預期例外: {e}", flush=True)
        try:
            _set_setting(BACKUP_LAST_FAIL_KEY, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            _set_setting(BACKUP_LAST_FAIL_REASON_KEY, f"未預期例外：{e}"[:500])
        except Exception:
            pass
        return {"success": False, "error": f"未預期例外：{e}"}
    finally:
        _release_backup_lease()


def _backup_due(now=None):
    """到期判斷：今天已過 backup_hour、今天還沒成功過、且距上次嘗試 ≥ 30 分鐘（失敗退避）。"""
    now = now or datetime.now()
    if now.hour < _get_backup_hour():
        return False
    today = now.strftime("%Y-%m-%d")
    if (_get_setting(BACKUP_LAST_OK_KEY, "") or "")[:10] == today:
        return False
    since_try = _seconds_since_setting_ts(BACKUP_LAST_ATTEMPT_KEY)
    if since_try is not None and since_try < BACKUP_RETRY_BACKOFF_SEC:
        return False
    return True


def _backup_tick():
    try:
        if _backup_due():
            run_backup_job(trigger="scheduled")
            return True
    except Exception as e:
        print(f"[backup] 定時任務例外: {e}", flush=True)
    return False


def _backup_loop():
    time.sleep(15)
    while True:
        _backup_tick()
        time.sleep(BACKUP_TICK_SEC)


def _start_backup_thread():
    """BACKUP_AUTO=0 可關閉（測試用）。未設定憑證則不啟動執行緒（警告已在上面印過一次）。"""
    if os.environ.get("BACKUP_AUTO", "1") != "1" or not _backup_enabled:
        return
    try:
        threading.Thread(target=_backup_loop, daemon=True, name="GDriveBackup").start()
        print(f"[backup] 每日自動備份已啟動（每日 {_get_backup_hour():02d}:00 後、每 {BACKUP_TICK_SEC}s 檢查到期；"
              f"本機副本 {BACKUP_LOCAL_DIR}）", flush=True)
    except Exception as e:
        print(f"[backup] 背景執行緒啟動失敗: {e}", flush=True)


_start_backup_thread()


def backup_status_snapshot():
    """給狀態端點與後台首頁用。stale = 超過 48 小時沒有成功備份（含從未成功）。"""
    last_ok = _get_setting(BACKUP_LAST_OK_KEY, "") or ""
    try:
        report = json.loads(_get_setting(BACKUP_REPORT_KEY, "") or "{}")
    except (ValueError, TypeError):
        report = {}
    since_ok = _seconds_since_setting_ts(BACKUP_LAST_OK_KEY)
    hours_since_ok = round(since_ok / 3600, 1) if since_ok is not None else None
    stale = (since_ok is None) or (since_ok > BACKUP_STALE_HOURS * 3600)
    try:
        local_count = len([n for n in os.listdir(BACKUP_LOCAL_DIR)
                           if n.startswith(gdrive_backup.FILE_PREFIX) and n.endswith(gdrive_backup.FILE_SUFFIX)])
    except OSError:
        local_count = 0
    return {
        "success": True,
        "enabled": gdrive_backup.is_configured(),
        "running": _backup_is_running(),
        "backup_hour": _get_backup_hour(),
        "last_ok_at": last_ok,
        "last_fail_at": _get_setting(BACKUP_LAST_FAIL_KEY, "") or "",
        "last_fail_reason": _get_setting(BACKUP_LAST_FAIL_REASON_KEY, "") or "",
        "last_attempt_at": _get_setting(BACKUP_LAST_ATTEMPT_KEY, "") or "",
        "hours_since_ok": hours_since_ok,
        "stale": stale,
        "stale_hours": BACKUP_STALE_HOURS,
        "drive_count": report.get("drive_count"),
        "latest_size": report.get("drive_latest_size") if report.get("success") else None,
        "latest_filename": report.get("filename") if report.get("success") else "",
        "local_count": local_count,
        "local_dir": BACKUP_LOCAL_DIR,
        "folder_id": _get_setting(BACKUP_FOLDER_KEY, "") or "",
        "last_report": report,
    }


# ============ 路由 ============

@app.route("/admin")
def admin_page():
    return render_template("admin.html")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config")
def get_config():
    """回傳前端所需設定（匯率等）"""
    return jsonify({
        "twd_to_jpy_rate": TWD_TO_JPY_RATE
    })


def get_admin_password():
    """取得管理員密碼：環境變數優先，否則 DB，最後預設"""
    env_pw = os.environ.get("ADMIN_PASSWORD", "")
    if env_pw:
        return env_pw
    conn = get_db()
    row = conn.execute("SELECT value FROM admin_settings WHERE key='admin_password'").fetchone()
    conn.close()
    if row:
        return row["value"]
    return "admin123"


def _ensure_super_admin():
    """確保至少有一個超級管理員"""
    try:
        conn = get_db()
        count = conn.execute("SELECT COUNT(*) as c FROM admin_users").fetchone()["c"]
        if count == 0:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pwd = get_admin_password()
            conn.execute(
                "INSERT INTO admin_users (username, password, role, created_at) VALUES (?, ?, 'super', ?)",
                ("admin", pwd, now)
            )
            conn.commit()
            print(f"[Admin] ✅ 已建立超級管理員帳號: admin / {pwd}", flush=True)
        else:
            env_pw = os.environ.get("ADMIN_PASSWORD", "")
            if env_pw:
                # ⚠️ 只重設 bootstrap 帳號「admin」的密碼。
                # 舊寫法是 WHERE role='super'，會把「所有老闆」的密碼一起蓋成同一組，
                # 導致其他老闆（例如 Flower）每次容器重啟後密碼被改掉而登不進來。
                conn.execute("UPDATE admin_users SET password=? WHERE username='admin'", (env_pw,))
                conn.commit()
            # 純密碼登入：密碼重複會讓系統認錯人，啟動時檢查並警告（不印出密碼本身）
            dups = conn.execute(
                "SELECT COUNT(*) AS c, GROUP_CONCAT(username) AS us FROM admin_users "
                "GROUP BY password HAVING c > 1"
            ).fetchall()
            for d in dups:
                print(f"[Admin] ⚠️ 密碼重複：{d['us']} 使用同一組密碼，請到後台改成不同密碼", flush=True)
            print(f"[Admin] ✅ 已有 {count} 個管理員帳號", flush=True)
        conn.close()
    except Exception as e:
        print(f"[Admin] ❌ 初始化失敗: {e}", flush=True)

_ensure_super_admin()
print("[App] ✅ 啟動完成", flush=True)


@app.route("/api/admin/verify", methods=["POST"])
def admin_verify():
    data = request.json
    username = (data.get("username") or "").strip()
    password = data.get("password", "")
    print(f"[Login] 嘗試登入: username='{username}'", flush=True)

    _ip = _client_ip()
    _lock_key = _admin_login_key(username)
    if _admin_login_locked(_lock_key, _ip):
        print(f"[admin_login_lock] key={_lock_key} ip={_ip}", flush=True)
        return jsonify({"success": False,
                        "error": "登入嘗試次數過多，請 15 分鐘後再試"}), 429

    conn = get_db()
    user = None
    user_type = None

    # 1) 先查管理員（admin_users）— 既有行為不變
    if username:
        user = conn.execute(
            "SELECT * FROM admin_users WHERE username=? AND password=?", (username, password)
        ).fetchone()
    else:
        # 相容舊的純密碼登入
        user = conn.execute(
            "SELECT * FROM admin_users WHERE password=?", (password,)
        ).fetchone()
    if user:
        user_type = "admin"

    # 2) admin 找不到 → 再查代理（agents）；代理模組關閉時不允許代理登入
    if not user and username and ENABLE_AGENTS:
        user = conn.execute(
            "SELECT * FROM agents WHERE username=? AND password=? AND status='active'",
            (username, password)
        ).fetchone()
        if user:
            user_type = "agent"
    conn.close()

    if user:
        _login_record(_lock_key, _ip, True)   # 成功 → 清掉該 key 視窗內的失敗紀錄
        # 寫入 session
        session.permanent = True
        session["user_type"] = user_type
        session["user_id"] = user["id"]
        session["username"] = user["username"]
        if user_type == "admin":
            session["role"] = user["role"]
            session["agent_id"] = 0  # 0 = 主管理員 / 你的員工，看全部
            print(f"[Login] ✅ admin 登入: {user['username']} ({user['role']})", flush=True)
            return jsonify({
                "success": True,
                "user_type": "admin",
                "user_id": user["id"],
                "username": user["username"],
                "role": user["role"]
            })
        else:
            session["role"] = "agent"
            session["agent_id"] = user["id"]
            session["prefix"] = user["prefix"]
            print(f"[Login] ✅ agent 登入: {user['username']} (prefix={user['prefix']}, id={user['id']})", flush=True)
            return jsonify({
                "success": True,
                "user_type": "agent",
                "username": user["username"],
                "name": user["name"],
                "prefix": user["prefix"]
            })

    print(f"[Login] ❌ 登入失敗", flush=True)
    _login_record(_lock_key, _ip, False)
    return jsonify({"success": False, "error": "帳號或密碼錯誤"})


# ===== 身份輔助函式 =====
def current_user():
    """回傳當前登入者資訊（從 session）"""
    if "user_type" not in session:
        return None
    return {
        "user_type": session.get("user_type"),
        "user_id": session.get("user_id"),
        "username": session.get("username"),
        "role": session.get("role"),
        "agent_id": session.get("agent_id", 0),
        "prefix": session.get("prefix", "G"),
    }

def is_super_admin():
    """是否為管理員（admin_users 表的人＝老闆或員工，日常作業都可）"""
    return session.get("user_type") == "admin"

# ── 登入失敗速率限制（/api/verify_customer）──
LOGIN_WINDOW_SEC = 900            # 視窗 15 分鐘
LOGIN_MAX_FAIL_GCODE = 5          # 同一客編視窗內失敗上限
LOGIN_MAX_FAIL_IP = 20            # 同一 IP 視窗內失敗上限
LOGIN_ATTEMPTS_RETENTION = 86400  # 紀錄保留 24 小時


def _client_ip():
    """Zeabur 在 proxy 後面，優先取 X-Forwarded-For 第一段。
    取不到就回空字串（代表無法辨識來源）。

    ⚠️ 不 fallback 到 request.remote_addr：在 proxy 後面那是 proxy 自己的位址，
    全站客戶會共用同一個值，IP 規則會把所有人一起鎖死。
    取不到就讓 IP 規則失效，只靠 g_code 規則擋 —— 寧可少擋，不可誤鎖全站。"""
    xff = (request.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
    return xff or ""


def _login_locked(g_code, ip):
    """視窗內失敗次數是否已達上限。DB 出狀況時 fail-open（寧可少擋，不可鎖死全站）。"""
    since = int(time.time()) - LOGIN_WINDOW_SEC
    try:
        conn = get_db()
        n = conn.execute(
            "SELECT COUNT(*) AS c FROM login_attempts "
            "WHERE success=0 AND g_code=? AND created_ts>=?", (g_code, since)
        ).fetchone()["c"]
        if n >= LOGIN_MAX_FAIL_GCODE:
            conn.close()
            return True
        if ip:   # ip 為空字串（取不到 XFF）→ 整條 IP 規則跳過
            n_ip = conn.execute(
                "SELECT COUNT(*) AS c FROM login_attempts "
                "WHERE success=0 AND ip=? AND created_ts>=?", (ip, since)
            ).fetchone()["c"]
            if n_ip >= LOGIN_MAX_FAIL_IP:
                conn.close()
                return True
        conn.close()
    except Exception as e:
        print(f"[login_lock] 檢查失敗（放行）: {e}", flush=True)
    return False


def _login_record(g_code, ip, success):
    """記一筆登入嘗試；成功時清掉該客編視窗內的失敗紀錄（只清自己的）。"""
    now = int(time.time())
    try:
        conn = get_db()
        conn.execute(
            "INSERT INTO login_attempts (g_code, ip, success, created_ts) VALUES (?, ?, ?, ?)",
            (g_code, ip, 1 if success else 0, now)
        )
        if success:
            conn.execute(
                "DELETE FROM login_attempts WHERE g_code=? AND success=0 AND created_ts>=?",
                (g_code, now - LOGIN_WINDOW_SEC)
            )
        # 順手清理過期紀錄
        conn.execute("DELETE FROM login_attempts WHERE created_ts < ?", (now - LOGIN_ATTEMPTS_RETENTION,))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[login_lock] 記錄失敗（不影響登入）: {e}", flush=True)


def _login_fail(g_code, ip, error, **extra):
    """記一次失敗並回原本的錯誤訊息（訊息維持原樣，不洩漏剩餘次數）。"""
    _login_record(g_code, ip, False)
    return jsonify({"success": False, "error": error, **extra})


# ── 後台登入 / 改密碼的速率限制 ──────────────────────────────
# 沿用客戶登入那張 login_attempts 表，key 加 "ADMIN:" 前綴與客編區隔
# （g_code 只有英數，不可能撞到冒號）。
#
# 為什麼後台需要：/api/admin/verify 與 /api/admin/change_password 原本完全沒有
# 次數限制。後者要求填對「目前密碼」才換，看似安全，但它等於一個可以無限次
# 測試密碼的 oracle，配合未授權的 /api/admin/users（可列出所有帳號）就是
# 一條完整的暴力破解路徑。兩支共用同一個 key，攻擊者不能拿改密碼那支
# 來繞過登入那支的計數。
#
# ⚠️ 純密碼登入（沒帶 username）不做「per 帳號」鎖定，只套 IP 規則：
#    那時還不知道對方是誰，共用一個計數器等於任何人連打幾次就能把老闆
#    鎖在自己的系統外面。營運時段擋掉自己人的傷害大於擋住攻擊者的好處。
#    帶 username 的登入有 per 帳號鎖定，而純密碼登入這條路徑始終留著，
#    剛好成為帳號被鎖時的安全閥。
ADMIN_LOGIN_MAX_FAIL = 8      # 同一後台帳號視窗（15 分鐘）內失敗上限
ADMIN_ANON_KEY = "ADMIN:"     # 純密碼登入：只記錄（讓 IP 規則看得到），不做 per 帳號鎖定


def _admin_login_key(username):
    u = (username or "").strip().lower()
    return ADMIN_ANON_KEY + u if u else ADMIN_ANON_KEY


def _admin_login_locked(key, ip):
    """後台登入/改密碼是否已達失敗上限。DB 出狀況時 fail-open（同 _login_locked）。"""
    since = int(time.time()) - LOGIN_WINDOW_SEC
    try:
        conn = get_db()
        try:
            if key and key != ADMIN_ANON_KEY:
                n = conn.execute(
                    "SELECT COUNT(*) AS c FROM login_attempts "
                    "WHERE success=0 AND g_code=? AND created_ts>=?", (key, since)
                ).fetchone()["c"]
                if n >= ADMIN_LOGIN_MAX_FAIL:
                    return True
            if ip:   # ip 為空字串（取不到 XFF）→ 整條 IP 規則跳過
                n_ip = conn.execute(
                    "SELECT COUNT(*) AS c FROM login_attempts "
                    "WHERE success=0 AND ip=? AND created_ts>=?", (ip, since)
                ).fetchone()["c"]
                if n_ip >= LOGIN_MAX_FAIL_IP:
                    return True
        finally:
            conn.close()
    except Exception as e:
        print(f"[admin_login_lock] 檢查失敗（放行）: {e}", flush=True)
    return False


def _require_customer(g_code):
    """驗證請求的 g_code 是否為目前登入者。
    回傳 (ok: bool, resp)。ok=False 時直接 return resp。
    後台 admin session 一律放行（後台有代客戶操作的流程）。

    背景：客戶端 API 原本只拿 request 傳來的 g_code 比對資料列的 g_code，
    客編連號（G0001…），一支迴圈就能把全站客戶的地址簿、申報人、包裹、帳單撈走。
    真正的帳密驗證在 /api/verify_customer，這裡把它的結果（session）接上來。"""
    if session.get("user_type") == "admin":
        return True, None
    cur = (session.get("cust_g_code") or "").upper()
    if not cur:
        return False, (jsonify({"success": False, "need_login": True,
                                "error": "登入階段已過期，請重新登入"}), 401)
    if cur != (g_code or "").strip().upper():
        return False, (jsonify({"success": False, "error": "無權存取"}), 403)
    return True, None


def is_boss():
    """是否為老闆（super）：可看營收統計、代理管理、管理員管理、變更密碼等敏感功能"""
    return session.get("user_type") == "admin" and session.get("role") == "super"

def is_staff():
    """是否為員工（admin_users 但非 super）：只能日常作業"""
    return session.get("user_type") == "admin" and session.get("role") != "super"

def current_operator():
    """當前操作者顯示名稱（給操作紀錄用）"""
    return session.get("username") or "?"

def log_op(action, target="", detail=""):
    """記一筆操作紀錄（誰、做了什麼、對象）。失敗不影響主流程。"""
    try:
        conn = get_db()
        conn.execute(
            "INSERT INTO operation_logs (operator, role, action, target, detail, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (current_operator(), session.get("role", ""), action, str(target), str(detail),
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[log_op] 失敗: {e}", flush=True)

def get_current_agent_id():
    """當前代理 id（>0 才是代理，0 = 主管理員看全部）"""
    return int(session.get("agent_id", 0))


def _parse_boxes(boxes_json):
    """安全解析 boxes_json → list（失敗回空）。"""
    if not boxes_json:
        return []
    try:
        v = json.loads(boxes_json)
        return v if isinstance(v, list) else []
    except (ValueError, TypeError):
        return []


def _safe_str(v):
    """安全把 None / 數字 / 字串 都轉為 strip 過的字串。

    主要解決 DB 內有些 phone 欄位被存成 float（912345678.0）的問題。
    無論進來是 float、字串 "912345678.0"、還是已是純字串，都正規化為合理形式：
      • None              → ""
      • 912345678.0       → "912345678"   （整數 float 去掉 .0）
      • "912345678.0"     → "912345678"   （SQLite 存進 TEXT 欄位後變字串）
      • "0912345678"      → "0912345678"  （原樣保留）
      • 3.14              → "3.14"        （非整數 float 保留）
    """
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v)).strip()
    s = str(v).strip()
    # SQLite TEXT 欄位存進 float 後變字串 "912345678.0" → 去掉 .0
    if s.endswith(".0"):
        prefix = s[:-2]
        if prefix.lstrip("-").isdigit():
            return prefix
    return s


def _parse_pkg_ids(raw):
    """容錯解析 package_ids 字串 → [int]（保序、不去重，與舊 list comprehension 行為一致）。

    支援這些髒格式：
      • "5,8,12"        → [5, 8, 12]   （正常）
      • "5.0,8.0,12.0"  → [5, 8, 12]   （migration 把整數 floatify；舊版用 .isdigit() 會整批解析成空）
      • " 5 , 8 "       → [5, 8]       （多餘空白）
      • "5，8、12"       → [5, 8, 12]   （全形逗號／頓號／空白分隔）
      • None / ""       → []
    只接受純整數或「整數.000」格式；"5.7" / "abc" / "-3" / "1e3" 一律忽略，
    避免把壞資料硬轉成錯誤 ID。
    """
    if raw is None:
        return []
    out = []
    for tok in re.split(r"[,，、\s]+", str(raw)):
        tok = tok.strip()
        if not tok:
            continue
        # 純整數，或整數後接全 0 的小數（"5"、"5.0"、"5.00"）
        if re.fullmatch(r"\d+(?:\.0+)?", tok):
            n = int(float(tok))
            if n > 0:
                out.append(n)
    return out


@app.route("/api/me", methods=["GET"])
def api_me():
    """前端查當前身份（只讀自己的 session，不吃 g_code → 無跨帳號存取問題）"""
    cust = session.get("cust_g_code") or ""
    u = current_user()
    if not u:
        return jsonify({"logged_in": bool(cust), "cust_g_code": cust})
    return jsonify({"logged_in": True, "cust_g_code": cust, **u})


@app.route("/api/admin/logout", methods=["POST"])
def admin_logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/api/admin/change_password", methods=["POST"])
def admin_change_password():
    # 需登入。原本未登入也能呼叫，只要填對「目前密碼」就換得掉。
    # 速率限制（下面那段）擋的是無限試誤，但擋不住「本來就拿到密碼的人」
    # 在沒有任何 session 的情況下直接改掉它 —— 例如密碼外流後被搶先改走鎖死帳號。
    # 前端只從已登入的後台面板呼叫（admin.html submitChangePwd），補上不影響流程。
    #
    # 順序刻意放在速率限制之前：未登入的請求直接 403，不進計數器。
    # 否則匿名者可以靠狂打這支去灌爆共用的 IP 額度，把真的使用者鎖在外面。
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    data = request.json
    username = (data.get("username") or "admin").strip()
    current = data.get("current", "")
    new_pwd = data.get("new_password", "").strip()
    confirm = data.get("confirm", "").strip()

    # 與 /api/admin/verify 共用同一把 key：這支要求填對「目前密碼」才換，
    # 沒有次數限制就是一個無限試誤的 oracle。共用 key 才不會被拿來繞過登入的計數。
    _ip = _client_ip()
    _lock_key = _admin_login_key(username)
    if _admin_login_locked(_lock_key, _ip):
        print(f"[admin_login_lock] change_password key={_lock_key} ip={_ip}", flush=True)
        return jsonify({"success": False,
                        "error": "嘗試次數過多，請 15 分鐘後再試"}), 429

    conn = get_db()
    user = conn.execute("SELECT * FROM admin_users WHERE username=?", (username,)).fetchone()
    if not user or user["password"] != current:
        conn.close()
        _login_record(_lock_key, _ip, False)
        return jsonify({"success": False, "error": "目前密碼錯誤"})
    if not new_pwd or len(new_pwd) < 4:
        conn.close()
        return jsonify({"success": False, "error": "新密碼至少 4 個字元"})
    if new_pwd != confirm:
        conn.close()
        return jsonify({"success": False, "error": "兩次密碼不一致"})

    conn.execute("UPDATE admin_users SET password=? WHERE username=?", (new_pwd, username))
    conn.commit()
    conn.close()
    _login_record(_lock_key, _ip, True)   # 換成功 → 清掉該 key 視窗內的失敗紀錄
    return jsonify({"success": True, "message": "密碼已更新"})


# ── 管理員帳號管理 ──

@app.route("/api/admin/users", methods=["GET"])
def admin_list_users():
    # 原本完全沒有身分檢查：未登入可列出所有後台帳號的 username 與 role
    # （誰是 super）。後台登入支援「純密碼登入」，這份名單等於餵給暴力破解的字典。
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    conn = get_db()
    rows = conn.execute("SELECT id, username, role, created_at FROM admin_users ORDER BY id").fetchall()
    conn.close()
    return jsonify({"success": True, "users": [dict(r) for r in rows]})


# ===== 代理帳號管理（只有主管理員可操作）=====

@app.route("/api/admin/agents", methods=["GET"])
def admin_list_agents():
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    rows = conn.execute(
        "SELECT id, username, prefix, name, min_rate, contact_phone, contact_email, status, note, created_at,"
        " contact_line, insurance_url, insurance_label, insurance_desc, signup_guide, promo_text, promo_price,"
        " owner_name, owner_address,"
        " bank_code, bank_name, bank_branch, bank_account, bank_account_name"
        " FROM agents ORDER BY id"
    ).fetchall()
    # 順便統計每個代理底下的會員數
    counts = {}
    for r in conn.execute("SELECT agent_id, COUNT(*) as c FROM members GROUP BY agent_id").fetchall():
        counts[r["agent_id"]] = r["c"]
    conn.close()
    # 每個代理的分潤摘要（各週明細 + 未撥款總額），讓你不必登入代理帳號就看得到
    result = []
    for r in rows:
        d = dict(r)
        d["member_count"] = counts.get(r["id"], 0)
        weeks = compute_agent_weekly(r["id"])
        d["weeks"] = weeks
        d["unpaid_total"] = round(sum(w["commission"] for w in weeks if not w["paid"]))
        d["paid_total"] = round(sum(w["commission"] for w in weeks if w["paid"]))
        result.append(d)
    return jsonify({"success": True, "agents": result})


# ===== 代理分潤撥款 =====

@app.route("/api/admin/agents/<int:agent_id>/payout", methods=["POST"])
def admin_agent_payout(agent_id):
    """標記某代理某週已撥款（填匯款後五碼 + 時間）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    period_key = (data.get("period_key") or "").strip()
    last5 = (data.get("payment_last5") or "").strip()
    note = (data.get("note") or "").strip()
    if not period_key:
        return jsonify({"success": False, "error": "缺少週次"}), 400
    if not last5:
        return jsonify({"success": False, "error": "請填匯款後五碼"}), 400

    # 金額以系統計算為準（避免前端竄改）
    weeks = compute_agent_weekly(agent_id)
    amount = next((w["commission"] for w in weeks if w["period_key"] == period_key), None)
    if amount is None:
        return jsonify({"success": False, "error": "該週無分潤資料"}), 400

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    paid_at = (data.get("paid_at") or "").strip() or now
    conn = get_db()
    conn.execute(
        "INSERT INTO agent_payouts (agent_id, period_key, amount, payment_last5, paid_at, note, created_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?) "
        "ON CONFLICT(agent_id, period_key) DO UPDATE SET "
        "amount=excluded.amount, payment_last5=excluded.payment_last5, paid_at=excluded.paid_at, note=excluded.note",
        (agent_id, period_key, amount, last5, paid_at, note, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "amount": amount, "paid_at": paid_at})


@app.route("/api/admin/agents/<int:agent_id>/payout", methods=["DELETE"])
def admin_agent_payout_cancel(agent_id):
    """取消某週撥款標記。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    period_key = (request.args.get("period_key") or "").strip()
    if not period_key:
        return jsonify({"success": False, "error": "缺少週次"}), 400
    conn = get_db()
    conn.execute("DELETE FROM agent_payouts WHERE agent_id=? AND period_key=?", (agent_id, period_key))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/agent/payouts", methods=["GET"])
def agent_my_payouts():
    """代理端：看自己的每週分潤與撥款狀態。"""
    aid = get_current_agent_id()
    if aid <= 0:
        return jsonify({"success": False, "error": "僅代理帳號可查看"}), 403
    return jsonify({"success": True, "weeks": compute_agent_weekly(aid)})


# ===== 代理品牌資訊（公開：referral URL 使用）=====
def _branding_dict(agent_row=None):
    """組合品牌資料：有 agent → 該代理；無 agent → 你的預設"""
    if agent_row:
        d = dict(agent_row)
        return {
            "is_agent": True,
            "agent_prefix": d.get("prefix", ""),
            "display_name": d.get("name") or BRAND["name"],
            "contact_line": d.get("contact_line") or "",
            "insurance_url": d.get("insurance_url") or "",
            "insurance_label": d.get("insurance_label") or "",
            "insurance_desc": d.get("insurance_desc") or "",
            "signup_guide": d.get("signup_guide") or "",
            "promo_text": d.get("promo_text") or "",
            "promo_price": d.get("promo_price") or "",
        }
    # 預設（FWT JAPAN 品牌，見 brand.py）
    return {
        "is_agent": False,
        "agent_prefix": "G",
        "display_name": BRAND["name"],
        "contact_line": BRAND["line_id"],
        "insurance_url": BRAND["insurance_url"],
        "insurance_label": BRAND["insurance_label"],
        "insurance_desc": BRAND["insurance_desc"],
        "signup_guide": "",  # 預設由前端原本的內容處理
        "promo_text": "",   # 預設徽章由前端寫死（限時特價招生 NT$200）
        "promo_price": "",
    }


@app.route("/api/branding", methods=["GET"])
def api_branding():
    """
    公開 API：依 ?a=PREFIX 回傳對應代理的品牌資訊。
    無 a 或找不到 → 回預設（你的 FWT JAPAN 內容）
    """
    prefix = (request.args.get("a") or "").strip().upper()
    if not prefix or not ENABLE_AGENTS:
        return jsonify({"success": True, **_branding_dict(None)})
    try:
        conn = get_db()
        row = conn.execute(
            "SELECT * FROM agents WHERE prefix=? AND status='active'", (prefix,)
        ).fetchone()
        conn.close()
        return jsonify({"success": True, **_branding_dict(row)})
    except Exception as e:
        print(f"[api_branding] {e}", flush=True)
        return jsonify({"success": True, **_branding_dict(None)})


@app.route("/api/agent/my_branding", methods=["GET", "PUT"])
def agent_my_branding():
    """代理自助：查看與編輯自己的品牌設定"""
    aid = get_current_agent_id()
    if aid <= 0:
        return jsonify({"success": False, "error": "僅代理可使用此功能"}), 403
    conn = get_db()
    if request.method == "GET":
        row = conn.execute("SELECT * FROM agents WHERE id=?", (aid,)).fetchone()
        conn.close()
        if not row:
            return jsonify({"success": False, "error": "代理資料異常"}), 500
        d = dict(row)
        return jsonify({
            "success": True,
            "agent": {
                "id": d["id"],
                "name": d.get("name", ""),
                "prefix": d.get("prefix", ""),
                "min_rate": d.get("min_rate", 0),
                "contact_phone": d.get("contact_phone", ""),
                "contact_email": d.get("contact_email", ""),
                "contact_line": d.get("contact_line", ""),
                "insurance_url": d.get("insurance_url", ""),
                "insurance_label": d.get("insurance_label", ""),
                "insurance_desc": d.get("insurance_desc", ""),
                "signup_guide": d.get("signup_guide", ""),
                "promo_text": d.get("promo_text", ""),
                "promo_price": d.get("promo_price", ""),
                "owner_name": d.get("owner_name", ""),
                "owner_address": d.get("owner_address", ""),
                "bank_code": d.get("bank_code", ""),
                "bank_name": d.get("bank_name", ""),
                "bank_branch": d.get("bank_branch", ""),
                "bank_account": d.get("bank_account", ""),
                "bank_account_name": d.get("bank_account_name", ""),
            }
        })
    # PUT：更新自己的品牌欄位（不能改帳號、前綴、密碼、狀態、min_rate）
    data = request.json or {}
    fields, values = [], []
    # 允許代理自己改的欄位：聯絡方式 + 品牌 + 特價徽章 + 負責人資訊 + 銀行帳戶
    for col in ["name", "contact_phone", "contact_email", "contact_line",
                "insurance_url", "insurance_label", "insurance_desc", "signup_guide",
                "promo_text", "promo_price",
                "owner_name", "owner_address",
                "bank_code", "bank_name", "bank_branch", "bank_account", "bank_account_name"]:
        if col in data:
            fields.append(f"{col}=?"); values.append((data[col] or "").strip())
    # min_rate 開放代理自設費率
    if "min_rate" in data:
        try:
            mr = float(data["min_rate"])
            fields.append("min_rate=?"); values.append(mr)
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": "費率必須為數字"})
    if not fields:
        conn.close()
        return jsonify({"success": False, "error": "沒有可更新欄位"})
    values.append(aid)
    conn.execute(f"UPDATE agents SET {', '.join(fields)} WHERE id=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "已儲存"})


@app.route("/api/admin/agents", methods=["POST"])
def admin_create_agent():
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    prefix = (data.get("prefix") or "").strip().upper()
    name = (data.get("name") or "").strip()
    min_rate = float(data.get("min_rate") or 180)

    if not username or not password or not prefix or not name:
        return jsonify({"success": False, "error": "帳號、密碼、前綴、名稱皆為必填"})
    if not re.fullmatch(r"[A-Z]", prefix):
        return jsonify({"success": False, "error": "前綴必須為單一英文字母（A-Z）"})
    if prefix == "G":
        return jsonify({"success": False, "error": "前綴 G 已保留給主管理員"})
    # min_rate 已開放代理自由設定（不再有 180 下限）

    conn = get_db()
    try:
        conn.execute(
            """INSERT INTO agents (username, password, prefix, name, min_rate, contact_phone, contact_email,
                                   status, note, created_at, contact_line, insurance_url, insurance_label, insurance_desc, signup_guide,
                                   promo_text, promo_price, owner_name, owner_address)
               VALUES (?, ?, ?, ?, ?, ?, ?, 'active', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (username, password, prefix, name, min_rate,
             data.get("contact_phone", ""), data.get("contact_email", ""),
             data.get("note", ""), datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             data.get("contact_line", ""), data.get("insurance_url", ""),
             data.get("insurance_label", ""), data.get("insurance_desc", ""),
             data.get("signup_guide", ""),
             data.get("promo_text", ""), data.get("promo_price", ""),
             data.get("owner_name", ""), data.get("owner_address", ""))
        )
        conn.commit()
        new_id = conn.execute("SELECT last_insert_rowid() as id").fetchone()["id"]
    except sqlite3.IntegrityError as e:
        conn.close()
        msg = str(e)
        if "agents.username" in msg:
            return jsonify({"success": False, "error": f"帳號「{username}」已被使用"})
        if "agents.prefix" in msg:
            return jsonify({"success": False, "error": f"前綴「{prefix}」已被使用"})
        return jsonify({"success": False, "error": f"資料庫錯誤：{msg}"})
    conn.close()
    return jsonify({"success": True, "id": new_id, "message": f"代理「{name}」已建立（前綴 {prefix}）"})


@app.route("/api/admin/agents/<int:agent_id>", methods=["PUT"])
def admin_update_agent(agent_id):
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    conn = get_db()
    existing = conn.execute("SELECT * FROM agents WHERE id=?", (agent_id,)).fetchone()
    if not existing:
        conn.close()
        return jsonify({"success": False, "error": "代理不存在"})

    fields = []
    values = []
    # 可改：name, min_rate, contact_phone, contact_email, status, note, password
    if "name" in data:
        fields.append("name=?"); values.append((data["name"] or "").strip())
    if "min_rate" in data:
        try:
            mr = float(data["min_rate"])
            fields.append("min_rate=?"); values.append(mr)
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": "費率必須為數字"})
    if "contact_phone" in data:
        fields.append("contact_phone=?"); values.append(data["contact_phone"] or "")
    if "contact_email" in data:
        fields.append("contact_email=?"); values.append(data["contact_email"] or "")
    if "status" in data and data["status"] in ("active", "disabled"):
        fields.append("status=?"); values.append(data["status"])
    if "note" in data:
        fields.append("note=?"); values.append(data["note"] or "")
    if data.get("password"):
        fields.append("password=?"); values.append(data["password"])
    # 品牌欄位（referral URL + 登入後內容客製） + 銀行帳戶（撥款用）
    for col in ["contact_line", "insurance_url", "insurance_label", "insurance_desc", "signup_guide",
                "promo_text", "promo_price", "owner_name", "owner_address",
                "bank_code", "bank_name", "bank_branch", "bank_account", "bank_account_name"]:
        if col in data:
            fields.append(f"{col}=?"); values.append((data[col] or "").strip())
    # 前綴與帳號名建立後不可改（避免關聯混亂）

    if not fields:
        conn.close()
        return jsonify({"success": False, "error": "沒有可更新的欄位"})
    values.append(agent_id)
    conn.execute(f"UPDATE agents SET {', '.join(fields)} WHERE id=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "已更新"})


@app.route("/api/admin/agents/<int:agent_id>", methods=["DELETE"])
def admin_delete_agent(agent_id):
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    transfer = request.args.get("transfer") == "1"
    conn = get_db()
    # 統計關聯資料
    cm = conn.execute("SELECT COUNT(*) as c FROM members WHERE agent_id=?", (agent_id,)).fetchone()["c"]
    cp = conn.execute("SELECT COUNT(*) as c FROM packages WHERE agent_id=?", (agent_id,)).fetchone()["c"]
    cf = conn.execute("SELECT COUNT(*) as c FROM forecasts WHERE agent_id=?", (agent_id,)).fetchone()["c"]
    cs = conn.execute("SELECT COUNT(*) as c FROM shipment_requests WHERE agent_id=?", (agent_id,)).fetchone()["c"]
    has_data = (cm + cp + cf + cs) > 0
    if has_data and not transfer:
        conn.close()
        return jsonify({
            "success": False,
            "needs_transfer": True,
            "stats": {"members": cm, "packages": cp, "forecasts": cf, "shipment_requests": cs},
            "error": f"此代理底下尚有 {cm} 位會員 / {cp} 個包裹 / {cf} 個預報 / {cs} 個出貨紀錄。請改用「離職移交」將資料轉回主管理員。"
        })
    if has_data and transfer:
        # 全部 agent_id 改為 0 （= 主管理員 / 你）
        conn.execute("UPDATE members SET agent_id=0 WHERE agent_id=?", (agent_id,))
        conn.execute("UPDATE packages SET agent_id=0 WHERE agent_id=?", (agent_id,))
        conn.execute("UPDATE forecasts SET agent_id=0 WHERE agent_id=?", (agent_id,))
        conn.execute("UPDATE shipment_requests SET agent_id=0 WHERE agent_id=?", (agent_id,))
        print(f"[transfer] agent_id={agent_id} 移交 {cm} 會員 / {cp} 包裹 / {cf} 預報 / {cs} 出貨 給主管理員", flush=True)
    conn.execute("DELETE FROM agents WHERE id=?", (agent_id,))
    conn.commit()
    conn.close()
    if has_data and transfer:
        return jsonify({
            "success": True,
            "transferred": {"members": cm, "packages": cp, "forecasts": cf, "shipment_requests": cs},
            "message": f"已離職移交：{cm} 位會員、{cp} 個包裹、{cf} 個預報、{cs} 個出貨紀錄已轉回主管理員。會員編號保留不變。"
        })
    return jsonify({"success": True, "message": "代理已刪除"})


# ===== 統一會員查詢（本地 members 優先、找不到回退 Shopify）=====
def get_agent_id_for_g_code(g_code):
    """
    依 g_code 找出歸屬的 agent_id。
    - 在 members 表（代理的客戶）→ 回那個代理的 id
    - 不在 members 表（你 Shopify 來的客戶）→ 回 0（主管理員）
    """
    if not g_code:
        return 0
    try:
        conn = get_db()
        row = conn.execute("SELECT agent_id FROM members WHERE g_code=?", (g_code,)).fetchone()
        conn.close()
        if row:
            return int(row["agent_id"] or 0)
    except Exception as e:
        print(f"[get_agent_id_for_g_code] 失敗: {e}", flush=True)
    return 0


def check_record_ownership(table, record_id):
    """
    檢查當前使用者是否可存取該筆紀錄。
    回傳 (allowed: bool, record_dict_or_none)
    - 主管理員：永遠可以
    - 代理：只有當 record.agent_id == 自己的 agent_id 才可以
    """
    aid = get_current_agent_id()
    conn = get_db()
    row = conn.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
    conn.close()
    if not row:
        return False, None
    if aid == 0 or is_super_admin():
        return True, dict(row)
    return (int(row["agent_id"] or 0) == aid), dict(row)


def get_member_unified(g_code):
    """
    回傳 {g_code, name, agent_id, phone, address, source} 或 None
    - source='local'  → 來自代理建的會員（agent_id > 0）
    - source='shopify' → 來自你的 Shopify（agent_id = 0）
    """
    if not g_code:
        return None
    conn = get_db()
    row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
    conn.close()
    if row:
        d = dict(row)
        d["source"] = "local"
        return d
    # 回退到 Shopify 快取
    try:
        for c in get_all_goyoutati_customers():
            if (c.get("g_code") or "") == g_code:
                return {
                    "g_code": g_code,
                    "name": c.get("name", ""),
                    "phone": c.get("phone", ""),
                    "address": c.get("address", ""),
                    "agent_id": 0,
                    "source": "shopify"
                }
    except Exception as e:
        print(f"[get_member_unified] Shopify 查詢失敗: {e}", flush=True)
    return None


@app.route("/api/admin/users", methods=["POST"])
def admin_create_user():
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    role = data.get("role", "staff")
    if role not in ("super", "staff"):
        role = "staff"
    if not username or not password:
        return jsonify({"success": False, "error": "帳號和密碼為必填"})
    if len(password) < 4:
        return jsonify({"success": False, "error": "密碼至少 4 個字元"})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    # 純密碼登入 → 密碼必須全站唯一（否則分不出是誰）
    dup = conn.execute("SELECT username FROM admin_users WHERE password=?", (password,)).fetchone()
    if dup:
        conn.close()
        return jsonify({"success": False, "error": f"此密碼已被「{dup['username']}」使用，請換一組（密碼須唯一）"})
    try:
        conn.execute("INSERT INTO admin_users (username, password, role, created_at) VALUES (?, ?, ?, ?)",
                     (username, password, role, now))
        conn.commit()
        conn.close()
        return jsonify({"success": True})
    except:
        conn.close()
        return jsonify({"success": False, "error": "帳號已存在"})

@app.route("/api/admin/users/<int:user_id>", methods=["DELETE"])
def admin_delete_user(user_id):
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    user = conn.execute("SELECT role FROM admin_users WHERE id=?", (user_id,)).fetchone()
    if not user:
        conn.close()
        return jsonify({"success": False, "error": "找不到"})
    if user["role"] == "super":
        conn.close()
        return jsonify({"success": False, "error": "無法刪除超級管理員"})
    conn.execute("DELETE FROM admin_users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/users/<int:user_id>/password", methods=["POST"])
def admin_reset_user_password(user_id):
    """老闆變更任一管理員的密碼（純密碼登入 → 唯一檢查）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    new_pwd = ((request.json or {}).get("password") or "").strip()
    if len(new_pwd) < 4:
        return jsonify({"success": False, "error": "密碼至少 4 個字元"})
    conn = get_db()
    dup = conn.execute("SELECT username FROM admin_users WHERE password=? AND id!=?", (new_pwd, user_id)).fetchone()
    if dup:
        conn.close()
        return jsonify({"success": False, "error": f"此密碼已被「{dup['username']}」使用，請換一組"})
    u = conn.execute("SELECT username FROM admin_users WHERE id=?", (user_id,)).fetchone()
    if not u:
        conn.close()
        return jsonify({"success": False, "error": "找不到此管理員"})
    conn.execute("UPDATE admin_users SET password=? WHERE id=?", (new_pwd, user_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/users/<int:user_id>/role", methods=["POST"])
def admin_change_user_role(user_id):
    """變更管理員身分（老闆 super ⇄ 員工 staff）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    new_role = ((request.json or {}).get("role") or "").strip()
    if new_role not in ("super", "staff"):
        return jsonify({"success": False, "error": "身分只能是 super 或 staff"}), 400
    if session.get("user_id") == user_id:
        return jsonify({"success": False, "error": "不能變更自己的身分"}), 400
    conn = get_db()
    u = conn.execute("SELECT username, role FROM admin_users WHERE id=?", (user_id,)).fetchone()
    if not u:
        conn.close()
        return jsonify({"success": False, "error": "找不到此管理員"}), 404
    # 保護：不可把最後一個老闆降為員工
    if u["role"] == "super" and new_role == "staff":
        supers = conn.execute("SELECT COUNT(*) AS c FROM admin_users WHERE role='super'").fetchone()["c"]
        if supers <= 1:
            conn.close()
            return jsonify({"success": False, "error": "至少要保留一位老闆"}), 400
    conn.execute("UPDATE admin_users SET role=? WHERE id=?", (new_role, user_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "username": u["username"], "role": new_role})


@app.route("/api/admin/operation_logs", methods=["GET"])
def admin_operation_logs():
    """操作紀錄（老闆專用）：誰做了什麼。?q= 關鍵字、?page=、?limit= 分頁。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    q = (request.args.get("q") or "").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except (ValueError, TypeError):
        page = 1
    try:
        limit = min(200, max(1, int(request.args.get("limit", 50))))
    except (ValueError, TypeError):
        limit = 50
    offset = (page - 1) * limit
    conn = get_db()
    where, params = "", []
    if q:
        like = f"%{q}%"
        where = " WHERE operator LIKE ? OR action LIKE ? OR target LIKE ? OR detail LIKE ?"
        params = [like, like, like, like]
    total = conn.execute(f"SELECT COUNT(*) AS c FROM operation_logs{where}", params).fetchone()["c"]
    rows = conn.execute(
        f"SELECT * FROM operation_logs{where} ORDER BY id DESC LIMIT ? OFFSET ?",
        params + [limit, offset]
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "logs": [dict(r) for r in rows],
                    "total": total, "page": page, "limit": limit,
                    "has_more": offset + len(rows) < total})


# TODO: 此函式會就地修改傳入 list 的元素（每個 dict 寫入 disabled/disabled_reason/disabled_at），
# get_all_members 交進來的是 Shopify 快取裡的 dict，所以快取元素也會被加上這三個 key。
# 目前冪等無害（每次都從 disabled_members 重算後覆寫、不讀舊值、不影響筆數）；
# 日後若要隔離，改成回傳新 dict（[{**m, "disabled": ...} for m in members]），不要用 deepcopy。
def _mark_disabled(members):
    """為 members list 標記停用狀態（依 disabled_members 表）。"""
    try:
        conn = get_db()
        dis = {r["g_code"]: r for r in conn.execute("SELECT g_code, reason, disabled_at FROM disabled_members").fetchall()}
        conn.close()
    except Exception:
        dis = {}
    for m in members:
        d = dis.get(m.get("g_code"))
        m["disabled"] = bool(d)
        m["disabled_reason"] = d["reason"] if d else ""
        m["disabled_at"] = d["disabled_at"] if d else ""
    return members


@app.route("/api/admin/members", methods=["GET"])
def get_all_members():
    try:
        aid = get_current_agent_id()
        # ===== 代理：只看自己本地建的會員 =====
        if aid > 0:
            conn = get_db()
            agent = conn.execute("SELECT prefix, min_rate FROM agents WHERE id=?", (aid,)).fetchone()
            prefix = agent["prefix"] if agent else "X"
            min_rate = float(agent["min_rate"] or 180) if agent else 180.0
            rows = conn.execute(
                "SELECT * FROM members WHERE agent_id=? ORDER BY g_code", (aid,)
            ).fetchall()
            conn.close()
            members = []
            used_numbers = set()
            for r in rows:
                d = dict(r)
                # 會員專屬費率 > 0 → 用該費率；否則 fallback 到代理 min_rate
                member_rate = float(d.get("shipping_rate") or 0)
                effective_rate = member_rate if member_rate > 0 else min_rate
                members.append({
                    "g_code": d.get("g_code", ""),
                    "name": d.get("name", ""),
                    "phone": d.get("phone", ""),
                    "address": d.get("address", ""),
                    "line_id": d.get("line_id", ""),
                    "email": d.get("email", ""),
                    "shipping_rate": effective_rate,
                    "shipping_rate_raw": member_rate,  # 0 表示沿用 min_rate
                    "note": d.get("note", ""),
                    "status": d.get("status", "active"),
                    "source": "local",
                })
                gc = d.get("g_code", "")
                if gc.startswith(prefix):
                    try:
                        used_numbers.add(int(gc[len(prefix):]))
                    except (ValueError, TypeError):
                        pass
            max_number = max(used_numbers) if used_numbers else 0
            next_number = 1
            while next_number in used_numbers:
                next_number += 1
            next_g_code = f"{prefix}{next_number:04d}"
            return jsonify({
                "success": True,
                "members": _mark_disabled(members),
                "total": len(members),
                "max_number": max_number,
                "next_g_code": next_g_code,
                "default_shipping_rate": min_rate,
                "twd_to_jpy_rate": TWD_TO_JPY_RATE,
                "min_rate": min_rate,
                "prefix": prefix,
                "source": "agent_local",
            })

        # ===== 主管理員：Shopify + 全部本地會員（含所有代理底下的）=====
        force = request.args.get("refresh") == "1"
        members = get_all_goyoutati_customers(force_refresh=force)
        # 附加所有本地會員（含代理 agent_id>0 的、以及離職移交 agent_id=0 的）
        try:
            conn0 = get_db()
            local_rows = conn0.execute(
                "SELECT m.*, a.name AS agent_name, a.prefix AS agent_prefix, a.min_rate AS agent_min_rate "
                "FROM members m LEFT JOIN agents a ON a.id = m.agent_id "
                "ORDER BY m.g_code"
            ).fetchall()
            conn0.close()
            for r in local_rows:
                d = dict(r)
                m_rate = float(d.get("shipping_rate") or 0)
                aid_of_member = int(d.get("agent_id") or 0)
                if aid_of_member > 0:
                    # 代理底下的會員：member rate > 0 用會員專屬、否則用代理 min_rate
                    agent_min = float(d.get("agent_min_rate") or DEFAULT_SHIPPING_RATE)
                    effective_rate = m_rate if m_rate > 0 else agent_min
                    source = "agent"
                    agent_name = d.get("agent_name", "") or ""
                else:
                    # 離職移交 / 主管理員直接管
                    effective_rate = m_rate if m_rate > 0 else DEFAULT_SHIPPING_RATE
                    source = "transferred"
                    agent_name = ""
                members.append({
                    "g_code": d.get("g_code", ""),
                    "name": d.get("name", ""),
                    "phone": d.get("phone", ""),
                    "address": d.get("address", ""),
                    "line_id": d.get("line_id", ""),
                    "email": d.get("email", ""),
                    "shipping_rate": effective_rate,
                    "note": d.get("note", ""),
                    "status": d.get("status", "active"),
                    "source": source,
                    "agent_name": agent_name,
                    "agent_id": aid_of_member,
                    "customer_id": "",  # 本地會員無 Shopify ID
                })
        except Exception as e:
            print(f"[admin members] 抓本地會員失敗: {e}", flush=True)
        members.sort(key=lambda x: x["g_code"])
        used_numbers = set()
        for m in members:
            if m["g_code"].startswith("G"):
                try:
                    used_numbers.add(int(m["g_code"][1:]))
                except:
                    pass
        max_number = max(used_numbers) if used_numbers else 0
        next_number = 1
        while next_number in used_numbers:
            next_number += 1
        next_g_code = f"G{next_number:04d}"
        return jsonify({
            "success": True,
            "members": _mark_disabled(members),
            "total": len(members),
            "max_number": max_number,
            "next_g_code": next_g_code,
            "default_shipping_rate": DEFAULT_SHIPPING_RATE,  # 台幣
            "twd_to_jpy_rate": TWD_TO_JPY_RATE
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/search_members", methods=["GET"])
def admin_search_members():
    """
    跨會員搜尋（給新增到貨等場景使用，倉庫人員打字即時搜尋）
    - 主管理員：搜本地 members 表 + Shopify 兩邊（含所有代理客戶）
    - 代理：只搜自己 members 表的客戶
    - 比對欄位：g_code、name、phone（去空白後）
    """
    q = (request.args.get("q") or "").strip()
    try:
        limit = int(request.args.get("limit", 15))
    except (ValueError, TypeError):
        limit = 15
    limit = max(1, min(limit, 50))

    if not q or len(q) < 1:
        return jsonify({"success": True, "results": []})

    q_upper = q.upper()
    q_phone = normalize_phone(q)
    pattern = f"%{q_upper}%"
    pattern_phone = f"%{q_phone}%"
    aid = get_current_agent_id()
    results = []

    conn = get_db()
    # 本地 members（代理建的客戶）
    if aid > 0:
        local = conn.execute("""
            SELECT g_code, name, phone, address, agent_id, status
            FROM members
            WHERE agent_id=? AND status='active'
              AND (UPPER(g_code) LIKE ? OR UPPER(name) LIKE ? OR phone LIKE ?)
            ORDER BY g_code LIMIT ?
        """, (aid, pattern, pattern, pattern_phone, limit)).fetchall()
    else:
        local = conn.execute("""
            SELECT m.g_code, m.name, m.phone, m.address, m.agent_id, m.status,
                   a.name as agent_name, a.prefix as agent_prefix
            FROM members m
            LEFT JOIN agents a ON a.id = m.agent_id
            WHERE m.status='active'
              AND (UPPER(m.g_code) LIKE ? OR UPPER(m.name) LIKE ? OR m.phone LIKE ?)
            ORDER BY m.g_code LIMIT ?
        """, (pattern, pattern, pattern_phone, limit)).fetchall()
    conn.close()

    for r in local:
        d = dict(r)
        results.append({
            "g_code": d.get("g_code"),
            "name": d.get("name") or "",
            "phone": d.get("phone") or "",
            "address": d.get("address") or "",
            "source": "agent",
            "agent_id": d.get("agent_id") or 0,
            "agent_name": d.get("agent_name") or "",
        })

    # Shopify 客戶（僅主管理員視角）
    if aid == 0:
        try:
            customers = get_all_goyoutati_customers()
            for c in customers:
                if len(results) >= limit:
                    break
                gc = (c.get("g_code") or "").upper()
                nm = (c.get("name") or "").upper()
                ph = c.get("phone") or ""
                if q_upper in gc or q_upper in nm or (q_phone and q_phone in ph):
                    results.append({
                        "g_code": c.get("g_code"),
                        "name": c.get("name") or "",
                        "phone": ph,
                        "address": c.get("address") or "",
                        "source": "shopify",
                        "agent_id": 0,
                        "agent_name": "",
                    })
        except Exception as e:
            print(f"[search_members] Shopify 搜尋失敗：{e}", flush=True)

    return jsonify({"success": True, "results": results[:limit]})


# ===== 停用 / 啟用會員（集運系統層級，不動 Shopify）=====

@app.route("/api/admin/members/<g_code>/disable", methods=["POST"])
def admin_disable_member(g_code):
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    g_code = (g_code or "").strip().upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"}), 400
    reason = ((request.json or {}).get("reason") or "").strip()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    conn.execute(
        "INSERT INTO disabled_members (g_code, reason, disabled_at) VALUES (?, ?, ?) "
        "ON CONFLICT(g_code) DO UPDATE SET reason=excluded.reason, disabled_at=excluded.disabled_at",
        (g_code, reason, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/members/<g_code>/disable", methods=["DELETE"])
def admin_enable_member(g_code):
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    g_code = (g_code or "").strip().upper()
    conn = get_db()
    conn.execute("DELETE FROM disabled_members WHERE g_code=?", (g_code,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ── 會員主檔同步（M1）：手動觸發 + 唯讀狀態。員工也能按（is_super_admin）──
@app.route("/api/admin/members/sync", methods=["POST"])
def admin_members_sync():
    """手動觸發 Shopify → members 同步。手貼完 metafield 可立刻按，不用等 30 分鐘。
    不受定時間隔限制，但同一時間只允許一個（租約搶不到就回「同步進行中」）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    report = run_member_sync(trigger=f"manual:{current_operator()}")
    if report.get("skipped"):
        return jsonify({"success": False, "error": report.get("error") or "同步進行中", "running": True}), 409
    try:
        log_op("member_sync", "members",
               f"inserted={report.get('inserted', 0)} updated={report.get('updated', 0)} "
               f"missing={report.get('shopify_missing', 0)} conflicts={len(report.get('conflicts') or [])} "
               f"duplicates={len(report.get('duplicates') or [])}"
               + (f" error={report.get('error')}" if not report.get("success") else ""))
    except Exception:
        pass
    return jsonify(report)


@app.route("/api/admin/members/sync_status", methods=["GET"])
def admin_members_sync_status():
    """最後一次同步時間、各項統計、source 分布。唯讀。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    try:
        last_report = json.loads(_get_setting(MEMBER_SYNC_REPORT_KEY, "") or "{}")
    except (ValueError, TypeError):
        last_report = {}
    conn = get_db()
    try:
        n_shopify = conn.execute("SELECT COUNT(*) AS c FROM members_shopify").fetchone()["c"]
        missing_total = conn.execute(
            "SELECT COUNT(*) AS c FROM members_shopify WHERE status='shopify_missing'"
        ).fetchone()["c"]
        # members 表只讀：agent_id>0 = 代理會員、=0 = 本地（離職移交/主帳號直接管）
        n_agent = conn.execute("SELECT COUNT(*) AS c FROM members WHERE agent_id>0").fetchone()["c"]
        n_local = conn.execute("SELECT COUNT(*) AS c FROM members WHERE agent_id=0").fetchone()["c"]
    finally:
        conn.close()
    dist = {"shopify": n_shopify, "agent": n_agent, "local": n_local}
    fb = _fallback_stats_snapshot()
    return jsonify({
        "success": True,
        "last_sync_at": _get_setting(MEMBER_SYNC_LAST_KEY, ""),
        "running": _member_sync_is_running(),
        "last_report": last_report,
        "source_dist": dist,
        "local_total": n_shopify,   # members_shopify 筆數（同步的主檔）
        "members_total": n_agent + n_local,
        "shopify_missing_total": missing_total,
        "interval_min": _get_sync_interval_min(),
        # 登入 fallback 計數（自程序啟動起算、不持久化、多 worker 各自計）：
        # attempts 高但 hits≈0 → 多半是打錯客編，不是手貼 metafield 的空窗期
        "fallback_attempts": fb["attempts"],
        "fallback_hits": fb["hits"],
    })


@app.route("/api/admin/settings/sync_interval", methods=["GET"])
def admin_get_sync_interval():
    """同步間隔（老闆＋員工都能看）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    return jsonify({"success": True, "interval_min": _get_sync_interval_min(), "default": MEMBER_SYNC_INTERVAL_DEFAULT_MIN,
                    "min": MEMBER_SYNC_INTERVAL_MIN_MIN, "max": MEMBER_SYNC_INTERVAL_MAX_MIN})


@app.route("/api/admin/settings/sync_interval", methods=["PUT"])
def admin_set_sync_interval():
    """修改同步間隔（只有老闆）。M2 起這個值 = 本地會員資料（含計費用 shipping_rate）的新鮮度上限。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以修改此設定"}), 403
    raw = (request.json or {}).get("interval_min")
    msg = f"間隔必須是 {MEMBER_SYNC_INTERVAL_MIN_MIN}~{MEMBER_SYNC_INTERVAL_MAX_MIN} 的整數（分鐘）"
    try:
        v = int(str(raw).strip())
    except (ValueError, TypeError, AttributeError):
        return jsonify({"success": False, "error": msg}), 400
    if not (MEMBER_SYNC_INTERVAL_MIN_MIN <= v <= MEMBER_SYNC_INTERVAL_MAX_MIN):
        return jsonify({"success": False, "error": msg}), 400
    old = _get_sync_interval_min()
    _set_setting(MEMBER_SYNC_INTERVAL_KEY, str(v))
    log_op("修改會員同步間隔", "member_sync_interval", f"{old} → {v} 分鐘")
    return jsonify({"success": True, "interval_min": v, "old": old})


# ── SQLite 自動備份（Google Drive）：手動觸發（老闆）、狀態（老闆+員工）、時間設定（老闆）──
@app.route("/api/admin/maintenance/backup_now", methods=["POST"])
def admin_backup_now():
    """立即備份一次。同一時間只允許一個（租約搶不到回 409）。未設定憑證回 400。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以觸發備份"}), 403
    report = run_backup_job(trigger=f"manual:{current_operator()}")
    if report.get("disabled"):
        return jsonify(report), 400
    if report.get("skipped"):
        return jsonify({"success": False, "error": report.get("error") or "備份進行中", "running": True}), 409
    try:
        log_op("手動備份", "gdrive",
               (f"ok {report.get('filename')} {report.get('size')}B rows={report.get('rows')} "
                f"drive={report.get('drive_count')}") if report.get("success")
               else f"失敗 {report.get('error')}")
    except Exception:
        pass
    return jsonify(report)


@app.route("/api/admin/maintenance/backup_status", methods=["GET"])
def admin_backup_status():
    """最後成功/失敗時間與原因、Drive 份數、最新檔案大小、是否逾時（>48h 沒成功）。唯讀。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    try:
        return jsonify(backup_status_snapshot())
    except Exception as e:
        return jsonify({"success": False, "error": f"讀取備份狀態失敗：{e}"}), 500


@app.route("/api/admin/settings/backup_hour", methods=["PUT"])
def admin_set_backup_hour():
    """修改每日備份時間（0~23，容器本地時間；只有老闆）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以修改此設定"}), 403
    raw = (request.json or {}).get("backup_hour")
    try:
        v = int(str(raw).strip())
    except (ValueError, TypeError, AttributeError):
        return jsonify({"success": False, "error": "時間必須是 0~23 的整數"}), 400
    if not (0 <= v <= 23):
        return jsonify({"success": False, "error": "時間必須是 0~23 的整數"}), 400
    old = _get_backup_hour()
    _set_setting(BACKUP_HOUR_KEY, str(v))
    log_op("修改備份時間", "backup_hour", f"{old} → {v} 時")
    return jsonify({"success": True, "backup_hour": v, "old": old})


@app.route("/api/admin/members", methods=["POST"])
def admin_create_member():
    """代理新增會員（自動補前綴與 agent_id）。主管理員建議直接在 Shopify 操作。"""
    aid = get_current_agent_id()
    if aid <= 0:
        return jsonify({"success": False, "error": "主管理員的會員請在 Shopify 後台建立"}), 400
    data = request.json or {}
    name = (data.get("name") or "").strip()
    phone = normalize_phone((data.get("phone") or "").strip())
    address = (data.get("address") or "").strip()
    if not name:
        return jsonify({"success": False, "error": "會員姓名必填"})
    if not phone:
        return jsonify({"success": False, "error": "電話必填（會員用此登入）"})

    conn = get_db()
    ag = conn.execute("SELECT prefix FROM agents WHERE id=?", (aid,)).fetchone()
    if not ag:
        conn.close()
        return jsonify({"success": False, "error": "代理資料異常"}), 500
    prefix = ag["prefix"]

    # 自動產生下一個 g_code（前綴+四位流水）
    g_code_in = (data.get("g_code") or "").strip().upper()
    if g_code_in:
        # 手動指定的：必須以該代理前綴開頭
        if not g_code_in.startswith(prefix):
            conn.close()
            return jsonify({"success": False, "error": f"會員編號必須以「{prefix}」開頭"})
        # 不可重複
        exists = conn.execute("SELECT 1 FROM members WHERE g_code=?", (g_code_in,)).fetchone()
        if exists:
            conn.close()
            return jsonify({"success": False, "error": f"編號「{g_code_in}」已使用"})
        g_code = g_code_in
    else:
        used = set()
        for r in conn.execute("SELECT g_code FROM members WHERE agent_id=?", (aid,)).fetchall():
            gc = r["g_code"] or ""
            if gc.startswith(prefix):
                try:
                    used.add(int(gc[len(prefix):]))
                except (ValueError, TypeError):
                    pass
        n = 1
        while n in used:
            n += 1
        g_code = f"{prefix}{n:04d}"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 處理會員專屬費率（>= 代理 min_rate；留空/0 = 沿用 min_rate）
    raw_rate = data.get("shipping_rate")
    rate_val = 0.0
    if raw_rate not in (None, "", 0, "0"):
        try:
            rate_val = float(raw_rate)
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": "運費必須為數字"})
        # 代理可自由設定費率，無下限
    try:
        conn.execute(
            """INSERT INTO members (g_code, agent_id, name, phone, address, line_id, email, note, status, created_at, shipping_rate)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active', ?, ?)""",
            (g_code, aid, name, phone, address,
             (data.get("line_id") or "").strip(), (data.get("email") or "").strip(),
             (data.get("note") or "").strip(), now, rate_val)
        )
        conn.commit()
    except sqlite3.IntegrityError as e:
        conn.close()
        return jsonify({"success": False, "error": f"資料庫錯誤：{e}"})
    conn.close()
    return jsonify({"success": True, "g_code": g_code, "message": f"已建立「{g_code} {name}」"})


@app.route("/api/admin/members/<g_code>", methods=["PUT"])
def admin_update_member(g_code):
    """代理更新自己會員的資料（姓名、電話、地址等）"""
    g_code = g_code.upper()
    aid = get_current_agent_id()
    conn = get_db()
    row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
    if not row:
        conn.close()
        return jsonify({"success": False, "error": "找不到會員"})
    if aid > 0 and int(row["agent_id"] or 0) != aid:
        conn.close()
        return jsonify({"success": False, "error": "權限不足"}), 403

    data = request.json or {}
    fields, values = [], []
    for col in ["name", "phone", "address", "line_id", "email", "note", "status"]:
        if col in data:
            v = (data[col] or "").strip()
            if col == "phone":
                v = normalize_phone(v)
            if col == "status" and v not in ("active", "disabled"):
                continue
            fields.append(f"{col}=?"); values.append(v)
    # 會員專屬費率
    if "shipping_rate" in data:
        raw = data["shipping_rate"]
        if raw in (None, "", 0, "0"):
            rate_val = 0.0
        else:
            try:
                rate_val = float(raw)
            except (ValueError, TypeError):
                conn.close()
                return jsonify({"success": False, "error": "運費必須為數字"})
            # 代理可自由設定費率，無下限
        fields.append("shipping_rate=?"); values.append(rate_val)
    if not fields:
        conn.close()
        return jsonify({"success": False, "error": "沒有可更新欄位"})
    values.append(g_code)
    conn.execute(f"UPDATE members SET {', '.join(fields)} WHERE g_code=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/members/<g_code>", methods=["DELETE"])
def admin_delete_member(g_code):
    """代理刪除自己會員（若會員底下有任何包裹/預報/出貨紀錄，擋下，建議停用）"""
    g_code = g_code.upper()
    aid = get_current_agent_id()
    conn = get_db()
    row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
    if not row:
        conn.close()
        return jsonify({"success": False, "error": "找不到會員"})
    if aid > 0 and int(row["agent_id"] or 0) != aid:
        conn.close()
        return jsonify({"success": False, "error": "權限不足"}), 403
    # 安全檢查：是否有關聯資料
    p = conn.execute("SELECT COUNT(*) as c FROM packages WHERE g_code=?", (g_code,)).fetchone()["c"]
    f = conn.execute("SELECT COUNT(*) as c FROM forecasts WHERE g_code=?", (g_code,)).fetchone()["c"]
    s = conn.execute("SELECT COUNT(*) as c FROM shipment_requests WHERE g_code=?", (g_code,)).fetchone()["c"]
    if p + f + s > 0:
        conn.close()
        return jsonify({
            "success": False,
            "error": f"此會員已有 {p} 個包裹/{f} 個預報/{s} 個出貨紀錄，無法刪除。請改為「停用」狀態。"
        })
    conn.execute("DELETE FROM members WHERE g_code=?", (g_code,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "會員已刪除"})


@app.route("/api/admin/shipping_rate", methods=["POST"])
def set_shipping_rate():
    # ⚠️ 原本完全沒有身分檢查，而這支是「寫入」：直接打 Shopify metafieldsSet
    # 改任一客戶的每公斤運費。守門層級比照同樣會動 Shopify 的 /api/admin/members/sync：
    # is_super_admin() = admin_users 的人（老闆＋員工）。代理是 user_type='agent'，
    # 擋掉是對的 —— 這支改的是主帳號 Shopify 客戶的費率，代理客戶走 members 那條路。
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    data = request.json
    customer_gid = data.get("customer_gid", "")
    shipping_rate = data.get("shipping_rate", "")  # 台幣
    if not customer_gid:
        return jsonify({"success": False, "error": "缺少客戶 ID"})
    if shipping_rate == "" or shipping_rate is None:
        return jsonify({"success": False, "error": "請輸入運費"})
    try:
        rate_val = int(shipping_rate)
        if rate_val < 0:
            return jsonify({"success": False, "error": "運費不能為負數"})
    except ValueError:
        return jsonify({"success": False, "error": "運費必須為整數"})

    mutation = """
    mutation metafieldsSet($metafields: [MetafieldsSetInput!]!) {
        metafieldsSet(metafields: $metafields) {
            metafields { key value }
            userErrors { field message }
        }
    }
    """
    variables = {
        "metafields": [{
            "ownerId": customer_gid,
            "namespace": "custom",
            "key": "shipping_rate",
            "type": "single_line_text_field",
            "value": str(rate_val)  # 儲存台幣值
        }]
    }
    try:
        result = shopify_graphql(mutation, variables)
        if "data" in result:
            mutation_result = result["data"].get("metafieldsSet", {})
            user_errors = mutation_result.get("userErrors", [])
            if user_errors:
                return jsonify({"success": False, "error": "; ".join([e["message"] for e in user_errors])})
            if mutation_result.get("metafields"):
                # M2：讀取走本地 members_shopify，Shopify 寫成功後順手更新本地，不用等下一輪同步
                try:
                    cid = customer_gid.split("/")[-1] if "/" in customer_gid else customer_gid
                    _c = get_db()
                    _c.execute("UPDATE members_shopify SET shipping_rate=?, synced_at=? WHERE shopify_customer_id=?",
                               (float(rate_val), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cid))
                    _c.commit(); _c.close()
                except Exception as _e:
                    print(f"[shipping_rate] 本地 members_shopify 更新失敗（下一輪同步會補上）: {_e}", flush=True)
                return jsonify({
                    "success": True,
                    "shipping_rate_twd": rate_val,
                    "shipping_rate_jpy": twd_to_jpy(rate_val)
                })
        if "errors" in result:
            return jsonify({"success": False, "error": str(result["errors"])})
        return jsonify({"success": False, "error": "設定失敗，請重試"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# ============ 管理員：到貨包裹管理 ============

@app.route("/api/admin/packages", methods=["GET"])
def admin_list_packages():
    g_code = request.args.get("g_code", "")
    aid = get_current_agent_id()
    conn = get_db()
    if g_code:
        # 單一客戶的包裹（供會員明細等用）：維持原樣，不分頁
        if aid > 0:
            rows = conn.execute(
                "SELECT * FROM packages WHERE g_code=? AND agent_id=? ORDER BY id DESC",
                (g_code.upper(), aid)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM packages WHERE g_code=? ORDER BY id DESC", (g_code.upper(),)
            ).fetchall()
        conn.close()
        return jsonify({"success": True, "packages": [dict(r) for r in rows]})

    # 到貨管理列表：後端分頁 + 狀態 + 搜尋
    status = (request.args.get("status") or "").strip()
    q = (request.args.get("q") or "").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except (ValueError, TypeError):
        page = 1
    try:
        limit = min(200, max(1, int(request.args.get("limit", 50))))
    except (ValueError, TypeError):
        limit = 50
    offset = (page - 1) * limit

    where, params = [], []
    if aid > 0:
        where.append("agent_id=?"); params.append(aid)
    if q:
        like = f"%{q}%"
        where.append("(g_code LIKE ? OR logis_num LIKE ? OR product_name LIKE ?)")
        params += [like, like, like]   # 搜尋時忽略狀態、跨全部
    else:
        if status == "未出貨":
            where.append("status!='已出貨'")
        elif status and status not in ("全部", "最近50"):
            where.append("status=?"); params.append(status)
        # 全部 / 最近50 → 不加狀態條件（最近50 由分頁自然呈現）

    wsql = (" WHERE " + " AND ".join(where)) if where else ""
    total = conn.execute(f"SELECT COUNT(*) AS c FROM packages{wsql}", params).fetchone()["c"]
    rows = conn.execute(
        f"SELECT * FROM packages{wsql} ORDER BY id DESC LIMIT ? OFFSET ?",
        params + [limit, offset]
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "packages": [dict(r) for r in rows],
                    "total": total, "page": page, "limit": limit,
                    "has_more": offset + len(rows) < total})


# ===== 無主包裹認領牆 =====

def _uc_days(date_str):
    """到倉天數（登記日算起）。日期壞掉或沒填就回 None，前端顯示 '-'。"""
    try:
        d = datetime.strptime((date_str or "")[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None
    return max(0, (datetime.now().date() - d).days)


def _uc_mask_logis(logis_num, keep=4):
    """物流單號只給末四碼，避免整組單號被任意會員拿去查詢。"""
    s = (logis_num or "").strip()
    if not s:
        return ""
    return s[-keep:] if len(s) > keep else s


def _uc_valid_customer(g_code):
    """判斷 g_code 是不是有效客戶（不驗密碼），來源與 verify_customer() 完全一致：
    停用名單 → 擋；先查本地 members（代理建的會員），查不到再回退 Shopify 客戶清單。
    回傳 (normalized_g_code, source, name, error)；error 非 None 代表不放行。"""
    g_code = (g_code or "").strip().upper()
    if not g_code:
        return "", "", "", "請先登入"
    # 沒有英文前綴 → 預設加 G（與 verify_customer 相同）
    if not g_code[:1].isalpha():
        g_code = "G" + g_code

    # 0) 停用名單（集運系統層級）
    try:
        conn = get_db()
        if conn.execute("SELECT 1 FROM disabled_members WHERE g_code=?", (g_code,)).fetchone():
            conn.close()
            return g_code, "", "", "您的帳號已停用，請聯繫客服"
        # 1) 本地 members 表
        row = conn.execute("SELECT name, status FROM members WHERE g_code=?", (g_code,)).fetchone()
        conn.close()
        if row:
            if (row["status"] or "") == "disabled":
                return g_code, "", "", "此會員帳號已停用，請聯絡您的代理"
            return g_code, "agent", (row["name"] or ""), None
    except Exception as e:
        print(f"[_uc_valid_customer] 本地查詢失敗：{e}", flush=True)

    # 2) 回退 Shopify 客戶（只在本地查不到時才走，走快取不打 API）
    try:
        for c in get_all_goyoutati_customers():
            if c.get("g_code") == g_code:
                return g_code, "shopify", (c.get("name") or ""), None
    except Exception as e:
        print(f"[_uc_valid_customer] Shopify 查詢失敗：{e}", flush=True)

    return g_code, "", "", "查無此會員編號"


@app.route("/api/unclaimed", methods=["GET"])
def member_list_unclaimed():
    """會員端認領牆：單號末四碼 + 到倉天數（不回傳收件人姓名，避免外洩其他客戶資料）。
    認領牆本來就是登入後的功能，這裡綁 session。"""
    # ⚠️ 順序不可對調：_require_customer 必須是第一件事。
    #    若先跑 _uc_valid_customer，未登入者會先拿到「客編不存在／已停用」這種
    #    區別性錯誤，等於留下一個「這個客編存不存在」的線上查詢器。
    ok, resp = _require_customer(request.args.get("g_code"))
    if not ok:
        return resp
    g_code, _source, _name, err = _uc_valid_customer(request.args.get("g_code"))
    if err:
        return jsonify({"success": False, "error": err}), 403

    conn = get_db()
    rows = conn.execute("SELECT * FROM unclaimed_packages ORDER BY id DESC").fetchall()
    counts = {r["unclaimed_id"]: r["c"] for r in conn.execute(
        "SELECT unclaimed_id, COUNT(*) AS c FROM unclaimed_claims GROUP BY unclaimed_id").fetchall()}
    mine = {r["unclaimed_id"] for r in conn.execute(
        "SELECT unclaimed_id FROM unclaimed_claims WHERE g_code=?", (g_code,)).fetchall()}
    conn.close()

    items = []
    for row in rows:
        r = dict(row)
        items.append({
            "id":              r["id"],
            "logis_tail":      _uc_mask_logis(r.get("logis_num")),
            "product_name":    r.get("product_name") or "",
            "weight":          r.get("weight") or "",
            "note":            r.get("note") or "",
            "registered_date": r.get("registered_date") or "",
            "days":            _uc_days(r.get("registered_date")),
            "claim_count":     counts.get(r["id"], 0),
            "claimed_by_me":   r["id"] in mine,
        })
    return jsonify({"success": True, "items": items, "total": len(items)})


@app.route("/api/unclaimed/<int:uid>/claim_request", methods=["POST"])
def member_request_unclaimed(uid):
    """會員申請認領：只留申請紀錄，包裹仍留在牆上，由管理員確認後才轉入。"""
    data = request.json or {}
    note = (data.get("note") or "").strip()[:200]
    g_code, _source, member_name, err = _uc_valid_customer(data.get("g_code"))
    if err:
        return jsonify({"success": False, "error": err}), 403
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp

    conn = get_db()
    if not conn.execute("SELECT 1 FROM unclaimed_packages WHERE id=?", (uid,)).fetchone():
        conn.close()
        return jsonify({"success": False, "error": "此包裹已被認領或已移除"}), 404

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        conn.execute(
            """INSERT INTO unclaimed_claims (unclaimed_id, g_code, member_name, note, created_at)
               VALUES (?, ?, ?, ?, ?)""",
            (uid, g_code, member_name, note, now)
        )
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"success": True, "duplicated": True,
                        "message": "你已經申請過這件了，請等候倉庫確認"})
    conn.close()
    log_op("申請認領無主件", g_code, f"無主件 #{uid}" + (f"／{note}" if note else ""))
    return jsonify({"success": True, "message": "已送出認領申請，倉庫確認後會轉入你的包裹"})


@app.route("/api/unclaimed/<int:uid>/claim_request", methods=["DELETE"])
def member_cancel_unclaimed(uid):
    """會員自行撤回認領申請（按錯了）。"""
    raw = (request.json or {}).get("g_code") or request.args.get("g_code")
    g_code, _source, _name, err = _uc_valid_customer(raw)
    if err:
        return jsonify({"success": False, "error": err}), 403
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    conn.execute("DELETE FROM unclaimed_claims WHERE unclaimed_id=? AND g_code=?", (uid, g_code))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/unclaimed", methods=["GET"])
def admin_list_unclaimed():
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    rows = conn.execute("SELECT * FROM unclaimed_packages ORDER BY id DESC").fetchall()
    claim_rows = conn.execute(
        "SELECT unclaimed_id, g_code, member_name, note, created_at "
        "FROM unclaimed_claims ORDER BY id ASC").fetchall()
    conn.close()
    claims = {}
    for c in claim_rows:
        claims.setdefault(c["unclaimed_id"], []).append(dict(c))
    items = []
    for r in rows:
        d = dict(r)
        d["claims"] = claims.get(d["id"], [])   # 會員申請認領的名單（可能多人搶同一件）
        items.append(d)
    return jsonify({"success": True, "items": items})


@app.route("/api/admin/unclaimed", methods=["POST"])
def admin_add_unclaimed():
    """倉庫登記一件無主包裹（記到倉日期）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today = datetime.now().strftime("%Y-%m-%d")
    reg_date = (data.get("registered_date") or "").strip() or today
    conn = get_db()
    cur = conn.execute(
        """INSERT INTO unclaimed_packages (recipient_name, logis_num, product_name, weight, note, registered_date, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        ((data.get("recipient_name") or "").strip(), (data.get("logis_num") or "").strip(),
         (data.get("product_name") or "").strip(), (data.get("weight") or "").strip(),
         (data.get("note") or "").strip(), reg_date, now)
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": new_id})


@app.route("/api/admin/unclaimed/<int:uid>", methods=["DELETE"])
def admin_delete_unclaimed(uid):
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    conn.execute("DELETE FROM unclaimed_packages WHERE id=?", (uid,))
    conn.execute("DELETE FROM unclaimed_claims WHERE unclaimed_id=?", (uid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/unclaimed/<int:uid>/claim", methods=["POST"])
def admin_claim_unclaimed(uid):
    """認領：填客編 → 轉入 packages（保留原始到倉日 in_date）→ 從認領牆移除。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    g_code = ((request.json or {}).get("g_code") or "").strip().upper()
    if not g_code:
        return jsonify({"success": False, "error": "請輸入客戶編號"}), 400
    if not g_code[:1].isalpha():
        g_code = "G" + g_code

    conn = get_db()
    u = conn.execute("SELECT * FROM unclaimed_packages WHERE id=?", (uid,)).fetchone()
    if not u:
        conn.close()
        return jsonify({"success": False, "error": "找不到此無主包裹"}), 404
    u = dict(u)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    in_date = u.get("registered_date") or datetime.now().strftime("%Y-%m-%d")  # 保留原始到倉日
    pkg_agent_id = get_agent_id_for_g_code(g_code)
    conn.execute(
        """INSERT INTO packages (g_code, logis_num, product_name, weight, status, note, in_date, created_at, agent_id, pkg_type)
           VALUES (?, ?, ?, ?, '已到貨', ?, ?, ?, ?, '包裹')""",
        (g_code, u.get("logis_num", ""), u.get("product_name", ""), u.get("weight", ""),
         u.get("note", ""), in_date, now, pkg_agent_id)
    )
    conn.execute("DELETE FROM unclaimed_packages WHERE id=?", (uid,))
    conn.execute("DELETE FROM unclaimed_claims WHERE unclaimed_id=?", (uid,))
    conn.commit()
    conn.close()
    log_op("認領無主件", g_code, f"{u.get('recipient_name','')} → 到倉日 {in_date}")
    return jsonify({"success": True, "g_code": g_code, "in_date": in_date})


@app.route("/api/admin/packages", methods=["POST"])
def admin_add_package():
    data = request.json
    g_code      = (data.get("g_code") or "").strip().upper()
    logis_num   = (data.get("logis_num") or "").strip()
    product_name= (data.get("product_name") or "").strip()
    weight      = (data.get("weight") or "").strip()
    note        = (data.get("note") or "").strip()
    status      = data.get("status", "已到貨")
    pkg_type    = data.get("pkg_type", "包裹")
    if pkg_type not in ("包裹", "信件"):
        pkg_type = "包裹"

    if not g_code:
        return jsonify({"success": False, "error": "請輸入客戶編號"})
    # 開頭非字母 → 補上對應前綴（代理用自己的前綴、主管理員預設 G）
    if not g_code[:1].isalpha():
        aid = get_current_agent_id()
        if aid > 0:
            conn0 = get_db()
            ag = conn0.execute("SELECT prefix FROM agents WHERE id=?", (aid,)).fetchone()
            conn0.close()
            g_code = (ag["prefix"] if ag else "G") + g_code
        else:
            g_code = "G" + g_code

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today = datetime.now().strftime("%Y-%m-%d")
    pkg_agent_id = get_agent_id_for_g_code(g_code)

    # 代理只能幫自己的客戶建包裹
    aid = get_current_agent_id()
    if aid > 0 and pkg_agent_id != aid:
        return jsonify({"success": False, "error": f"客戶編號「{g_code}」不屬於你的代理帳號"}), 403

    conn = get_db()
    cur = conn.execute(
        """INSERT INTO packages (g_code, logis_num, product_name, weight, status, note, in_date, created_at, agent_id, pkg_type)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (g_code, logis_num, product_name, weight, status, note, today, now, pkg_agent_id, pkg_type)
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    log_op("登記到貨", g_code, f"{product_name or logis_num} {weight}kg")
    return jsonify({"success": True, "id": new_id})


@app.route("/api/admin/packages/<int:pkg_id>", methods=["PUT"])
def admin_update_package(pkg_id):
    ok, row = check_record_ownership("packages", pkg_id)
    if not row:
        return jsonify({"success": False, "error": "找不到包裹"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    fields = []
    values = []
    for key in ["g_code", "logis_num", "product_name", "weight", "status", "note", "in_date", "pkg_type"]:
        if key in data:
            val = data[key]
            if key == "g_code":
                val = val.strip().upper()
                # 改 g_code 時驗證新 g_code 仍屬於同代理
                aid_chk = get_current_agent_id()
                if aid_chk > 0 and get_agent_id_for_g_code(val) != aid_chk:
                    return jsonify({"success": False, "error": f"無法將包裹改至非自己客戶「{val}」"}), 403
            fields.append(f"{key}=?")
            values.append(val)
    if not fields:
        return jsonify({"success": False, "error": "沒有要更新的欄位"})
    values.append(pkg_id)
    conn = get_db()
    conn.execute(f"UPDATE packages SET {', '.join(fields)} WHERE id=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/packages/bulk_ship", methods=["POST"])
def admin_bulk_ship():
    data = request.json
    ids = data.get("ids", [])
    if not ids:
        return jsonify({"success": False, "error": "沒有選取任何包裹"})
    aid = get_current_agent_id()
    conn = get_db()
    if aid > 0:
        # 驗證所有 id 都屬於該代理
        placeholders = ",".join(["?"] * len(ids))
        owned = conn.execute(
            f"SELECT id FROM packages WHERE id IN ({placeholders}) AND agent_id=?",
            ids + [aid]
        ).fetchall()
        owned_ids = [r["id"] for r in owned]
        if len(owned_ids) != len(ids):
            conn.close()
            return jsonify({"success": False, "error": "部分包裹不屬於你的代理帳號"}), 403
        ids = owned_ids
    conn.execute(
        f"UPDATE packages SET status='已出貨' WHERE id IN ({','.join(['?']*len(ids))})",
        ids
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "updated": len(ids)})


@app.route("/api/admin/packages/<int:pkg_id>", methods=["DELETE"])
def admin_delete_package(pkg_id):
    ok, row = check_record_ownership("packages", pkg_id)
    if not row:
        return jsonify({"success": False, "error": "找不到包裹"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    conn.execute("DELETE FROM packages WHERE id=?", (pkg_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


def _last5_num_core(v):
    """比對用：去掉 .0 尾巴與前導零，只留數字本身；非數字原樣。"""
    t = str(v or "").strip()
    mm = re.fullmatch(r"(\d+)\.0+", t)
    if mm:
        t = mm.group(1)
    return (t.lstrip("0") or "0") if t.isdigit() else t


def _last5_log_originals(conn):
    """從 operation_logs 解析「管理員確認付款」當下的原始輸入（含前導零的 ground truth）。
    回傳 ({req_id: 原值}, 有解析到的紀錄數)。同一單多筆取 created_at 最新者。"""
    rows = conn.execute(
        "SELECT target, detail, created_at, id FROM operation_logs "
        "WHERE action='帳單確認付款' ORDER BY created_at ASC, id ASC"
    ).fetchall()
    by_req, parsed = {}, 0
    for lr in rows:
        m_id = re.search(r"\d+", str(lr["target"] or ""))        # 容忍「出貨單#123」以外的格式
        if not m_id:
            continue
        m_val = re.match(r"^\s*後五碼\s*(.*)$", str(lr["detail"] or ""))
        if not m_val:
            continue
        by_req[int(m_id.group(0))] = m_val.group(1).strip()
        parsed += 1
    return by_req, parsed


def _strip_dot_zero(txt):
    """'463.0' → '463'；其餘原樣。只處理型別造成的尾巴，不補零。"""
    mm = re.fullmatch(r"(-?\d+)\.0+", str(txt))
    return mm.group(1) if mm else str(txt)


# ============ 維運：欄位 affinity 誤宣告的損害範圍診斷（唯讀）============
# 線上 schema 有一批欄位被宣告成 REAL，但程式碼的 ALTER 清單裡它們都是 TEXT。
# REAL affinity 會把「看起來像數字」的字串自動轉成數字，前導零被吃掉。
# ship_phone 尤其嚴重：09 開頭的手機號被存成 9 碼，已流入交給清關行的出貨單。

AFFINITY_SUSPECT_COLS = [
    "payment_last5", "payment_at", "tracking_num", "extra_services",
    "ship_recipient", "ship_phone", "ship_address",
]
# 宣告成這些型別 + DEFAULT '' 幾乎一定是誤宣告（空字串預設值配數字型別沒有意義）
_NUMERIC_DECL_PREFIXES = ("REAL", "NUMERIC", "DOUBLE", "FLOAT", "DECIMAL")


# ============ 維運：電話欄位 affinity 全庫檢查（唯讀）============
# 客戶登入密碼就是手機號碼。若存手機的欄位是 REAL affinity，
# '0912345678' 會被存成 912345678.0，登入比對就會出問題。
# 任務 K 的掃描條件是「REAL/NUMERIC 且 DEFAULT ''」，會漏掉「REAL 但無 DEFAULT
# 或 DEFAULT NULL」的欄位，所以這裡放寬條件重掃一次。

_PHONE_NAME_HINTS = ("phone", "tel", "mobile", "電話")
# 放寬後要列出的宣告型別（含「無型別」→ BLOB affinity）
_LOOSE_NUMERIC_PREFIXES = ("REAL", "NUMERIC", "INTEGER", "INT", "DOUBLE", "FLOAT", "DECIMAL")

# 客戶登入比對的實際邏輯（見 /api/verify_customer；此處僅為文字說明，不影響行為）
_LOGIN_COMPARE_NOTE = {
    "endpoint": "/api/verify_customer",
    "local_members": (
        "stored_phone = normalize_phone(m.get('phone') or '')；"
        "if stored_phone != password_clean → 密碼錯誤。"
        "比對欄位＝members.phone"
    ),
    "shopify": (
        "if c['phone'] and c['phone'] == password_clean → 通過。"
        "比對欄位＝Shopify 客戶資料的 phone（不經過本地 DB）"
    ),
    "normalize_phone": (
        "只做 replace(' ','') 與 replace('-','')，再把 +886 / +81 前綴換成 0。"
        "★ 沒有 strip()、沒有 zfill、沒有去 .0、沒有任何型別轉換。"
    ),
    "risk_if_real_affinity": (
        "若 members.phone 是 REAL affinity，讀回來會是 float（例如 912345678.0），"
        "normalize_phone 對 float 呼叫 .replace 會拋 AttributeError，"
        "被 verify_customer 的 except 吞掉並印出「本地查詢失敗」，"
        "接著 fallback 去查 Shopify → 本地會員將完全登不進去。"
    ),
}


@app.route("/api/admin/maintenance/phone_affinity_diag", methods=["GET"])
def admin_phone_affinity_diag():
    """電話欄位 affinity 全庫檢查（唯讀）。

    ⚠️ 只做 SELECT / PRAGMA，不含任何 UPDATE / INSERT / DELETE / ALTER / CREATE。
    ⚠️ 所有值一律 CAST(... AS TEXT) 才輸出。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以執行維運作業"}), 403

    conn = get_db()
    try:
        tables = [r["name"] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        ).fetchall() if not r["name"].startswith("sqlite_")]

        def _cols(t):
            try:
                return [dict(r) for r in conn.execute(f'PRAGMA table_info("{t}")').fetchall()]
            except Exception:
                return []

        def _pk_expr(cinfo):
            """回傳可當識別欄位的運算式與別名；沒有主鍵就用 rowid。"""
            pks = [c["name"] for c in cinfo if c.get("pk")]
            return (f'"{pks[0]}"', pks[0]) if len(pks) == 1 else ("rowid", "rowid")

        # ── 1) 全庫放寬重掃：宣告為數值型別或無型別的欄位，不論 DEFAULT ──
        loose = []
        for t in tables:
            for c in _cols(t):
                ctype = str(c.get("type") or "").strip()
                up = ctype.upper()
                if (not ctype) or up.startswith(_LOOSE_NUMERIC_PREFIXES):
                    loose.append({
                        "table": t, "column": c.get("name"),
                        "declared_type": ctype or "(無型別)",
                        "dflt_value": c.get("dflt_value"),
                        "notnull": c.get("notnull"), "pk": c.get("pk"),
                    })

        # ── 2) 欄位名含 phone / tel / mobile / 電話 者，逐一檢查 ──
        phone_cols = []
        for t in tables:
            cinfo = _cols(t)
            pk_expr, pk_name = _pk_expr(cinfo)
            for c in cinfo:
                cname = str(c.get("name") or "")
                if not any(h in cname.lower() for h in _PHONE_NAME_HINTS[:3]) \
                        and "電話" not in cname:
                    continue
                dist = [{"typeof": r["t"], "count": r["c"]} for r in conn.execute(
                    f'SELECT typeof("{cname}") AS t, COUNT(*) AS c FROM "{t}" '
                    f"GROUP BY t ORDER BY c DESC").fetchall()]
                damaged = conn.execute(
                    f'SELECT COUNT(*) FROM "{t}" WHERE typeof("{cname}") IN (\'real\',\'integer\')'
                ).fetchone()[0]
                samples = [
                    {"pk_column": pk_name, "pk": str(r["pk_val"]), "typeof": r["t"],
                     "raw": r["raw"], "len": r["len"], "stripped": _strip_dot_zero(r["raw"] or "")}
                    for r in conn.execute(
                        f'SELECT {pk_expr} AS pk_val, typeof("{cname}") AS t, '
                        f'       CAST("{cname}" AS TEXT) AS raw, '
                        f'       length(CAST("{cname}" AS TEXT)) AS len '
                        f'FROM "{t}" WHERE COALESCE("{cname}",\'\') <> \'\' '
                        f'ORDER BY (typeof("{cname}") IN (\'real\',\'integer\')) DESC, {pk_expr} DESC '
                        f"LIMIT 20"
                    ).fetchall()
                ]
                phone_cols.append({
                    "table": t, "column": cname,
                    "declared_type": c.get("type"), "dflt_value": c.get("dflt_value"),
                    "typeof_distribution": dist, "damaged": damaged, "samples": samples,
                })

        # ── 3) members 表細看（登入密碼比對用的就是這張表的 phone）──
        members_detail = None
        if "members" in tables:
            cinfo = _cols("members")
            pk_expr, pk_name = _pk_expr(cinfo)
            login_col = "phone" if any(c["name"] == "phone" for c in cinfo) else None
            detail = {
                "table_info": [{"name": c.get("name"), "type": c.get("type"),
                                "dflt_value": c.get("dflt_value"), "notnull": c.get("notnull"),
                                "pk": c.get("pk")} for c in cinfo],
                "login_compare_column": login_col,
            }
            if login_col:
                detail["typeof_distribution"] = [
                    {"typeof": r["t"], "count": r["c"]} for r in conn.execute(
                        f'SELECT typeof("{login_col}") AS t, COUNT(*) AS c FROM members '
                        f"GROUP BY t ORDER BY c DESC").fetchall()
                ]
                len_dist, bad = {}, []
                for r in conn.execute(
                    f'SELECT {pk_expr} AS pk_val, typeof("{login_col}") AS t, '
                    f'       CAST("{login_col}" AS TEXT) AS raw '
                    f'FROM members WHERE COALESCE("{login_col}",\'\') <> \'\' '
                    f"ORDER BY {pk_expr}"
                ).fetchall():
                    raw = r["raw"] or ""
                    k = str(len(raw))
                    len_dist[k] = len_dist.get(k, 0) + 1
                    if len(raw) != 10 and len(bad) < 20:
                        bad.append({"pk_column": pk_name, "pk": str(r["pk_val"]), "typeof": r["t"],
                                    "raw": raw, "len": len(raw),
                                    "stripped": _strip_dot_zero(raw)})
                detail["length_distribution"] = dict(sorted(len_dist.items(), key=lambda kv: int(kv[0])))
                detail["not_10_chars"] = {
                    "total": sum(c for k, c in len_dist.items() if int(k) != 10),
                    "shown": len(bad), "samples": bad,
                }
            members_detail = detail
        conn.close()

        return jsonify({
            "success": True,
            "read_only": True,
            "tables_scanned": len(tables),
            "loose_numeric_columns": {"total": len(loose), "items": loose},
            "phone_like_columns": phone_cols,
            "members": members_detail,
            "note": _LOGIN_COMPARE_NOTE,
        })
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/admin/maintenance/affinity_diag", methods=["GET"])
def admin_affinity_diag():
    """七個誤宣告為 REAL 的欄位，損害範圍診斷。

    ⚠️ 只做 SELECT / PRAGMA，不含任何 UPDATE / INSERT / DELETE / ALTER / CREATE。
    ⚠️ 所有值一律 CAST(... AS TEXT) 才輸出，避免 JSON 序列化再吃掉一次型別。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以執行維運作業"}), 403

    conn = get_db()
    try:
        table = "shipment_requests"
        info = {dict(r)["name"]: dict(r) for r in
                conn.execute(f"PRAGMA table_info({table})").fetchall()}
        total_rows = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]

        columns = {}
        for col in AFFINITY_SUSPECT_COLS:
            if col not in info:
                columns[col] = {"exists": False}
                continue
            meta = info[col]
            dist = [{"typeof": r["t"], "count": r["c"]} for r in conn.execute(
                f'SELECT typeof("{col}") AS t, COUNT(*) AS c FROM {table} GROUP BY t ORDER BY c DESC'
            ).fetchall()]
            damaged = conn.execute(
                f"SELECT COUNT(*) FROM {table} WHERE typeof(\"{col}\") IN ('real','integer')"
            ).fetchone()[0]
            samples = [
                {"id": r["id"], "g_code": r["g_code"], "typeof": r["t"],
                 "raw": r["raw"], "len": r["len"], "stripped": _strip_dot_zero(r["raw"] or "")}
                for r in conn.execute(
                    f'SELECT id, g_code, typeof("{col}") AS t, '
                    f'       CAST("{col}" AS TEXT) AS raw, '
                    f'       length(CAST("{col}" AS TEXT)) AS len '
                    f"FROM {table} WHERE typeof(\"{col}\") IN ('real','integer') "
                    f"ORDER BY id DESC LIMIT 30"
                ).fetchall()
            ]
            columns[col] = {
                "exists": True,
                "declared_type": meta.get("type"),
                "dflt_value": meta.get("dflt_value"),
                "typeof_distribution": dist,
                "damaged": damaged,
                "samples": samples,
            }

        # ── ship_phone：去 .0 後的長度分布（預期大量 9 碼＝被吃掉前導 0）──
        phone_len, phone_short = {}, []
        if "ship_phone" in info:
            for r in conn.execute(
                f'SELECT id, g_code, CAST(ship_phone AS TEXT) AS raw, typeof(ship_phone) AS t '
                f"FROM {table} WHERE COALESCE(ship_phone,'') <> '' ORDER BY id DESC"
            ).fetchall():
                v = _strip_dot_zero(r["raw"] or "")
                k = str(len(v))
                phone_len[k] = phone_len.get(k, 0) + 1
                # 8 碼或更短 = 不只掉一個 0，或本來就不是手機 → 要特別看
                if len(v) <= 8 and len(phone_short) < 50:
                    phone_short.append({"id": r["id"], "g_code": r["g_code"], "typeof": r["t"],
                                        "raw": r["raw"], "stripped": v, "len": len(v)})
        ship_phone_extra = {
            "length_distribution": dict(sorted(phone_len.items(), key=lambda kv: int(kv[0]))),
            "expected_9_digits": phone_len.get("9", 0),      # 09xxxxxxxx 掉前導 0 → 9 碼
            "intact_10_digits": phone_len.get("10", 0),
            "short_le_8": {"total": sum(c for k, c in phone_len.items() if int(k) <= 8),
                           "samples": phone_short},
        }

        # ── tracking_num：去 .0 後長度 > 15 → 進過 REAL 會失去精度，屬不可逆損壞 ──
        long_tracking = []
        long_total = 0
        if "tracking_num" in info:
            for r in conn.execute(
                f'SELECT id, g_code, CAST(tracking_num AS TEXT) AS raw, typeof(tracking_num) AS t '
                f"FROM {table} WHERE COALESCE(tracking_num,'') <> '' ORDER BY id DESC"
            ).fetchall():
                v = _strip_dot_zero(r["raw"] or "")
                if len(v) > 15:
                    long_total += 1
                    if len(long_tracking) < 50:
                        long_tracking.append({"id": r["id"], "g_code": r["g_code"], "typeof": r["t"],
                                              "raw": r["raw"], "stripped": v, "len": len(v)})
        tracking_extra = {
            "over_15_chars": {"total": long_total, "samples": long_tracking},
            "note": "超長數字進 REAL 會失去精度，這種損壞不可逆；typeof 為 real 且長度>15 者要特別確認",
        }

        # ── 全資料庫掃描：宣告為 REAL/NUMERIC 但預設值是空字串的欄位 ──
        # 表名由 sqlite_master 取，欄位細節用 PRAGMA table_info（比 regex 解析 SQL 可靠）
        suspects = []
        for tr in conn.execute(
            "SELECT name, sql FROM sqlite_master WHERE type='table' ORDER BY name"
        ).fetchall():
            tname = tr["name"]
            if tname.startswith("sqlite_"):
                continue
            try:
                cinfo = conn.execute(f'PRAGMA table_info("{tname}")').fetchall()
            except Exception:
                continue
            for c in cinfo:
                cd = dict(c)
                ctype = str(cd.get("type") or "").strip().upper()
                dflt = str(cd.get("dflt_value") if cd.get("dflt_value") is not None else "")
                if ctype.startswith(_NUMERIC_DECL_PREFIXES) and dflt in ("''", '""'):
                    suspects.append({"table": tname, "column": cd.get("name"),
                                     "declared_type": cd.get("type"), "dflt_value": cd.get("dflt_value")})
        conn.close()

        return jsonify({
            "success": True,
            "read_only": True,
            "table": table,
            "row_count": total_rows,
            "columns": columns,
            "ship_phone_extra": ship_phone_extra,
            "tracking_num_extra": tracking_extra,
            "db_wide_suspects": {"total": len(suspects), "items": suspects},
        })
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/admin/maintenance/last5_diag", methods=["GET"])
def admin_last5_diag():
    """payment_last5 儲存型別診斷（唯讀）。

    2026-09 的 REAL affinity 事件就是靠這支查出來的（原專案事件紀錄）：
    寫入路徑都寫 zfill 過的字串，照理不該產生 REAL，實際卻是欄位型別被宣告成 REAL。
    資料已修復，這支留著給下次懷疑資料有問題時用。

    ⚠️ 本端點只做 SELECT / PRAGMA，不含任何 UPDATE / INSERT / DELETE / ALTER。
    ⚠️ 所有值一律 CAST(... AS TEXT) 才輸出：JSON 序列化會把 REAL 變成數字、
       型別資訊就消失了（這正是我們現在踩到的坑，診斷工具不能重蹈覆轍）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以執行維運作業"}), 403

    conn = get_db()
    try:
        # 1) schema：兩張表的 payment_last5 欄位定義
        schema = {}
        for t in ("shipment_requests", "agent_payouts"):
            col = None
            for r in conn.execute(f"PRAGMA table_info({t})").fetchall():
                d = dict(r)
                if d.get("name") == "payment_last5":
                    col = {"name": d.get("name"), "type": d.get("type"),
                           "notnull": d.get("notnull"), "dflt_value": d.get("dflt_value")}
                    break
            schema[t] = col

        # 2) 儲存類別分布（SQLite 的實際 storage class，不是宣告型別）
        storage = {}
        for t in ("shipment_requests", "agent_payouts"):
            storage[t] = [
                {"typeof": r["t"], "count": r["c"]}
                for r in conn.execute(
                    f"SELECT typeof(payment_last5) AS t, COUNT(*) AS c "
                    f"FROM {t} GROUP BY t ORDER BY c DESC"
                ).fetchall()
            ]

        # 3) 時間軸：按月看儲存類別 → 一直都髒？還是某個時間點開始髒？
        timeline = [
            {"ym": r["ym"], "typeof": r["t"], "count": r["c"]}
            for r in conn.execute(
                "SELECT substr(payment_at,1,7) AS ym, typeof(payment_last5) AS t, COUNT(*) AS c "
                "FROM shipment_requests "
                "WHERE COALESCE(payment_at,'') <> '' "
                "GROUP BY ym, t ORDER BY ym DESC LIMIT 60"
            ).fetchall()
        ]

        # 4) 模糊樣本：純數字、去掉小數後不足 5 位 → 清理時唯一靠「末五碼必為 5 位」推斷補 0 的部分
        ambiguous = [
            {"id": r["id"], "g_code": r["g_code"],
             "raw_text": r["raw_text"], "typeof": r["t"],
             "stripped": r["stripped"], "payment_at": r["payment_at"]}
            for r in conn.execute(
                "SELECT id, g_code, "
                "       CAST(payment_last5 AS TEXT) AS raw_text, "
                "       typeof(payment_last5) AS t, "
                "       replace(CAST(payment_last5 AS TEXT), '.0', '') AS stripped, "
                "       payment_at "
                "FROM shipment_requests "
                "WHERE COALESCE(payment_last5,'') <> '' "
                "  AND length(replace(CAST(payment_last5 AS TEXT), '.0', '')) < 5 "
                "  AND replace(CAST(payment_last5 AS TEXT), '.0', '') GLOB '[0-9]*' "
                "ORDER BY id DESC LIMIT 50"
            ).fetchall()
        ]
        ambiguous_total = conn.execute(
            "SELECT COUNT(*) AS c FROM shipment_requests "
            "WHERE COALESCE(payment_last5,'') <> '' "
            "  AND length(replace(CAST(payment_last5 AS TEXT), '.0', '')) < 5 "
            "  AND replace(CAST(payment_last5 AS TEXT), '.0', '') GLOB '[0-9]*'"
        ).fetchone()["c"]

        # 5) 從 operation_logs 還原原值（ground truth，優先於任何推斷）
        #    欄位 affinity 是 REAL，寫入的字串被 SQLite 轉成數字、前導零被吃掉；
        #    但 operation_logs.detail 是真正的 TEXT，管理員確認付款當下的原始輸入完整保存。
        by_req, parsed_logs = _last5_log_originals(conn)
        _num_core = _last5_num_core

        cur_map = {}
        req_ids = list(by_req.keys())
        for i in range(0, len(req_ids), 500):     # 分批避開 SQLite 變數上限
            chunk = req_ids[i:i + 500]
            ph = ",".join(["?"] * len(chunk))
            for cr in conn.execute(
                f"SELECT id, g_code, CAST(payment_last5 AS TEXT) AS cur_text, "
                f"       typeof(payment_last5) AS t "
                f"FROM shipment_requests WHERE id IN ({ph})", chunk
            ).fetchall():
                cur_map[cr["id"]] = dict(cr)

        would_recover, conflicts = [], []
        for rid in sorted(by_req.keys(), reverse=True):
            cur = cur_map.get(rid)
            if not cur:
                continue
            logged = by_req[rid]
            cur_text = cur["cur_text"] if cur["cur_text"] is not None else ""
            item = {
                "id": rid,
                "g_code": cur["g_code"],
                "current": cur_text,          # 已 CAST AS TEXT
                "current_typeof": cur["t"],
                "logged": logged,             # operation_logs 內的原始輸入
                "same": cur_text == logged,
            }
            if len(would_recover) < 100:
                would_recover.append(item)
            if _num_core(cur_text) != _num_core(logged):
                conflicts.append(item)

        # 有末五碼、但完全找不到對應操作紀錄的單（客戶自報路徑，只能靠推斷）
        no_log_total = 0
        no_log_samples = []
        for nr in conn.execute(
            "SELECT id, g_code, CAST(payment_last5 AS TEXT) AS cur_text, typeof(payment_last5) AS t "
            "FROM shipment_requests WHERE COALESCE(payment_last5,'') <> '' ORDER BY id DESC"
        ).fetchall():
            if nr["id"] in by_req:
                continue
            no_log_total += 1
            if len(no_log_samples) < 50:
                no_log_samples.append({"id": nr["id"], "g_code": nr["g_code"],
                                       "current": nr["cur_text"], "current_typeof": nr["t"]})

        conn.close()

        return jsonify({
            "success": True,
            "read_only": True,
            "schema": schema,
            "storage_classes": storage,
            "timeline": timeline,
            "ambiguous": {"total": ambiguous_total, "shown": len(ambiguous), "samples": ambiguous},
            "recovered": {
                "total_logs": parsed_logs,
                "unique_requests": len(by_req),
                "matched": len(cur_map),
                "differs": sum(1 for x in would_recover if not x["same"]),
                "would_recover": would_recover,
                "conflicts": {"total": len(conflicts), "samples": conflicts[:50]},
            },
            "no_log": {"total": no_log_total, "shown": len(no_log_samples), "samples": no_log_samples},
        })
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500


# ============ 最低計費重量 2kg 未套用：唯讀範圍診斷 ============
# 2026/09/01 起最低計費重量 1kg → 2kg（commit 0a3a45c），但正式站出現 9 月的單
# billed_weight 仍是 1，代表存檔當下瀏覽器跑的是舊版 admin.html。
# 這裡只量「有幾張、少收多少」，不改任何資料。
#
# ⚠️ 下面的重算是逐行對照 templates/admin.html 的 calcBoxes() / calcHandling()
#    （全站只有這兩處算計費重量）翻成 Python，不是憑公式說明重寫。
#    改前端規則時這裡要跟著改，否則診斷會失真。

BILLING_MIN_KG = 1.0          # FWT JAPAN：最低 1kg，超過照實際重量（小數點後兩位）計費
BILLING_RULE_SINCE = "2026-09-01"


def _js_parse_float(v):
    """模擬 JS 的 parseFloat(v) || 0：前綴能解析成數字就取，否則 0。"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    m = re.match(r"\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?", str(v))
    if not m:
        return 0.0
    try:
        return float(m.group(0))
    except ValueError:
        return 0.0


def _js_round(x):
    """模擬 JS Math.round（.5 一律往正無限大，Python round 是四捨六入五成雙）。"""
    return int(math.floor(x + 0.5))


def _js_parse_int(v):
    """模擬 JS parseInt(v) || 0（截斷小數；解析失敗 0）。"""
    return int(_js_parse_float(v))


def _calc_box_billed(box, rate, min_kg=BILLING_MIN_KG):
    """對照 admin.html calcBoxes() 內的單箱迴圈：回 (billed, sub)。

    JS 原文：
        var weight = parseFloat(b.actual_weight) || 0;
        var vL = parseFloat(b.length)||0, vW = parseFloat(b.width)||0, vH = parseFloat(b.height)||0;
        var raw = weight;
        if (vL>0 && vW>0 && vH>0) {
            var vol = (vL*vW*vH)/6000;
            var over3 = vol > weight*3;
            raw = over3 ? Math.max(weight, vol) : weight;
        }
        var billed = weight<=0 ? 0 : (raw<2 ? 2 : Math.ceil(raw*2)/2);
        var longFee = (vlong_over(vL,vW,vH)) ? 86 : 0;   // max(L,W,H) > 150
        var sub = Math.round(billed*rate) + longFee;
    """
    weight = _js_parse_float(box.get("actual_weight"))
    vL = _js_parse_float(box.get("length"))
    vW = _js_parse_float(box.get("width"))
    vH = _js_parse_float(box.get("height"))
    raw = weight  # FWT JAPAN：只算實際重量、不算材積
    if weight <= 0:
        billed = 0.0
    else:
        billed = float(min_kg) if raw < min_kg else round(raw, 2)
    sub = _js_round(billed * rate)
    return billed, sub


def _calc_handling(hw, min_kg=BILLING_MIN_KG):
    """對照 admin.html calcHandling()：

        if (hw <= 0) handling = 0;
        else {
            var billedHw = hw < 2 ? 2 : Math.ceil(hw * 2) / 2;
            var wholeKg = Math.floor(billedHw);
            var hasHalf = (billedHw - wholeKg) >= 0.5 ? 1 : 0;
            var handling = wholeKg * 11 + hasHalf * 6;
        }
    """
    return 0  # FWT JAPAN：不收理貨費


def _is_half_kg_multiple(x):
    return abs(x * 2 - round(x * 2)) < 1e-6


def _billing_recalc_row(rd, min_kg=BILLING_MIN_KG):
    """用現行規則重算一張單（純函式，不碰 DB）。

    回 dict：mode / recalc_billed / recalc_shipping / recalc_handling / recalc_total / flagged。
    total 只換運費與理貨兩項，合箱費、信件費、加值服務照存檔值帶入
    （calcTotal = shipping + handling + consolidation + letter + extras）。
    """
    stored_billed = float(rd.get("billed_weight") or 0)
    stored_shipping = float(rd.get("shipping_fee") or 0)
    stored_handling = float(rd.get("handling_fee") or 0)
    stored_total = float(rd.get("total_fee") or 0)
    rate = _js_parse_int(rd.get("rate_per_kg"))
    boxes = [b for b in _parse_boxes(rd.get("boxes_json")) if isinstance(b, dict)]
    weighted = [b for b in boxes if _js_parse_float(b.get("actual_weight")) > 0]

    if weighted:
        # 新制多箱：逐箱重算（每箱各自套下限、各自 0.5 進位），總計費重量 = 各箱加總
        mode = "boxes"
        total_billed, total_shipping = 0.0, 0
        for b in boxes:
            billed, sub = _calc_box_billed(b, rate, min_kg)
            total_billed += billed
            total_shipping += sub
        recalc_billed = float(f"{total_billed:.1f}")   # JS: totalBilled.toFixed(1) 再 parseFloat
        recalc_shipping = total_shipping
        shipping_residual = 0
    else:
        # 舊制單箱（沒有 boxes_json 或箱子都沒重量）：只有 billed_weight 可用，
        # 拿存檔的計費重量再套一次現行規則（舊值本來就 ≥ 實重，套規則不會變小）。
        # 運費：存檔運費扣掉 round(舊計費重 × 費率) 的餘額（特長件 86 等）原樣保留。
        mode = "legacy"
        if stored_billed <= 0:
            recalc_billed = 0.0
        else:
            recalc_billed = float(min_kg) if stored_billed < min_kg else round(stored_billed, 2)
        shipping_residual = int(stored_shipping - _js_round(stored_billed * rate))
        recalc_shipping = _js_round(recalc_billed * rate) + shipping_residual

    # calcBoxes 尾端：理貨重量預設 = 總計費重量 → calcHandling()
    recalc_handling = _calc_handling(recalc_billed, min_kg)
    recalc_total = _js_round(stored_total - stored_shipping - stored_handling + recalc_shipping + recalc_handling)

    weight_mismatch = abs(recalc_billed - stored_billed) > 0.005
    not_half_kg = False  # FWT JAPAN：不做 0.5kg 進位，不檢查
    return {
        "mode": mode,
        "box_count": len(boxes),
        "rate": rate,
        "recalc_billed": recalc_billed,
        "recalc_shipping": recalc_shipping,
        "recalc_handling": recalc_handling,
        "recalc_total": recalc_total,
        "shipping_residual": shipping_residual,
        "flagged": weight_mismatch or not_half_kg,
        "reason": ("計費重量與現行規則不符" if weight_mismatch else "") +
                  ("；" if weight_mismatch and not_half_kg else "") +
                  ("存檔計費重量不是 0.5 倍數" if not_half_kg else ""),
    }


def _billing_min_weight_scan(conn, min_kg=BILLING_MIN_KG, since=BILLING_RULE_SINCE):
    """掃 shipment_requests：total_fee > 0 且 exported_at / updated_at 在規則生效日之後的單。
    只做 SELECT。min_kg 參數只給測試用（把下限改回 1 應該掃不到任何單）。"""
    rows = conn.execute(
        "SELECT id, g_code, customer_name, status, created_at, updated_at, exported_at, "
        "       billed_weight, rate_per_kg, shipping_fee, handling_fee, "
        "       consolidation_fee, letter_fee, total_fee, "
        "       payment_last5, payment_at, boxes_json "
        "FROM shipment_requests "
        "WHERE COALESCE(total_fee, 0) > 0 "
        "  AND (COALESCE(exported_at,'') >= ? OR COALESCE(updated_at,'') >= ?) "
        "ORDER BY updated_at DESC, id DESC",
        (since, since)
    ).fetchall()

    # 誰按的出貨：operation_logs（出貨處理 / 出貨單#id），取最後一筆
    op_by_req = {}
    for lr in conn.execute(
        "SELECT operator, role, detail, created_at FROM operation_logs "
        "WHERE action='出貨處理' AND created_at >= ? ORDER BY id ASC",
        (since,)
    ).fetchall():
        m = re.search(r"出貨單#(\d+)", lr["detail"] or "")
        if m:
            op_by_req[int(m.group(1))] = {"operator": lr["operator"], "role": lr["role"],
                                           "at": lr["created_at"]}

    items = []
    scanned = 0
    for r in rows:
        rd = dict(r)
        scanned += 1
        calc = _billing_recalc_row(rd, min_kg)
        if not calc["flagged"]:
            continue
        stored_billed = float(rd.get("billed_weight") or 0)
        stored_total = float(rd.get("total_fee") or 0)
        paid = bool((str(rd.get("payment_last5") or "")).strip())
        # 舊規則（下限 1kg）算出來是否剛好等於存檔值 → 直接證明是舊版 JS 存的
        old = _billing_recalc_row(rd, 1.0)
        matches_old_rule = (abs(old["recalc_billed"] - stored_billed) < 0.005
                            and old["recalc_total"] == _js_round(stored_total))
        billed_day = (rd.get("updated_at") or rd.get("exported_at") or "")[:10]
        op = op_by_req.get(rd["id"])
        items.append({
            "id": rd["id"],
            "g_code": rd.get("g_code"),
            "customer_name": rd.get("customer_name"),
            "status": rd.get("status"),
            "日期": billed_day,
            "updated_at": rd.get("updated_at"),
            "exported_at": rd.get("exported_at"),
            "箱數": calc["box_count"],
            "mode": calc["mode"],
            "rate_per_kg": calc["rate"],
            "現存_billed_weight": stored_billed,
            "現存_shipping_fee": float(rd.get("shipping_fee") or 0),
            "現存_handling_fee": float(rd.get("handling_fee") or 0),
            "現存_total_fee": stored_total,
            "重算_billed_weight": calc["recalc_billed"],
            "重算_shipping_fee": calc["recalc_shipping"],
            "重算_handling_fee": calc["recalc_handling"],
            "重算_total_fee": calc["recalc_total"],
            "差額": calc["recalc_total"] - _js_round(stored_total),
            "是否已付款": paid,
            "payment_at": rd.get("payment_at") or "",
            "符合舊規則": matches_old_rule,
            "operator": (op or {}).get("operator", ""),
            "operator_role": (op or {}).get("role", ""),
            "reason": calc["reason"],
        })

    paid_items = [x for x in items if x["是否已付款"]]
    unpaid_items = [x for x in items if not x["是否已付款"]]
    by_date, by_operator = {}, {}
    for x in items:
        d = by_date.setdefault(x["日期"], {"count": 0, "shortfall": 0})
        d["count"] += 1
        d["shortfall"] += x["差額"]
        o = by_operator.setdefault(x["operator"] or "(無操作紀錄)", {"count": 0, "shortfall": 0})
        o["count"] += 1
        o["shortfall"] += x["差額"]

    return {
        "rule": {"min_kg": min_kg, "since": since,
                 "source": "templates/admin.html calcBoxes()/calcHandling()"},
        "scanned": scanned,
        "affected_count": len(items),
        "total_shortfall": sum(x["差額"] for x in items),
        "已付款": {"count": len(paid_items), "shortfall": sum(x["差額"] for x in paid_items)},
        "未付款": {"count": len(unpaid_items), "shortfall": sum(x["差額"] for x in unpaid_items)},
        "符合舊規則筆數": sum(1 for x in items if x["符合舊規則"]),
        "負差額筆數": sum(1 for x in items if x["差額"] < 0),
        "依日期": [{"日期": k, **v} for k, v in sorted(by_date.items())],
        "依操作者": [{"operator": k, **v} for k, v in sorted(by_operator.items(), key=lambda kv: -kv[1]["count"])],
        "items": items,
    }


@app.route("/api/admin/maintenance/billing_min_weight_diag", methods=["GET"])
def admin_billing_min_weight_diag():
    """最低計費重量 2kg 未套用的單：範圍診斷（唯讀）。

    ⚠️ 本端點只做 SELECT，沒有任何寫入語句，
       也不重算、不改任何帳單——只回報「有幾張、少收多少、集中在哪幾天／誰」。
    """
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以執行維運作業"}), 403

    conn = get_db()
    try:
        result = _billing_min_weight_scan(conn)
        conn.close()
        return jsonify({"success": True, "read_only": True, **result})
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500


# ============ 階梯運費（按團主上月累計 kg 決定本月費率）============
# 規則：某 g_code「上月」累計 total_kg 落在哪一階，「本月」整月的每公斤運費就是該階。
#       當月帳單一出即定、不追溯回算；新團主沒有上月資料 → 給第一階（最貴）。
#
# 階梯值存 admin_settings 由老闆調整，程式碼裡不寫死（比照 BOX_LIMIT 的作法）。
# 邊界一律取「小於等於」：kg<=30→220、<=60→210、<=100→200、<=200→190、其餘→180。
# 所以 30.5kg 落在 210 那階（規格寫「31–60」時沒定義的小數區間，在這裡被補完）。
#
# 方案B（維持抽成）退路：shipping_rate_mode='flat' → _effective_rate_for() 直接回舊邏輯
# （會員專屬費率 / 代理 min_rate），其餘程式碼一行都不必改。

RATE_MODE_KEY = "shipping_rate_mode"          # 'tier' = 階梯制；'flat' = 方案B 舊單一費率
RATE_TIERS_KEY = "shipping_rate_tiers"        # JSON: [{"max_kg": 30, "rate": 220}, ...]
MARGIN_COST_KEY = "gross_margin_cost_per_kg"  # 集運自家成本，算毛利用

# 注意：這個 140 跟 app.py 裡代理分潤的 base_cost=180 是兩個不同概念，不要合併。
# 180 是「甲方批發成本」，寫在代理合約第六條，動它等於改合約。
# 140 是我們自己的集運成本，只影響毛利報表的呈現。
MARGIN_COST_DEFAULT = 140.0

# max_kg=None 代表最後一階（無上限）。排序由 _get_rate_tiers 保證。
RATE_TIERS_DEFAULT = [
    {"max_kg": 30,   "rate": 220},
    {"max_kg": 60,   "rate": 210},
    {"max_kg": 100,  "rate": 200},
    {"max_kg": 200,  "rate": 190},
    {"max_kg": None, "rate": 180},
]


def _get_rate_mode():
    """'tier'（階梯制）或 'flat'（方案B）。值壞掉一律回 'flat'，
    因為 flat 是舊行為 —— 設定壞掉時要退回「跟這功能上線前一樣」，不是退回新規則。"""
    v = (_get_setting(RATE_MODE_KEY, "") or "").strip().lower()
    return "tier" if v == "tier" else "flat"


def _normalize_rate_tiers(tiers):
    """驗證＋正規化階梯表。回 (tiers, None) 或 (None, 錯誤訊息)。

    要求：至少一階、rate 皆 > 0、max_kg 遞增、且最後一階必須是無上限（max_kg=None）。
    最後一階沒有無上限的話，超過最高門檻的重量會無費率可套 —— 寧可存不進去也不要
    在出帳當下才炸。"""
    if not isinstance(tiers, list) or not tiers:
        return None, "階梯表至少要有一階"
    out = []
    for t in tiers:
        if not isinstance(t, dict):
            return None, "每一階必須是物件，例：{\"max_kg\": 30, \"rate\": 220}"
        raw_max = t.get("max_kg", None)
        try:
            rate = float(t.get("rate"))
        except (ValueError, TypeError):
            return None, "rate 必須是數字"
        if rate <= 0:
            return None, "rate 必須大於 0"
        if raw_max in (None, "", "null"):
            max_kg = None
        else:
            try:
                max_kg = float(raw_max)
            except (ValueError, TypeError):
                return None, "max_kg 必須是數字或 null（代表最後一階無上限）"
            if max_kg <= 0:
                return None, "max_kg 必須大於 0"
        out.append({"max_kg": max_kg, "rate": rate})

    capped = [t for t in out if t["max_kg"] is not None]
    uncapped = [t for t in out if t["max_kg"] is None]
    if len(uncapped) != 1:
        return None, "必須剛好有一階 max_kg=null（最後一階，無上限）"
    capped.sort(key=lambda t: t["max_kg"])
    for a, b in zip(capped, capped[1:]):
        if a["max_kg"] == b["max_kg"]:
            return None, f"max_kg 重複：{a['max_kg']}"
    return capped + uncapped, None


def _get_rate_tiers():
    """階梯表；設定缺漏或壞掉時回預設值，不讓出帳卡住（比照 _get_box_limit）。"""
    raw = _get_setting(RATE_TIERS_KEY, "")
    if raw:
        try:
            tiers, err = _normalize_rate_tiers(json.loads(raw))
            if tiers and not err:
                return tiers
            print(f"[rate_tiers] ⚠️ 設定值不合法（{err}），改用預設階梯", flush=True)
        except Exception as e:
            print(f"[rate_tiers] ⚠️ 設定值解析失敗（{e}），改用預設階梯", flush=True)
    return [dict(t) for t in RATE_TIERS_DEFAULT]


def _rate_for_kg(kg, tiers=None):
    """依累計 kg 取每公斤費率。邊界為「小於等於」：30.5 → 210。"""
    tiers = tiers or _get_rate_tiers()
    try:
        kg = float(kg or 0)
    except (ValueError, TypeError):
        kg = 0.0
    for t in tiers:
        if t["max_kg"] is None or kg <= t["max_kg"]:
            return float(t["rate"])
    return float(tiers[-1]["rate"])   # _normalize 保證有無上限階，正常到不了這行


def _get_margin_cost():
    """集運毛利的每公斤成本；壞值回預設 140。"""
    try:
        v = float(str(_get_setting(MARGIN_COST_KEY, "") or MARGIN_COST_DEFAULT).strip())
        return v if v >= 0 else MARGIN_COST_DEFAULT
    except (ValueError, TypeError):
        return MARGIN_COST_DEFAULT


def _get_tenant_row(g_code, conn=None):
    """查 GoyouLink 團主對照表；只認 status='active'。不是團主回 None。

    階梯運費的「適用對象」判定就靠這支。任何異常（表不存在、SQL 出錯）一律回
    None＝視為散客＝維持原費率。失敗方向刻意偏向「漏套」而不是「誤套」：
    階梯最高階 220，散客若無累計量會被判 220，多半高於他原本的費率，
    等於對真實客戶超收。寧可階梯沒生效，也不能讓散客被誤套。

    conn 可傳入既有連線（_effective_rate_for 已經開著一條，不要再開巢狀的）。"""
    g_code = (g_code or "").strip().upper()
    if not g_code:
        return None
    own = conn is None
    try:
        c = conn or get_db()
        try:
            return c.execute(
                "SELECT * FROM goyoulink_tenants WHERE UPPER(TRIM(g_code))=? AND status='active'",
                (g_code,)
            ).fetchone()
        finally:
            if own:
                c.close()
    except Exception as e:
        print(f"[tenant] {g_code} 查詢失敗（視為散客）: {e}", flush=True)
        return None


def _is_goyoulink_tenant(g_code, conn=None):
    return _get_tenant_row(g_code, conn) is not None


def _tenant_basis_kg(conn, g_code, ym, since_at):
    """算某團主在 ym 這個月、且「成為團主之後」的累計計費重量。

    與 monthly_kg_snapshots.total_kg 的差別：這裡多了 updated_at >= since_at 的條件。
    剛登錄的團主，登錄前那段是散客時期的量，不算團主貢獻，也避免
    「先以散客身分衝量、再登錄直接套最低價」。所以兩個數字可能不同，
    這是刻意的；存進 rate_basis_kg 當存證。"""
    ex_sql, ex_params = _stats_exclude_sql()
    row = conn.execute(
        f"""SELECT COALESCE(SUM(billed_weight), 0) AS kg FROM shipment_requests
             WHERE status='已出貨' AND UPPER(TRIM(g_code))=?
               AND COALESCE(updated_at,'') != ''
               AND substr(updated_at, 1, 7) = ?
               AND updated_at >= ?{ex_sql}""",
        [g_code.strip().upper(), ym, since_at] + ex_params
    ).fetchone()
    return float(row["kg"] or 0)


def _prev_ym(ym):
    """'2026-09' → '2026-08'。格式不對回 ''。"""
    try:
        y, m = ym.split("-")
        y, m = int(y), int(m)
        if m == 1:
            return f"{y - 1}-12"
        return f"{y}-{m - 1:02d}"
    except (ValueError, AttributeError):
        return ""


def _current_ym():
    return datetime.now().strftime("%Y-%m")


def _refresh_monthly_snapshot(conn, ym, g_code=None):
    """重算某年月（可限定單一 g_code）的衍生統計並 upsert。

    只碰衍生欄位，applied_rate / rate_locked_at / rate_basis_kg 一律不動 ——
    那是凍結值，只有 _effective_rate_for() 能寫，且只寫一次。
    冪等：同樣的 shipment_requests 重跑幾次結果都一樣。

    歸月依 updated_at（出貨處理時間），與 /api/admin/stats/monthly 同一個欄位。
    統計排除清單（測試帳號）在這裡就套用，避免 G8888 之類的量墊高階梯。"""
    ex_sql, ex_params = _stats_exclude_sql()
    where = ["status='已出貨'", "COALESCE(updated_at,'') != ''",
             "substr(updated_at, 1, 7) = ?"]
    params = [ym]
    if g_code:
        where.append("UPPER(TRIM(g_code)) = ?")
        params.append(g_code.strip().upper())
    sql = f"SELECT * FROM shipment_requests WHERE {' AND '.join(where)}{ex_sql}"
    rows = conn.execute(sql, params + ex_params).fetchall()

    cost = _get_margin_cost()
    acc = {}
    for row in rows:
        r = dict(row)
        gc = (r.get("g_code") or "").strip().upper()
        if not gc:
            continue
        a = acc.setdefault(gc, {"total_kg": 0.0, "paid_kg": 0.0, "cnt": 0,
                                "shipping": 0.0, "total": 0.0})
        kg = float(r.get("billed_weight") or 0)
        a["total_kg"] += kg
        a["cnt"] += 1
        a["shipping"] += float(r.get("shipping_fee") or 0)
        a["total"] += float(r.get("total_fee") or 0)
        # paid_kg 的口徑刻意與月報統計一致：有填匯款末五碼才算已收款
        if (r.get("payment_last5") or "").strip() and float(r.get("total_fee") or 0) > 0:
            a["paid_kg"] += kg

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for gc, a in acc.items():
        margin = a["shipping"] - cost * a["total_kg"]
        conn.execute(
            """INSERT INTO monthly_kg_snapshots
                   (g_code, ym, total_kg, paid_kg, shipment_count, shipping_fee,
                    total_fee, gross_margin, cost_per_kg, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(g_code, ym) DO UPDATE SET
                   total_kg=excluded.total_kg, paid_kg=excluded.paid_kg,
                   shipment_count=excluded.shipment_count,
                   shipping_fee=excluded.shipping_fee, total_fee=excluded.total_fee,
                   gross_margin=excluded.gross_margin, cost_per_kg=excluded.cost_per_kg,
                   updated_at=excluded.updated_at""",
            (gc, ym, round(a["total_kg"], 2), round(a["paid_kg"], 2), a["cnt"],
             round(a["shipping"]), round(a["total"]), round(margin), cost, now)
        )

    # 該月原本有列、現在算不出資料（單子被 revert / 被排除）→ 衍生欄位歸零。
    # 不刪列：applied_rate 可能已凍結，刪掉會讓同月後續帳單重新取得費率。
    stale_where = ["ym=?"]
    stale_params = [ym]
    if g_code:
        stale_where.append("g_code=?")
        stale_params.append(g_code.strip().upper())
    elif acc:
        ph = ",".join(["?"] * len(acc))
        stale_where.append(f"g_code NOT IN ({ph})")
        stale_params.extend(acc.keys())
    if not (g_code and g_code.strip().upper() in acc):
        conn.execute(
            f"""UPDATE monthly_kg_snapshots
                   SET total_kg=0, paid_kg=0, shipment_count=0, shipping_fee=0,
                       total_fee=0, gross_margin=0, updated_at=?
                 WHERE {' AND '.join(stale_where)}""",
            [now] + stale_params
        )
    return len(acc)


def _resync_snapshot(g_code, *yms):
    """出貨/還原/收款狀態變動後，把受影響月份的快照重算一次。

    自開連線（呼叫點通常剛 close 掉自己的寫入連線，避免兩個寫入交疊卡 lock）。
    永不拋例外：快照是衍生資料，算失敗頂多數字舊一點，
    絕不能讓出貨或收款這種主流程因此失敗。隨時可用 snapshot/refresh 補算。"""
    targets = {y for y in yms if y and re.fullmatch(r"\d{4}-\d{2}", y)}
    if not targets:
        return
    try:
        conn = get_db()
        try:
            for y in targets:
                _refresh_monthly_snapshot(conn, y, g_code)
            conn.commit()
        finally:
            conn.close()
    except Exception as e:
        print(f"[snapshot] 重算失敗 {g_code} {sorted(targets)}: {e}", flush=True)


def _effective_rate_for(g_code, ym=None):
    """某團主在某年月該套用的每公斤費率。

    回 dict：rate / mode / is_tenant / source / basis_kg / locked_at / tiers。
      source='locked'      已凍結，直接回存檔值（同月第二張之後的單都走這裡）
      source='tier'        依「上月、成為團主之後」的累計量決定並凍結
      source='tenant_new'  登錄當月的團主 → 一律第一階（不吃登錄前的散客量）
      source='flat'        方案B：回舊的會員/代理費率，不碰快照
      source='not_tenant'  階梯制已啟用，但這個 g_code 不是團主 → 維持原費率

    兩道關卡都要過才套階梯：① 全站 mode='tier' ② 該 g_code 是 active 團主。
    全站模式只是「啟用功能」，不代表所有人都套 —— 散客一律維持原費率。

    凍結是這支函式唯一的副作用，且用條件式 UPDATE（applied_rate<=0 才寫），
    兩個 gunicorn worker 同時開帳單視窗也只會有一個寫進去。
    非團主完全不寫 applied_rate，不佔用凍結欄位。"""
    g_code = (g_code or "").strip().upper()
    ym = ym or _current_ym()
    out = {"g_code": g_code, "ym": ym, "mode": _get_rate_mode(), "is_tenant": False,
           "rate": 0.0, "source": "", "basis_kg": 0.0, "locked_at": "",
           "tiers": _get_rate_tiers()}
    if not g_code:
        out["source"] = "not_tenant"
        out["rate"] = 0.0
        return out

    # ── 關卡①：方案B → 完全走舊路徑（會員專屬費率 > 代理 min_rate > 環境預設）
    if out["mode"] == "flat":
        out["source"] = "flat"
        out["rate"] = float(_flat_rate_for(g_code))
        return out

    conn = get_db()
    try:
        # ── 關卡②：不是 active 團主 → 散客，維持原費率，不碰快照、不凍結
        tenant = _get_tenant_row(g_code, conn)
        if tenant is None:
            out.update(is_tenant=False, source="not_tenant",
                       rate=float(_flat_rate_for(g_code)))
            return out
        out["is_tenant"] = True
        out["tenant_key"] = tenant["tenant_key"]
        out["display_name"] = tenant["display_name"] or ""

        row = conn.execute(
            "SELECT applied_rate, rate_locked_at, rate_basis_kg FROM monthly_kg_snapshots "
            "WHERE g_code=? AND ym=?", (g_code, ym)
        ).fetchone()
        if row and float(row["applied_rate"] or 0) > 0:
            out.update(rate=float(row["applied_rate"]), source="locked",
                       basis_kg=float(row["rate_basis_kg"] or 0),
                       locked_at=row["rate_locked_at"] or "")
            return out

        # ── 決定本月費率 ──
        # 登錄當月：一律第一階。登錄前的量是散客時期的，不是團主貢獻，
        # 也避免「先以散客身分衝量、再登錄直接套最低價」。
        reg_at = (tenant["created_at"] or "").strip()
        reg_ym = reg_at[:7]
        prev = _prev_ym(ym)
        basis_kg = 0.0
        if (not reg_ym) or ym <= reg_ym:
            src = "tenant_new"
        else:
            src = "tier"
            if prev:
                # 快照照樣重算（貢獻報表要用），但費率的依據另外算：
                # 多一個 updated_at >= 登錄時間 的條件，見 _tenant_basis_kg。
                _refresh_monthly_snapshot(conn, prev, g_code)
                conn.commit()
                basis_kg = _tenant_basis_kg(conn, g_code, prev, reg_at)

        rate = _rate_for_kg(basis_kg, out["tiers"])
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # 先確保本月列存在（新團主本月還沒出過貨 → 快照沒這列）
        conn.execute(
            "INSERT OR IGNORE INTO monthly_kg_snapshots (g_code, ym, cost_per_kg, updated_at) "
            "VALUES (?, ?, ?, ?)", (g_code, ym, _get_margin_cost(), now)
        )
        # 條件式 UPDATE：只有還沒凍結的列會被寫入
        conn.execute(
            "UPDATE monthly_kg_snapshots SET applied_rate=?, rate_locked_at=?, rate_basis_kg=? "
            "WHERE g_code=? AND ym=? AND COALESCE(applied_rate,0) <= 0",
            (rate, now, basis_kg, g_code, ym)
        )
        conn.commit()
        # 讀回實際落地的值：若剛好被另一個 worker 搶先，要用它寫的那個
        final = conn.execute(
            "SELECT applied_rate, rate_locked_at, rate_basis_kg FROM monthly_kg_snapshots "
            "WHERE g_code=? AND ym=?", (g_code, ym)
        ).fetchone()
        out.update(rate=float(final["applied_rate"] or rate),
                   basis_kg=float(final["rate_basis_kg"] or basis_kg),
                   locked_at=final["rate_locked_at"] or now,
                   source=src)
        return out
    finally:
        conn.close()


def _flat_rate_for(g_code):
    """方案B / 覆寫比對用的舊制費率：會員專屬 > 代理 min_rate > 環境預設。
    與 verify_customer 的取法一致。"""
    g_code = (g_code or "").strip().upper()
    conn = get_db()
    try:
        m = conn.execute("SELECT agent_id, shipping_rate FROM members WHERE g_code=?",
                         (g_code,)).fetchone()
        if m:
            own = float(m["shipping_rate"] or 0)
            if own > 0:
                return own
            ag = conn.execute("SELECT min_rate FROM agents WHERE id=?",
                              (m["agent_id"],)).fetchone()
            if ag and ag["min_rate"]:
                return float(ag["min_rate"])
            return float(DEFAULT_SHIPPING_RATE or 180)
        ms = conn.execute("SELECT shipping_rate FROM members_shopify WHERE g_code=?",
                          (g_code,)).fetchone()
        if ms and float(ms["shipping_rate"] or 0) > 0:
            return float(ms["shipping_rate"])
    except Exception as e:
        print(f"[flat_rate] {g_code} 查詢失敗: {e}", flush=True)
    finally:
        conn.close()
    return float(DEFAULT_SHIPPING_RATE or 180)


# ============ 申報人單日報關件數上限 ============
# 清關行規定同一位申報人、同一天可報關的件數有上限。各家不同且會變動，
# 所以存 admin_settings 由老闆調整，程式碼裡不寫死數字。
# 報關的單位是「出檔案批次」——一批交給清關行的就是同一天報的，
# 所以硬檢查點在出檔案，不在分箱。

BOX_LIMIT_KEY = "max_boxes_per_declarant_per_day"
BOX_LIMIT_DEFAULT = 3


def _get_box_limit():
    """申報人單日件數上限；設定值壞掉時回預設值，不讓出檔案卡住。"""
    try:
        v = int(str(_get_setting(BOX_LIMIT_KEY, "") or BOX_LIMIT_DEFAULT).strip())
        return v if 1 <= v <= 50 else BOX_LIMIT_DEFAULT
    except (ValueError, TypeError):
        return BOX_LIMIT_DEFAULT


def _declarant_key(box, ship):
    """(申報人姓名, 申報人電話)。取值順序與 vendors.build_rows 的三層 fallback
    完全一致：箱層級 → 出貨單主申報人 → 收件人。
    只用姓名會把同名不同人合併，所以 key 帶電話。"""
    name = (str(box.get("declarant_name") or "").strip()
            or str(ship.get("declarant_name") or "").strip()
            or (str(ship.get("ship_recipient")) if ship.get("ship_recipient") else ""))
    phone = (str(box.get("declarant_phone") or "").strip()
             or str(ship.get("declarant_phone") or "").strip()
             or (str(ship.get("ship_phone")) if ship.get("ship_phone") else ""))
    return name, phone


def _count_declarant_boxes(ships):
    """統計每位申報人的箱數。沒有箱資料的舊制單視為 1 箱
    （與 vendors.build_rows 合成一箱的行為一致）。"""
    counts = {}
    for sp in ships:
        boxes = sp.get("boxes") or [{}]
        for b in boxes:
            k = _declarant_key(b if isinstance(b, dict) else {}, sp)
            counts[k] = counts.get(k, 0) + 1
    return counts


def _same_day_declarant_counts(conn, day, exclude_ids=()):
    """同一天已出檔案的批次，每位申報人各報了幾箱。
    同一天出兩批、各自沒超過但加起來爆掉 —— 這一步就是為了抓那種情況。"""
    rows = conn.execute(
        "SELECT id, ship_recipient, ship_phone, declarant_name, declarant_phone, boxes_json "
        "FROM shipment_requests "
        "WHERE exported_at IS NOT NULL AND exported_at != '' AND date(exported_at) = ?",
        (day,)
    ).fetchall()
    ships = []
    for r in rows:
        rd = dict(r)
        if rd["id"] in exclude_ids:
            continue
        ships.append({
            "ship_recipient": _safe_str(rd.get("ship_recipient")),
            "ship_phone": _safe_str(rd.get("ship_phone")),
            "declarant_name": _safe_str(rd.get("declarant_name")),
            "declarant_phone": _safe_str(rd.get("declarant_phone")),
            "boxes": _parse_boxes(rd.get("boxes_json")),
        })
    return _count_declarant_boxes(ships)


def _check_declarant_box_limit(conn, shipments, limit=None, day=None):
    """回傳超過上限的申報人明細（空 list = 沒問題）。"""
    limit = _get_box_limit() if limit is None else limit
    day = day or datetime.now().strftime("%Y-%m-%d")
    batch = _count_declarant_boxes(shipments)
    already = _same_day_declarant_counts(conn, day, exclude_ids={sp.get("id") for sp in shipments})
    over = []
    for k, n in batch.items():
        prior = already.get(k, 0)
        if n + prior > limit:
            over.append({
                "declarant_name": k[0], "declarant_phone": k[1],
                "batch_boxes": n, "same_day_boxes": prior,
                "total": n + prior, "limit": limit,
            })
    over.sort(key=lambda x: (-x["total"], x["declarant_name"]))
    return over


@app.route("/api/admin/settings/box_limit", methods=["GET"])
def admin_get_box_limit():
    """申報人單日件數上限（老闆＋員工都能看，前端軟提示要用）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    return jsonify({"success": True, "limit": _get_box_limit(), "default": BOX_LIMIT_DEFAULT})


@app.route("/api/admin/settings/box_limit", methods=["PUT"])
def admin_set_box_limit():
    """修改上限（只有老闆）。各清關行規定不同且會變動，所以做成可調設定值。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以修改此設定"}), 403
    raw = (request.json or {}).get("limit")
    try:
        v = int(str(raw).strip())
    except (ValueError, TypeError, AttributeError):
        return jsonify({"success": False, "error": "上限必須是 1~50 的整數"}), 400
    if not (1 <= v <= 50):
        return jsonify({"success": False, "error": "上限必須是 1~50 的整數"}), 400
    old = _get_box_limit()
    _set_setting(BOX_LIMIT_KEY, str(v))
    log_op("修改申報人單日件數上限", "box_limit", f"{old} → {v}")
    return jsonify({"success": True, "limit": v, "old": old})


# ============ 統計排除客編（測試帳號不污染數字）============
# 老闆的測試帳號（預設 G8888）資料混在正式庫裡，統計會失真。
# 只排除 /api/admin/stats/* 的彙總；帳單列表、出檔案、會員、搜尋、對帳、客戶端一律照舊——
# 「不污染數字」與「看不見」是兩件事，測試帳號必須在日常操作中看得見才能拿來驗證系統。
# 做成設定值而非常數：之後還會開測試帳號，寫死就要再改程式。

STATS_EXCLUDED_KEY = "stats_excluded_codes"
STATS_EXCLUDED_DEFAULT = "G8888"
_STATS_CODE_RE = re.compile(r"^[A-Z0-9]+$")


def _normalize_stats_codes(raw):
    """逗號分隔 → 去空白、轉大寫、去重（保序）。含非英數字元 → 回 None（格式非法）。"""
    out = []
    for tok in str(raw or "").split(","):
        t = tok.strip().upper()
        if not t:
            continue
        if not _STATS_CODE_RE.match(t):
            return None
        if t not in out:
            out.append(t)
    return out


def _stats_excluded_codes():
    """目前要從統計排除的客編清單（已正規化）。設定壞掉時視為空清單，不讓統計頁掛掉。"""
    codes = _normalize_stats_codes(_get_setting(STATS_EXCLUDED_KEY, STATS_EXCLUDED_DEFAULT))
    return codes or []


def _stats_exclude_sql(alias=""):
    """回 (sql 片段, 參數)。清單為空 → ("", [])，行為與沒有這功能時完全一致。
    比對一律 UPPER(TRIM(...))，避免大小寫或空白造成漏排除。"""
    codes = _stats_excluded_codes()
    if not codes:
        return "", []
    col = f"{alias}.g_code" if alias else "g_code"
    ph = ",".join(["?"] * len(codes))
    return f" AND UPPER(TRIM(COALESCE({col},''))) NOT IN ({ph})", list(codes)


@app.route("/api/admin/settings/stats_excluded", methods=["GET"])
def admin_get_stats_excluded():
    """統計排除清單（老闆＋員工都能看：統計頁要標示「已排除：…」）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    return jsonify({"success": True, "codes": _stats_excluded_codes(),
                    "default": STATS_EXCLUDED_DEFAULT})


@app.route("/api/admin/settings/stats_excluded", methods=["PUT"])
def admin_set_stats_excluded():
    """修改統計排除清單（只有老闆）。格式：逗號分隔客編，只允許英數與逗號。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以修改此設定"}), 403
    raw = (request.json or {}).get("codes", "")
    if isinstance(raw, list):
        raw = ",".join(str(x) for x in raw)
    codes = _normalize_stats_codes(raw)
    if codes is None:
        return jsonify({"success": False, "error": "格式錯誤：只允許英數字與逗號，例如 G8888,G9999"}), 400
    old = ",".join(_stats_excluded_codes())
    new = ",".join(codes)
    _set_setting(STATS_EXCLUDED_KEY, new)
    log_op("修改統計排除客編", STATS_EXCLUDED_KEY, f"{old} → {new}")
    return jsonify({"success": True, "codes": codes, "old": old})


# ============ 階梯運費設定 / 查詢 API ============

@app.route("/api/admin/settings/rate_tiers", methods=["GET"])
def admin_get_rate_tiers():
    """階梯表與模式（老闆＋員工都能看：帳單視窗要顯示「本月階梯」）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    return jsonify({
        "success": True,
        "mode": _get_rate_mode(),
        "tiers": _get_rate_tiers(),
        "margin_cost_per_kg": _get_margin_cost(),
        "default_tiers": RATE_TIERS_DEFAULT,
        "default_margin_cost": MARGIN_COST_DEFAULT,
    })


@app.route("/api/admin/settings/rate_tiers", methods=["PUT"])
def admin_set_rate_tiers():
    """修改階梯表 / 切換方案B / 調整毛利成本（只有老闆）。

    改階梯不會動到任何「已凍結」的 applied_rate —— 那正是「不追溯回算」的意思。
    新階梯只對之後才第一次取得費率的（團主, 月份）生效。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以修改此設定"}), 403
    data = request.json or {}
    changes = []

    if "mode" in data:
        mode = str(data.get("mode") or "").strip().lower()
        if mode not in ("tier", "flat"):
            return jsonify({"success": False, "error": "mode 只能是 tier 或 flat"}), 400
        old = _get_rate_mode()
        if mode != old:
            _set_setting(RATE_MODE_KEY, mode)
            changes.append(f"模式 {old} → {mode}")

    if "tiers" in data:
        tiers, err = _normalize_rate_tiers(data.get("tiers"))
        if err:
            return jsonify({"success": False, "error": f"階梯表格式錯誤：{err}"}), 400
        old = json.dumps(_get_rate_tiers(), ensure_ascii=False)
        new = json.dumps(tiers, ensure_ascii=False)
        if old != new:
            _set_setting(RATE_TIERS_KEY, new)
            changes.append(f"階梯 {old} → {new}")

    if "margin_cost_per_kg" in data:
        try:
            cost = float(data.get("margin_cost_per_kg"))
        except (ValueError, TypeError):
            return jsonify({"success": False, "error": "毛利成本必須是數字"}), 400
        if cost < 0:
            return jsonify({"success": False, "error": "毛利成本不能為負"}), 400
        old = _get_margin_cost()
        if abs(cost - old) > 1e-9:
            _set_setting(MARGIN_COST_KEY, str(cost))
            changes.append(f"毛利成本 {old} → {cost}")

    if changes:
        log_op("修改階梯運費設定", RATE_TIERS_KEY, "；".join(changes))
    return jsonify({"success": True, "mode": _get_rate_mode(),
                    "tiers": _get_rate_tiers(),
                    "margin_cost_per_kg": _get_margin_cost(),
                    "changes": changes})


@app.route("/api/admin/billing/effective_rate", methods=["GET"])
def admin_effective_rate():
    """帳單視窗開啟時取「本月該套用的費率」。

    這支有副作用：階梯制下第一次查詢會把該（團主, 月份）的費率凍結起來。
    這是刻意的 —— 凍結點就是「開始為這個月開帳單」的那一刻。
    員工要能查（出貨是他們做的），但未登入不行。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    g_code = (request.args.get("g_code") or "").strip().upper()
    ym = (request.args.get("ym") or "").strip() or _current_ym()
    if not re.fullmatch(r"\d{4}-\d{2}", ym):
        return jsonify({"success": False, "error": "ym 格式須為 YYYY-MM"}), 400
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"}), 400
    try:
        info = _effective_rate_for(g_code, ym)
        info["flat_rate"] = _flat_rate_for(g_code)   # 供前端顯示「舊制會是多少」
        return jsonify({"success": True, **info})
    except Exception as e:
        print(f"[effective_rate] {g_code} {ym} 失敗: {e}", flush=True)
        # 取不到階梯費率不能讓出貨卡住：退回舊制費率，前端照常可改
        return jsonify({"success": True, "g_code": g_code, "ym": ym,
                        "mode": "flat", "source": "fallback", "is_tenant": False,
                        "rate": _flat_rate_for(g_code), "basis_kg": 0,
                        "locked_at": "", "tiers": [], "error": str(e)})


@app.route("/api/admin/stats/snapshot/refresh", methods=["POST"])
def admin_refresh_snapshot():
    """重算月結快照的衍生統計（不動已凍結的 applied_rate）。

    不帶 month → 重算本月與上月（日常用）。帶 month=YYYY-MM → 只重算該月。
    帶 all=1 → 重算 shipment_requests 裡出現過的所有月份（首次建表補歷史用）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以執行"}), 403
    month = (request.args.get("month") or "").strip()
    do_all = (request.args.get("all") or "") in ("1", "true", "yes")
    conn = get_db()
    try:
        if do_all:
            yms = [r["ym"] for r in conn.execute(
                "SELECT DISTINCT substr(updated_at, 1, 7) AS ym FROM shipment_requests "
                "WHERE status='已出貨' AND COALESCE(updated_at,'') != '' ORDER BY ym"
            ).fetchall() if r["ym"]]
        elif month:
            if not re.fullmatch(r"\d{4}-\d{2}", month):
                conn.close()
                return jsonify({"success": False, "error": "month 格式須為 YYYY-MM"}), 400
            yms = [month]
        else:
            cur = _current_ym()
            yms = [y for y in (_prev_ym(cur), cur) if y]
        done = {}
        for y in yms:
            done[y] = _refresh_monthly_snapshot(conn, y)
        conn.commit()
        conn.close()
        log_op("重算月結快照", ",".join(yms), f"{sum(done.values())} 個團主")
        return jsonify({"success": True, "months": done})
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/admin/stats/snapshot", methods=["GET"])
def admin_get_snapshot():
    """看某月的月結快照（驗收與除錯用；第 3 步的貢獻報表會用到同一批數字）。"""
    # 兩層：先擋未登入，再擋員工。只寫 is_staff() 會讓匿名請求穿過去
    # （匿名的 user_type 不是 admin，is_staff() 回 False）。
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    if is_staff():
        return jsonify({"success": False, "error": "權限不足"}), 403
    ym = (request.args.get("month") or "").strip() or _current_ym()
    if not re.fullmatch(r"\d{4}-\d{2}", ym):
        return jsonify({"success": False, "error": "month 格式須為 YYYY-MM"}), 400
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT * FROM monthly_kg_snapshots WHERE ym=? ORDER BY total_kg DESC, g_code",
            (ym,)
        ).fetchall()
        prev = _prev_ym(ym)
        prev_kg = {}
        if prev:
            for r in conn.execute(
                "SELECT g_code, total_kg FROM monthly_kg_snapshots WHERE ym=?", (prev,)
            ).fetchall():
                prev_kg[r["g_code"]] = float(r["total_kg"] or 0)
        out = []
        for r in rows:
            d = dict(r)
            d["prev_total_kg"] = prev_kg.get(d["g_code"], 0.0)
            d["tier_rate_now"] = _rate_for_kg(d["prev_total_kg"])
            out.append(d)
        conn.close()
        return jsonify({"success": True, "month": ym, "prev_month": prev,
                        "mode": _get_rate_mode(), "rows": out})
    except Exception as e:
        try:
            conn.close()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500


# ============ GoyouLink 團主對照表 ============
# 這張表同時是階梯運費的「適用對象」名單，所以增刪等於改計費對象 → 一律限老闆。

def _tenant_has_bill_this_month(conn, g_code, ym=None):
    """該 g_code 本月是否已經出過帳單（已出貨且有金額）。

    用來擋「月中登錄」：登錄前出的單用散客原費率、登錄後用階梯，
    同一團主同月兩種費率就是客訴來源。與階梯本身「上月量→本月價」的
    次月邏輯一致，建議次月生效。"""
    ym = ym or _current_ym()
    row = conn.execute(
        """SELECT COUNT(*) AS c, COALESCE(SUM(total_fee),0) AS amt
             FROM shipment_requests
            WHERE status='已出貨' AND UPPER(TRIM(g_code))=?
              AND COALESCE(updated_at,'') != '' AND substr(updated_at,1,7)=?
              AND COALESCE(total_fee,0) > 0""",
        (g_code.strip().upper(), ym)
    ).fetchone()
    return int(row["c"] or 0), float(row["amt"] or 0)


@app.route("/api/admin/goyoulink_tenants", methods=["GET"])
def admin_list_tenants():
    """列出所有 GoyouLink 團主對照。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    try:
        rows = [dict(r) for r in conn.execute(
            "SELECT * FROM goyoulink_tenants ORDER BY status, tenant_key").fetchall()]
        # 順手附上本月是否已有帳單，前端可提示「此人本月已出過單」
        for r in rows:
            n, amt = _tenant_has_bill_this_month(conn, r["g_code"])
            r["bills_this_month"] = n
            r["billed_amount_this_month"] = amt
        return jsonify({"success": True, "tenants": rows, "current_ym": _current_ym(),
                        "mode": _get_rate_mode()})
    finally:
        conn.close()


@app.route("/api/admin/goyoulink_tenants", methods=["POST"])
def admin_create_tenant():
    """登錄一個團主。

    ⚠️ created_at 不只是紀錄時間，它是「何時開始算團主」的依據：
    登錄當月一律第一階，次月起只採計 updated_at >= created_at 的出貨量。

    月中登錄且該月已有帳單 → 預設擋下（409），要硬做得帶 force=true。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    d = request.json or {}
    tenant_key = (d.get("tenant_key") or "").strip().lower()
    g_code = (d.get("g_code") or "").strip().upper()
    display_name = (d.get("display_name") or "").strip()
    if not tenant_key or not g_code:
        return jsonify({"success": False, "error": "tenant_key 與 g_code 皆為必填"}), 400
    if not re.fullmatch(r"[a-z0-9][a-z0-9_-]{0,39}", tenant_key):
        return jsonify({"success": False, "error": "tenant_key 只能用小寫英數與 - _，開頭須為英數"}), 400
    if not re.fullmatch(r"[A-Z0-9]{2,20}", g_code):
        return jsonify({"success": False, "error": "g_code 格式不正確"}), 400

    conn = get_db()
    try:
        if conn.execute("SELECT 1 FROM goyoulink_tenants WHERE tenant_key=?", (tenant_key,)).fetchone():
            return jsonify({"success": False, "error": f"tenant_key「{tenant_key}」已存在"}), 409
        if conn.execute("SELECT 1 FROM goyoulink_tenants WHERE UPPER(TRIM(g_code))=?", (g_code,)).fetchone():
            return jsonify({"success": False, "error": f"客編「{g_code}」已經對應到其他團主"}), 409

        n, amt = _tenant_has_bill_this_month(conn, g_code)
        force = d.get("force") in (True, 1, "1", "true", "True")
        if n > 0 and not force:
            return jsonify({
                "success": False, "needs_force": True,
                "bills_this_month": n, "billed_amount_this_month": amt,
                "error": (f"「{g_code}」本月（{_current_ym()}）已經出過 {n} 張帳單、"
                          f"共 NT${amt:,.0f}。現在登錄會造成同一個月前後兩種費率"
                          f"（先前的單是散客原費率，之後的單走階梯），容易產生客訴。\n"
                          f"建議下個月 1 號再登錄。確定要現在登錄請帶 force=true。")
            }), 409

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            """INSERT INTO goyoulink_tenants
                   (tenant_key, display_name, g_code, service_url, status, note, created_at)
               VALUES (?, ?, ?, ?, 'active', ?, ?)""",
            (tenant_key, display_name, g_code, (d.get("service_url") or "").strip(),
             (d.get("note") or "").strip(), now)
        )
        conn.commit()
        log_op("登錄 GoyouLink 團主", g_code,
               f"{tenant_key}（{display_name}）生效起算 {now}" + ("；強制月中登錄" if n > 0 else ""))
        return jsonify({"success": True, "tenant_key": tenant_key, "g_code": g_code,
                        "created_at": now, "forced": bool(n > 0)})
    finally:
        conn.close()


@app.route("/api/admin/goyoulink_tenants/<tenant_key>", methods=["PUT"])
def admin_update_tenant(tenant_key):
    """改團主資料。g_code 不給改 —— 換客編等於換計費對象，請刪掉重建，
    免得 created_at（生效起算時間）跟著舊客編留下來造成錯誤的階梯依據。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    tenant_key = (tenant_key or "").strip().lower()
    d = request.json or {}
    if "g_code" in d:
        return jsonify({"success": False,
                        "error": "不支援修改 g_code；請刪除後重新登錄（生效時間會重新起算）"}), 400
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM goyoulink_tenants WHERE tenant_key=?", (tenant_key,)).fetchone()
        if not row:
            return jsonify({"success": False, "error": "找不到該團主"}), 404
        sets, params, changes = [], [], []
        for col in ("display_name", "service_url", "note"):
            if col in d:
                sets.append(f"{col}=?"); params.append((d.get(col) or "").strip())
                changes.append(f"{col}: {row[col]} → {(d.get(col) or '').strip()}")
        if "status" in d:
            st = (d.get("status") or "").strip().lower()
            if st not in ("active", "disabled"):
                return jsonify({"success": False, "error": "status 只能是 active 或 disabled"}), 400
            sets.append("status=?"); params.append(st)
            changes.append(f"status: {row['status']} → {st}"
                           + ("（停用後該客編退回散客原費率）" if st == "disabled" else ""))
        if not sets:
            return jsonify({"success": False, "error": "沒有要修改的欄位"}), 400
        params.append(tenant_key)
        conn.execute(f"UPDATE goyoulink_tenants SET {','.join(sets)} WHERE tenant_key=?", params)
        conn.commit()
        log_op("修改 GoyouLink 團主", row["g_code"], f"{tenant_key}｜" + "；".join(changes))
        return jsonify({"success": True, "changes": changes})
    finally:
        conn.close()


@app.route("/api/admin/goyoulink_tenants/<tenant_key>", methods=["DELETE"])
def admin_delete_tenant(tenant_key):
    """移除團主登錄。該客編之後一律走散客原費率。

    已凍結的 applied_rate 不動 —— 那是已經開出去的帳單所依據的費率，
    不追溯回算是這套機制的基本前提。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    tenant_key = (tenant_key or "").strip().lower()
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM goyoulink_tenants WHERE tenant_key=?", (tenant_key,)).fetchone()
        if not row:
            return jsonify({"success": False, "error": "找不到該團主"}), 404
        conn.execute("DELETE FROM goyoulink_tenants WHERE tenant_key=?", (tenant_key,))
        conn.commit()
        log_op("刪除 GoyouLink 團主", row["g_code"], f"{tenant_key}（{row['display_name']}）")
        return jsonify({"success": True, "deleted": tenant_key, "g_code": row["g_code"]})
    finally:
        conn.close()


# ============ 客服手冊（內容存 admin_settings，不另建表）============

@app.route("/api/admin/handbook", methods=["GET"])
def admin_get_handbook():
    """客服手冊內容（老闆＋員工都能看）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    return jsonify({
        "success": True,
        "content": _get_setting("cs_handbook", ""),
        "updated_at": _get_setting("cs_handbook_at", ""),
    })


@app.route("/api/admin/handbook", methods=["PUT"])
def admin_save_handbook():
    """儲存客服手冊（只有老闆能改；員工端也不顯示編輯鈕，兩層都擋）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆可以編輯客服手冊"}), 403
    content = (request.json or {}).get("content")
    if content is None:
        return jsonify({"success": False, "error": "缺少內容"}), 400
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")   # TZ=Asia/Taipei
    _set_setting("cs_handbook", str(content))
    _set_setting("cs_handbook_at", now)
    log_op("編輯客服手冊", "handbook", "")
    return jsonify({"success": True, "updated_at": now})


# ============ 客戶端 API ============

@app.route("/api/verify_customer", methods=["POST"])
def verify_customer():
    data = request.json
    g_code = data.get("customer_id", "").strip().upper()
    password = data.get("password", "").strip()
    if not g_code:
        return jsonify({"success": False, "error": "請輸入會員編號"})
    if not password:
        return jsonify({"success": False, "error": "請輸入密碼"})
    # 沒有英文前綴 → 預設加 G（你的客戶）
    if not g_code[:1].isalpha():
        g_code = "G" + g_code
    password_clean = normalize_phone(password)

    # ===== 速率限制：視窗內失敗過多 → 一律擋下（不區分帳號存在與否、不回剩餘次數）=====
    _ip = _client_ip()
    if _login_locked(g_code, _ip):
        print(f"[login_lock] g_code={g_code} ip={_ip}", flush=True)
        return jsonify({"success": False, "error": "登入嘗試次數過多，請 15 分鐘後再試"}), 429

    # ===== 0) 停用名單檢查（集運系統層級）：被停用者一律擋下，整頁顯示停用訊息 =====
    try:
        _dconn = get_db()
        _drow = _dconn.execute("SELECT 1 FROM disabled_members WHERE g_code=?", (g_code,)).fetchone()
        _dconn.close()
        if _drow:
            return jsonify({"success": False, "disabled": True, "error": "您的帳號已停用，請聯繫客服"})
    except Exception:
        pass

    # ===== 1) 先查本地 members 表（代理建的會員）=====
    try:
        conn = get_db()
        row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
        if row:
            m = dict(row)
            # 狀態檢查
            if m.get("status") == "disabled":
                conn.close()
                return jsonify({"success": False, "error": "此會員帳號已停用，請聯絡您的代理"})
            # 比對電話（去除空白、橫線、+886/+81 等）
            stored_phone = normalize_phone(m.get("phone") or "")
            if stored_phone != password_clean:
                conn.close()
                return _login_fail(g_code, _ip, "密碼錯誤，請輸入您的手機號碼")
            # 找該代理（含品牌欄位）
            ag = conn.execute("SELECT * FROM agents WHERE id=?", (m["agent_id"],)).fetchone()
            conn.close()
            agent_min = float(ag["min_rate"]) if ag and ag["min_rate"] else float(DEFAULT_SHIPPING_RATE)
            # 會員專屬費率 > 0 → 用該費率；否則用代理 min_rate
            member_rate = float(m.get("shipping_rate") or 0)
            rate_twd = int(member_rate if member_rate > 0 else agent_min)
            branding = _branding_dict(ag) if ag else _branding_dict(None)
            session["cust_g_code"] = g_code       # 客戶登入狀態（客戶端 API 守門用）
            session.permanent = True
            _login_record(g_code, _ip, True)      # 成功 → 清掉該客編視窗內的失敗紀錄
            return jsonify({
                "success": True,
                "customer": {
                    "id": g_code,  # 本地會員無 Shopify customer_id，用 g_code
                    "g_code": g_code,
                    "name": m.get("name") or "會員",
                    "email": m.get("email") or "",
                    "phone": stored_phone,
                    "phone_raw": m.get("phone") or "",
                    "address": m.get("address") or "",
                    "shipping_rate_twd": rate_twd,
                    "shipping_rate_jpy": twd_to_jpy(rate_twd) if rate_twd else 0,
                    "source": "agent",
                    "agent_name": ag["name"] if ag else "",
                    "branding": branding,
                }
            })
        conn.close()
    except Exception as e:
        print(f"[verify_customer] 本地查詢失敗：{e}", flush=True)

    # ===== 2) 回退查 Shopify 會員（M2 起 get_all_goyoutati_customers 讀本地 members_shopify）=====
    try:
        customers = get_all_goyoutati_customers()
        c = next((x for x in customers if x["g_code"] == g_code), None)
        if c is None:
            # ===== 3) 本地都查無 → Shopify 單筆補撈（新客人剛貼完 metafield 的空窗期保險）=====
            # 已存在於本地的會員在上面就命中了，永遠不會走到這裡：正常登入不打 Shopify。
            c = _member_login_fallback(g_code)
        if c is not None:
            if c["phone"] and c["phone"] == password_clean:
                try:
                    rate_twd = int(c["shipping_rate"]) if c["shipping_rate"] else DEFAULT_SHIPPING_RATE
                except (ValueError, TypeError):
                    rate_twd = DEFAULT_SHIPPING_RATE
                rate_jpy = twd_to_jpy(rate_twd) if rate_twd else 0
                session["cust_g_code"] = g_code       # 客戶登入狀態（客戶端 API 守門用）
                session.permanent = True
                _login_record(g_code, _ip, True)      # 成功 → 清掉該客編視窗內的失敗紀錄
                return jsonify({
                    "success": True,
                    "customer": {
                        "id": c["customer_id"],
                        "g_code": g_code,
                        "name": c["name"] or "會員",
                        "email": c["email"],
                        "phone": c["phone"],
                        "phone_raw": c["phone_raw"],
                        "address": c.get("address", ""),
                        "shipping_rate_twd": rate_twd,
                        "shipping_rate_jpy": rate_jpy,
                        "source": "shopify",
                    }
                })
            else:
                return _login_fail(g_code, _ip, "密碼錯誤，請輸入您的手機號碼")
        return _login_fail(g_code, _ip, "找不到此會員編號，請確認後重試")
    except Exception as e:
        return jsonify({"success": False, "error": f"查詢失敗: {str(e)}"})


@app.route("/api/customer_logout", methods=["POST"])
def customer_logout():
    """客戶登出：只清客戶登入狀態，不動後台 session。"""
    session.pop("cust_g_code", None)
    return jsonify({"success": True})


@app.route("/api/forecast", methods=["POST"])
def create_forecast():
    data = request.json
    customer_id = data.get("customer_id")
    g_code = data.get("g_code", "")
    packages = data.get("packages", [])

    if not customer_id:
        return jsonify({"success": False, "error": "缺少客戶編號"})
    if not packages:
        return jsonify({"success": False, "error": "請至少填寫一個包裹"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp

    results = []
    for idx, pkg in enumerate(packages):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        local_logis_num = f"{g_code}-{timestamp}-{idx+1}"
        declare_list = []
        for item in pkg.get("items", []):
            declare_list.append({
                "product_name": item.get("name", "商品"),
                "product_name_local": item.get("name", "商品"),
                "product_num": int(item.get("quantity", 1)),
                "product_price": int(float(item.get("price", 0))),
                "product_url": item.get("url", "")
            })
        total_num = sum(int(item.get("quantity", 1)) for item in pkg.get("items", []))
        total_price = sum(int(float(item.get("price", 0))) * int(item.get("quantity", 1)) for item in pkg.get("items", []))
        forecast_data = {
            "packages": [{
                "local_logis_num": local_logis_num,
                "client_cid": g_code,
                "client_pid": pkg.get("client_pid") or local_logis_num,
                "client_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "warehouse_id": JPD_WAREHOUSE_ID,
                "product_name": declare_list[0]["product_name"] if declare_list else "商品",
                "product_num": total_num,
                "product_price": total_price,
                "declare_list": declare_list
            }]
        }
        result = jpd_request("TForecastPackage", forecast_data)
        if "OperationResult" in result:
            op_result = result["OperationResult"]
            if op_result["Request"]["IsValid"] == "True":
                result_data = op_result.get("Result", {})
                if result_data.get("Result") == "SUCCESS":
                    pkg_data = result_data.get("Data", [{}])[0]
                    results.append({
                        "success": True,
                        "local_logis_num": local_logis_num,
                        "package_id": pkg_data.get("package_id"),
                        "message": pkg_data.get("msg", "預報成功")
                    })
                    continue
        results.append({"success": False, "local_logis_num": local_logis_num, "error": "預報失敗"})

    return jsonify({"success": all(r["success"] for r in results), "results": results})


@app.route("/api/packages", methods=["GET"])
def get_packages():
    g_code = request.args.get("g_code") or request.args.get("customer_id")
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})

    g_code = g_code.upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM packages WHERE g_code=? ORDER BY id DESC",
        (g_code,)
    ).fetchall()

    # 找出該會員「進行中」的出貨申請（待處理／處理中），標記其包含的包裹
    # 用途：前端據此隱藏勾選框、顯示「已申請出貨」徽章，避免重複申請
    active_reqs = conn.execute(
        "SELECT id, package_ids FROM shipment_requests "
        "WHERE g_code=? AND status IN ('待處理', '處理中')",
        (g_code,)
    ).fetchall()
    conn.close()

    pkg_to_req = {}
    for r in active_reqs:
        ids_str = r["package_ids"] or ""
        for pid_str in ids_str.split(","):
            try:
                pkg_to_req[int(pid_str.strip())] = r["id"]
            except (ValueError, AttributeError):
                pass

    packages = []
    for row in rows:
        r = dict(row)
        packages.append({
            "id":           r["id"],
            "logis_num":    r["logis_num"] or "-",
            "product_name": r["product_name"] or "-",
            "weight":       r["weight"] or "",
            "status":       r["status"],
            "note":         r["note"] or "",
            "in_date":      r["in_date"] or "",
            "created_at":   r["created_at"],
            "pending_ship_request_id": pkg_to_req.get(r["id"]),  # None 表示未在出貨申請中
        })
    return jsonify({"success": True, "packages": packages})


@app.route("/api/orders", methods=["GET"])
def get_orders():
    """（已停用）原 JPD 運單查詢。已不與 JPD 合作，不再對其發送 API 請求。
    客戶端的運單查詢分頁改為只顯示「我的出貨申請」（含台灣配送貨況）。"""
    ok, resp = _require_customer(request.args.get("g_code") or request.args.get("customer_id"))
    if not ok:
        return resp
    return jsonify({"success": True, "orders": []})


def compute_agent_weekly(agent_id, conn=None):
    """算某代理各週分潤（與統計頁同一套算法，保證數字一致）。
    只算：已出貨 + 有金額 + 已收款（有匯款後五碼）。
    分潤 = Σ per_kg × kg，per_kg = max(該單費率 − 180, 20)。
    回傳 [{period_key, period_label, shipments, total_kg, commission, payout:{...}}, ...] 新→舊
    """
    own = False
    if conn is None:
        conn = get_db(); own = True
    rows = conn.execute("""
        SELECT * FROM shipment_requests
        WHERE status='已出貨' AND total_fee > 0 AND agent_id=?
          AND payment_last5 IS NOT NULL AND payment_last5 != ''
    """, (agent_id,)).fetchall()
    payouts = {
        p["period_key"]: dict(p)
        for p in conn.execute("SELECT * FROM agent_payouts WHERE agent_id=?", (agent_id,)).fetchall()
    }
    if own:
        conn.close()

    buckets = {}
    for row in rows:
        r = dict(row)
        date_str = r.get("updated_at") or r.get("created_at") or ""
        key, label = _period_key_from_date(date_str, "week")
        if not key:
            continue
        b = buckets.setdefault(key, {
            "period_key": key, "period_label": label,
            "shipments": 0, "total_kg": 0.0, "commission": 0.0,
        })
        kg = float(r.get("billed_weight") or 0)
        b["shipments"] += 1
        b["total_kg"] += kg
        if kg > 0:
            rate = float(r.get("rate_per_kg") or 0)
            b["commission"] += max(rate - 180, 20) * kg

    result = []
    for key in sorted(buckets.keys(), reverse=True):
        b = buckets[key]
        b["total_kg"] = round(b["total_kg"], 1)
        b["commission"] = round(b["commission"])
        p = payouts.get(key)
        b["paid"] = bool(p and p.get("paid_at"))
        b["payment_last5"] = (p or {}).get("payment_last5", "")
        b["paid_at"] = (p or {}).get("paid_at", "")
        b["payout_note"] = (p or {}).get("note", "")
        result.append(b)
    return result

@app.route("/api/admin/stats/monthly/detail", methods=["GET"])
def admin_monthly_detail():
    """取得指定週/月的出貨明細（month 參數值可為 '2026-06' 或 '2026-W23'）"""
    # ⚠️ 同 admin_monthly_stats：原本只有 is_staff() 一層，匿名可讀逐單明細
    #    （客編、姓名、重量、金額）。補上登入檢查。
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    if is_staff():
        return jsonify({"success": False, "error": "權限不足"}), 403
    month = request.args.get("month", "")
    if not month:
        return jsonify({"success": False, "error": "缺少期間"})
    aid = get_current_agent_id()
    try:
        conn = get_db()
        ex_sql, ex_params = _stats_exclude_sql()   # 統計排除測試帳號（清單空 → 無影響）
        if aid > 0:
            rows = conn.execute(f"""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0 AND agent_id=?{ex_sql}
                ORDER BY updated_at ASC
            """, [aid] + ex_params).fetchall()
        else:
            rows = conn.execute(f"""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0{ex_sql}
                ORDER BY updated_at ASC
            """, ex_params).fetchall()
        conn.close()
        details = []
        for r in rows:
            rd = dict(r)
            date_str = rd.get("updated_at") or rd.get("created_at") or ""
            if not _matches_period(date_str, month):
                continue
            extras = []
            try:
                extras = json.loads(rd.get("extra_services") or "[]")
            except:
                pass
            details.append({
                "date": date_str[:10],
                "g_code": rd.get("g_code", ""),
                "customer_name": rd.get("customer_name", ""),
                "ship_recipient": rd.get("ship_recipient", ""),
                "ship_phone": rd.get("ship_phone", ""),
                "ship_address": rd.get("ship_address", ""),
                "billed_weight": float(rd.get("billed_weight") or 0),
                "rate_per_kg": float(rd.get("rate_per_kg") or 0),
                "shipping_fee": float(rd.get("shipping_fee") or 0),
                "handling_fee": float(rd.get("handling_fee") or 0),
                "consolidation_fee": float(rd.get("consolidation_fee") or 0),
                "letter_fee": float(rd.get("letter_fee") or 0),
                "extra_services": extras,
                "total_fee": float(rd.get("total_fee") or 0),
            })
        # 代理檢視時計算分潤（依合約第六條第 4 項公式）
        # 公式：(該客戶運費 − NT$180/kg) × 包裹重量，最低 NT$20/kg × 包裹重量
        commission = None
        if aid > 0:
            base_cost = 180  # NT$/kg 甲方批發成本
            min_per_kg = 20  # NT$/kg 最低分潤
            total_commission = 0
            total_kg = 0
            for d in details:
                kg = d["billed_weight"]
                per_kg = max(d["rate_per_kg"] - base_cost, min_per_kg)
                d["commission"] = round(per_kg * kg)
                total_commission += d["commission"]
                total_kg += kg
            commission = {
                "total": round(total_commission),
                "total_kg": round(total_kg, 1),
                "min_per_kg": min_per_kg,
                "base_cost_per_kg": base_cost,
            }
        return jsonify({"success": True, "details": details, "is_agent": aid > 0, "commission": commission})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/stats/monthly/excel", methods=["GET"])
def admin_monthly_excel():
    """下載指定月份的出貨明細 Excel"""
    # ⚠️ 同 admin_monthly_stats：原本只有 is_staff() 一層，匿名可下載整月明細 Excel
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    if is_staff():
        return jsonify({"success": False, "error": "權限不足"}), 403
    month = request.args.get("month", "")  # e.g. "2026-04"
    if not month:
        return jsonify({"success": False, "error": "缺少月份參數"})
    aid = get_current_agent_id()
    try:
        conn = get_db()
        ex_sql, ex_params = _stats_exclude_sql()   # 統計排除測試帳號（清單空 → 無影響）
        if aid > 0:
            rows = conn.execute(f"""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0 AND agent_id=?{ex_sql}
                ORDER BY updated_at ASC
            """, [aid] + ex_params).fetchall()
        else:
            rows = conn.execute(f"""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0{ex_sql}
                ORDER BY updated_at ASC
            """, ex_params).fetchall()
        conn.close()

        # 篩選指定週/月
        filtered = []
        for r in rows:
            rd = dict(r)
            date_str = rd.get("updated_at") or rd.get("created_at") or ""
            if _matches_period(date_str, month):
                filtered.append(rd)

        wb = Workbook()
        ws = wb.active
        ws.title = f"{month} 出貨明細"

        # 標題樣式
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )

        # 合箱費原本整支函式都沒讀，各欄加總會比合計少掉一筆合箱費。
        # 這份檔案會下載出去跟客戶/代理對帳，欄位對不上合計是對方先發現。
        # 費用欄序定案：運費 → 理貨 → 合箱 → 加值明細 → 加值小計 → 信件 → 合計
        # 與帳單管理表、統計月表、月營收明細一致，跨報表不用重新對位。
        headers = ["出貨日期", "客戶編號", "客戶姓名", "寄送地址", "計費重量(kg)",
                    "運費單價", "運費小計", "理貨費", "合箱費", "加值服務明細", "加值服務小計",
                    "信件費", "合計(台幣)"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        total_kg = 0
        total_shipping = 0
        total_handling = 0
        total_extra = 0
        total_letter = 0
        total_consolidation = 0
        total_all = 0

        for i, r in enumerate(filtered, 2):
            date_str = r["updated_at"] or r["created_at"] or ""
            bw = float(r["billed_weight"] or 0)
            rate = float(r["rate_per_kg"] or 0)
            sf = float(r["shipping_fee"] or 0)
            hf = float(r["handling_fee"] or 0)
            lf = float(r["letter_fee"] or 0)
            cf = float(r["consolidation_fee"] or 0)
            tf = float(r["total_fee"] or 0)

            # 加值服務
            extra_desc = ""
            extra_total = 0
            try:
                extras = json.loads(r["extra_services"] or "[]")
                parts = []
                for e in extras:
                    qty = int(e.get("qty", 1))
                    price = int(e.get("price", 0))
                    sub = int(e.get("subtotal", qty * price))
                    parts.append(f"{e.get('name','')} ×{qty} = NT${sub}")
                    extra_total += sub
                extra_desc = " / ".join(parts)
            except:
                pass

            ship_addr = " ".join(filter(None, [str(r.get("ship_recipient") or ""), str(r.get("ship_phone") or ""), str(r.get("ship_address") or "")]))

            ws.cell(row=i, column=1, value=date_str[:10]).border = thin_border
            ws.cell(row=i, column=2, value=r["g_code"]).border = thin_border
            ws.cell(row=i, column=3, value=str(r.get("customer_name") or "")).border = thin_border
            ws.cell(row=i, column=4, value=ship_addr).border = thin_border
            ws.cell(row=i, column=5, value=bw).border = thin_border
            ws.cell(row=i, column=6, value=rate).border = thin_border
            ws.cell(row=i, column=7, value=sf).border = thin_border
            ws.cell(row=i, column=8, value=hf).border = thin_border
            ws.cell(row=i, column=9, value=cf).border = thin_border
            ws.cell(row=i, column=10, value=extra_desc).border = thin_border
            ws.cell(row=i, column=11, value=extra_total).border = thin_border
            ws.cell(row=i, column=12, value=lf).border = thin_border
            ws.cell(row=i, column=13, value=tf).border = thin_border

            total_kg += bw
            total_shipping += sf
            total_handling += hf
            total_extra += extra_total
            total_letter += lf
            total_consolidation += cf
            total_all += tf

        # 合計列
        sum_row = len(filtered) + 2
        sum_font = Font(bold=True, size=11)
        sum_fill = PatternFill(start_color="F39C12", end_color="F39C12", fill_type="solid")
        ws.cell(row=sum_row, column=1, value="合計").font = sum_font
        ws.cell(row=sum_row, column=1).fill = sum_fill
        ws.cell(row=sum_row, column=1).border = thin_border
        for c in range(2, 14):
            ws.cell(row=sum_row, column=c).border = thin_border
            ws.cell(row=sum_row, column=c).font = sum_font
        ws.cell(row=sum_row, column=2, value=f"{len(filtered)} 筆")
        ws.cell(row=sum_row, column=5, value=total_kg)
        ws.cell(row=sum_row, column=7, value=total_shipping)
        ws.cell(row=sum_row, column=8, value=total_handling)
        ws.cell(row=sum_row, column=9, value=total_consolidation)
        ws.cell(row=sum_row, column=11, value=total_extra)
        ws.cell(row=sum_row, column=12, value=total_letter)
        ws.cell(row=sum_row, column=13, value=total_all)

        # 欄寬
        # I=合箱費、J=加值服務明細（寬）、K=加值小計、L=信件費、M=合計(台幣)
        # M 原本就漏設（舊 widths 只到 K），一併補上
        widths = {'A':12, 'B':10, 'C':12, 'D':30, 'E':12, 'F':10, 'G':12, 'H':10, 'I':12, 'J':30,
                  'K':12, 'L':12, 'M':14}
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        filename = f"{BRAND['slug']}_{month}_出貨明細.xlsx"
        return send_file(buf, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


def _period_key_from_date(date_str, period_type):
    """
    把 'YYYY-MM-DD HH:MM:SS' 轉成 (period_key, period_label)
    - month → ('2026-06', '2026-06')
    - week  → ('2026-W23', '6/1 - 6/7')  (ISO 週、週一開始)
    """
    if period_type == "week":
        from datetime import datetime as _dt, timedelta as _td
        try:
            d = _dt.strptime(date_str[:10], "%Y-%m-%d")
        except Exception:
            return None, None
        iso_year, iso_week, iso_weekday = d.isocalendar()
        monday = d - _td(days=iso_weekday - 1)
        sunday = monday + _td(days=6)
        key = f"{iso_year}-W{iso_week:02d}"
        if monday.year == sunday.year and monday.month == sunday.month:
            label = f"{monday.month}/{monday.day} - {sunday.day}"
        elif monday.year == sunday.year:
            label = f"{monday.month}/{monday.day} - {sunday.month}/{sunday.day}"
        else:
            label = f"{monday.year}/{monday.month}/{monday.day} - {sunday.year}/{sunday.month}/{sunday.day}"
        return key, label
    else:
        # month
        if len(date_str) < 7:
            return None, None
        key = date_str[:7]
        return key, key


def _matches_period(date_str, period_key):
    """date_str 是否屬於 period_key（自動偵測週/月）"""
    if not date_str or not period_key:
        return False
    if "W" in period_key:
        k, _ = _period_key_from_date(date_str, "week")
        return k == period_key
    else:
        return date_str[:7] == period_key


@app.route("/api/admin/stats/customer-activity", methods=["GET"])
def admin_customer_activity():
    """客戶活躍度（老闆專用）：依最後活動（已出貨日 或 到貨日 取較近）分類。
    活躍≤7天／觀察7-30／沉睡30-60／僵屍>60／未啟用(從沒動過)。"""
    if not is_boss():
        return jsonify({"success": False, "error": "權限不足"}), 403
    seg = (request.args.get("seg") or "").strip()       # active/watch/sleep/zombie/inactive/''(全部)
    q = (request.args.get("q") or "").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except (ValueError, TypeError):
        page = 1
    try:
        limit = min(200, max(1, int(request.args.get("limit", 50))))
    except (ValueError, TypeError):
        limit = 50

    conn = get_db()
    ex_sql, ex_params = _stats_exclude_sql()   # 統計排除測試帳號（清單空 → 無影響）
    excluded = set(_stats_excluded_codes())
    # 每個客編最後已出貨日
    last_ship = {}
    for r in conn.execute(
        "SELECT g_code, MAX(substr(COALESCE(NULLIF(updated_at,''),created_at),1,10)) AS d "
        "FROM shipment_requests WHERE status='已出貨' AND g_code IS NOT NULL AND g_code!=''"
        + ex_sql + " GROUP BY g_code", ex_params
    ).fetchall():
        if r["d"]:
            last_ship[r["g_code"].upper()] = r["d"]
    # 每個客編最後到貨日
    last_arr = {}
    for r in conn.execute(
        "SELECT g_code, MAX(substr(in_date,1,10)) AS d FROM packages "
        "WHERE g_code IS NOT NULL AND g_code!='' AND in_date IS NOT NULL AND in_date!=''"
        + ex_sql + " GROUP BY g_code", ex_params
    ).fetchall():
        if r["d"]:
            last_arr[r["g_code"].upper()] = r["d"]
    conn.close()

    # 客戶名單（Shopify 快取；用於抓「未啟用＝註冊但從沒動」）
    customers = get_all_goyoutati_customers() or []

    today = datetime.now().date()

    def classify(days):
        if days is None:
            return "inactive"
        if days <= 7:
            return "active"
        if days <= 30:
            return "watch"
        if days <= 60:
            return "sleep"
        return "zombie"

    rows = []
    counts = {"active": 0, "watch": 0, "sleep": 0, "zombie": 0, "inactive": 0}
    for c in customers:
        gc = (c.get("g_code") or "").strip().upper()
        if not gc or gc in excluded:      # 測試帳號不進客戶清單、不計入各分類數
            continue
        s = last_ship.get(gc); a = last_arr.get(gc)
        last = max([x for x in (s, a) if x], default=None)
        days = None
        if last:
            try:
                days = (today - datetime.strptime(last, "%Y-%m-%d").date()).days
            except (ValueError, TypeError):
                days = None
        cat = classify(days)
        counts[cat] += 1
        rows.append({
            "g_code": gc, "name": c.get("name", ""),
            "last_ship": s or "", "last_arrival": a or "",
            "last_active": last or "", "days_inactive": days,
            "segment": cat,
        })

    # 篩選
    filtered = rows
    if seg in counts:
        filtered = [r for r in filtered if r["segment"] == seg]
    if q:
        ql = q.upper()
        filtered = [r for r in filtered if ql in r["g_code"] or ql in (r["name"] or "").upper()]
    # 排序：最久沒動排前面（未啟用視為最久；未啟用 days=None 放最後另計）
    filtered.sort(key=lambda r: (r["days_inactive"] is None, -(r["days_inactive"] or 0)))
    # 但未啟用要排最前面時另處理：預設把「已動過的」依天數多→少，未啟用放最後
    total = len(filtered)
    start = (page - 1) * limit
    page_rows = filtered[start:start + limit]
    return jsonify({"success": True, "rows": page_rows, "counts": counts,
                    "total": total, "page": page, "limit": limit,
                    "has_more": start + len(page_rows) < total,
                    "total_customers": len(rows)})


@app.route("/api/admin/stats/daily", methods=["GET"])
def admin_daily_ops():
    """每日營運（老闆＋員工）：每天的進倉件數/重量、出貨單數/重量/金額。
    進倉＝packages 依到倉日 in_date；出貨＝shipment_requests 依出貨日(updated_at)、
    算『當天標記已出貨的全部單』(看工作量/產出，不論是否收款)。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    date_from = (request.args.get("date_from") or "").strip()
    date_to = (request.args.get("date_to") or "").strip()
    # 預設近 7 天
    if not date_from and not date_to:
        date_to = datetime.now().strftime("%Y-%m-%d")
        date_from = (datetime.now() - timedelta(days=6)).strftime("%Y-%m-%d")
    aid = get_current_agent_id()
    conn = get_db()

    # 進倉：依到倉日 group（in_date 為 YYYY-MM-DD；防呆截前10碼）
    ex_sql, ex_params = _stats_exclude_sql()   # 統計排除測試帳號（清單空 → 無影響）
    inb_where = ["substr(COALESCE(in_date,''),1,10) BETWEEN ? AND ?"]
    inb_params = [date_from, date_to]
    if aid > 0:
        inb_where.append("agent_id=?"); inb_params.append(aid)
    if ex_sql:
        inb_where.append(ex_sql[len(" AND "):]); inb_params += ex_params
    inbound = {}
    for r in conn.execute(
        f"SELECT substr(in_date,1,10) AS d, COUNT(*) AS cnt, "
        f"COALESCE(SUM(CAST(weight AS REAL)),0) AS kg FROM packages "
        f"WHERE {' AND '.join(inb_where)} GROUP BY d", inb_params
    ).fetchall():
        inbound[r["d"]] = {"in_count": r["cnt"], "in_kg": round(r["kg"] or 0, 1)}

    # 出貨：依出貨日 group（出貨日＝updated_at 那刻；當天全部已出貨單）
    out_where = ["status='已出貨'",
                 "substr(COALESCE(NULLIF(updated_at,''), created_at),1,10) BETWEEN ? AND ?"]
    out_params = [date_from, date_to]
    if aid > 0:
        out_where.append("agent_id=?"); out_params.append(aid)
    if ex_sql:
        out_where.append(ex_sql[len(" AND "):]); out_params += ex_params
    outbound = {}
    for r in conn.execute(
        f"SELECT substr(COALESCE(NULLIF(updated_at,''), created_at),1,10) AS d, "
        f"COUNT(*) AS cnt, COALESCE(SUM(CAST(billed_weight AS REAL)),0) AS kg, "
        f"COALESCE(SUM(CAST(total_fee AS REAL)),0) AS fee FROM shipment_requests "
        f"WHERE {' AND '.join(out_where)} GROUP BY d", out_params
    ).fetchall():
        outbound[r["d"]] = {"out_count": r["cnt"], "out_kg": round(r["kg"] or 0, 1),
                            "out_fee": round(r["fee"] or 0)}
    conn.close()

    # 合併每一天（區間內每天一列，新到舊）
    days = sorted(set(list(inbound.keys()) + list(outbound.keys())), reverse=True)
    rows = []
    tot = {"in_count": 0, "in_kg": 0.0, "out_count": 0, "out_kg": 0.0, "out_fee": 0}
    for d in days:
        i = inbound.get(d, {}); o = outbound.get(d, {})
        row = {
            "date": d,
            "in_count": i.get("in_count", 0), "in_kg": i.get("in_kg", 0),
            "out_count": o.get("out_count", 0), "out_kg": o.get("out_kg", 0),
            "out_fee": o.get("out_fee", 0),
        }
        rows.append(row)
        tot["in_count"] += row["in_count"]; tot["in_kg"] += row["in_kg"]
        tot["out_count"] += row["out_count"]; tot["out_kg"] += row["out_kg"]; tot["out_fee"] += row["out_fee"]
    tot["in_kg"] = round(tot["in_kg"], 1); tot["out_kg"] = round(tot["out_kg"], 1)
    return jsonify({"success": True, "rows": rows, "total": tot,
                    "date_from": date_from, "date_to": date_to})


@app.route("/api/admin/stats/monthly", methods=["GET"])
def admin_monthly_stats():
    """月/週統計：代理→按週、主帳號→可選月/週（?period=month|week，預設 month）"""
    # ⚠️ 原本只有 is_staff() 一層，匿名請求會直接穿過去（未登入的 user_type 不是
    #    'admin'，is_staff() 回 False）→ 全站營收月報可未登入讀取。補上登入檢查。
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    if is_staff():
        return jsonify({"success": False, "error": "權限不足"}), 403
    aid = get_current_agent_id()
    if aid > 0:
        period_type = "week"
    else:
        period_type = request.args.get("period", "month")
        if period_type not in ("month", "week"):
            period_type = "month"
    try:
        conn = get_db()
        ex_sql, ex_params = _stats_exclude_sql()   # 統計排除測試帳號（清單空 → 無影響）
        if aid > 0:
            rows = conn.execute(f"""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0 AND agent_id=?
                  AND payment_last5 IS NOT NULL AND payment_last5 != ''{ex_sql}
                ORDER BY updated_at DESC
            """, [aid] + ex_params).fetchall()
        else:
            rows = conn.execute(f"""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0
                  AND payment_last5 IS NOT NULL AND payment_last5 != ''{ex_sql}
                ORDER BY updated_at DESC
            """, ex_params).fetchall()
        conn.close()

        buckets = {}
        for row in rows:
            r = dict(row)
            date_str = r.get("updated_at") or r.get("created_at") or ""
            if not date_str:
                continue
            key, label = _period_key_from_date(date_str, period_type)
            if not key:
                continue

            if key not in buckets:
                buckets[key] = {
                    "month": key,            # 保留欄位名稱以維持向後相容（前端、明細 API、Excel 都用此）
                    "period_label": label,
                    "shipments": 0,
                    "total_kg": 0,
                    "shipping_fee": 0,
                    "handling_fee": 0,
                    "consolidation_fee": 0,
                    "letter_fee": 0,
                    "extra_fee": 0,
                    "total_revenue": 0,
                    "commission": 0,        # 代理分潤累計（主帳號為 0）
                    "customers": set()
                }

            m = buckets[key]
            m["shipments"] += 1
            kg = float(r["billed_weight"] or 0)
            m["total_kg"] += kg
            m["shipping_fee"] += float(r["shipping_fee"] or 0)
            m["handling_fee"] += float(r["handling_fee"] or 0)
            m["consolidation_fee"] += float(r.get("consolidation_fee") or 0)
            m["letter_fee"] += float(r.get("letter_fee") or 0)
            m["total_revenue"] += float(r["total_fee"] or 0)
            m["customers"].add(r["g_code"])

            # 代理分潤累加：(rate - 180) × kg，最低 20 × kg
            if aid > 0 and kg > 0:
                rate = float(r["rate_per_kg"] or 0)
                per_kg = max(rate - 180, 20)
                m["commission"] += per_kg * kg

            try:
                extras = json.loads(r["extra_services"] or "[]")
                for e in extras:
                    m["extra_fee"] += int(e.get("subtotal") or e.get("qty", 1) * e.get("price", 0) or 0)
            except:
                pass

        result = []
        for key in sorted(buckets.keys(), reverse=True):
            m = buckets[key]
            m["customer_count"] = len(m["customers"])
            m["commission"] = round(m["commission"])
            del m["customers"]
            result.append(m)

        # 代理端：附上每週撥款狀態（讓代理在統計頁看得到「已撥款/後五碼」）
        if aid > 0 and result:
            conn2 = get_db()
            payouts = {
                p["period_key"]: dict(p)
                for p in conn2.execute("SELECT * FROM agent_payouts WHERE agent_id=?", (aid,)).fetchall()
            }
            conn2.close()
            for m in result:
                p = payouts.get(m["month"])
                m["paid"] = bool(p and p.get("paid_at"))
                m["payment_last5"] = (p or {}).get("payment_last5", "")
                m["paid_at"] = (p or {}).get("paid_at", "")

        return jsonify({
            "success": True,
            "period_type": period_type,
            "is_agent": aid > 0,
            "monthly": result
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# ============ JPD 自動建單 ============

@app.route("/api/admin/shipment_requests/<int:req_id>/jpd_create", methods=["POST"])
def admin_create_jpd_order(req_id):
    """從出貨申請自動在 JPD 建立運單"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到出貨申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    req = conn.execute("SELECT * FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
    req = dict(req)
    g_code = req.get("g_code", "")
    
    # 收件人資訊
    recipient = str(req.get("ship_recipient") or "")
    phone = str(req.get("ship_phone") or "")
    address = str(req.get("ship_address") or "")
    note = str(req.get("note") or "")
    
    if not recipient or not phone or not address:
        conn.close()
        return jsonify({"success": False, "error": "缺少寄送地址資訊"})
    
    # 從預報取得申報品項
    forecasts = conn.execute(
        "SELECT * FROM forecasts WHERE g_code=? AND status='待處理' ORDER BY id", (g_code,)
    ).fetchall()
    
    declare_list = []
    for fc in forecasts:
        fc = dict(fc)
        try:
            items = json.loads(fc.get("items_json") or "[]")
            for item in items:
                declare_list.append({
                    "product_name": item.get("name", "雑貨"),
                    "product_name_local": item.get("name", "雑貨"),
                    "product_num": int(item.get("quantity", 1)),
                    "product_price": int(float(item.get("price", 0))),
                    "product_url": item.get("url", "")
                })
        except Exception as e:
            print(f"[JPD] declare_list 解析失敗: {e}", flush=True)

    # 如果沒有預報品項，給一個預設品項
    if not declare_list:
        declare_list = [{
            "product_name": "雑貨",
            "product_name_local": "雑貨",
            "product_num": 1,
            "product_price": 0,
            "product_url": ""
        }]
    
    # 入庫包裹 ID 由管理員手動輸入（多個用逗號/頓號/空白分隔）
    body = request.get_json(silent=True) or {}
    raw_ids = body.get("package_ids", "")
    if isinstance(raw_ids, list):
        id_tokens = [str(x).strip() for x in raw_ids]
    else:
        id_tokens = re.split(r"[,，、\s]+", str(raw_ids))
    jpd_package_ids = []
    for tok in id_tokens:
        tok = tok.strip()
        if not tok:
            continue
        try:
            jpd_package_ids.append(int(tok))
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": f"包裹 ID「{tok}」格式錯誤，應為數字"})

    if not jpd_package_ids:
        conn.close()
        return jsonify({"success": False, "error": "請填入 JPD 入庫包裹 ID（可在 JPD 雲倉的入庫列表查到，多個用逗號分隔）"})

    # 客戶運單號前綴：客編 + 日期
    today_str = datetime.now().strftime("%m%d")
    order_prefix = f"{g_code}-{today_str}"

    # 方案 A：一個包裹建一張運單，運單號一律加流水號（-1, -2, -3...）
    created_orders = []
    failed_orders = []

    for idx, pid in enumerate(jpd_package_ids):
        customer_order_id = f"{order_prefix}-{idx + 1}"
        order_data = {
            "customer_order_id": customer_order_id,
            "deliv_id": JPD_DELIV_ID,
            "recipient": recipient,
            "id_issure": "",
            "area": 3,
            "addr1": address,
            "addr2": "",
            "addr3": "",
            "addr4": "",
            "tel": phone,
            "memo": note,
            "create_order_pdf": "n",
            "warehouse_id": JPD_WAREHOUSE_ID,
            "create_package": "n",
            "create_sender": "y",
            "packages": [{"package_id": pid, "declare_list": declare_list}]
        }
        print(f"[JPD] 建立運單: {customer_order_id}, 收件人: {recipient}, 包裹 id={pid}", flush=True)
        result = jpd_request("TCreateOrder", order_data)

        ok = False
        if "OperationResult" in result:
            op = result["OperationResult"]
            if op.get("Request", {}).get("IsValid") == "True":
                res_data = op.get("Result", {})
                if res_data.get("Result") == "SUCCESS":
                    data = res_data.get("Data", {})
                    jpd_order_id = str(data.get("order_id", ""))
                    jpd_logis_num = str(data.get("logis_num", ""))
                    print(f"[JPD] ✅ 建單成功: {customer_order_id} order_id={jpd_order_id}, logis_num={jpd_logis_num}", flush=True)
                    created_orders.append({
                        "customer_order_id": customer_order_id,
                        "jpd_order_id": jpd_order_id,
                        "jpd_logis_num": jpd_logis_num,
                        "package_id": pid
                    })
                    ok = True
                else:
                    err = res_data.get("ErrorMsg", str(res_data))
                    print(f"[JPD] ❌ 建單失敗 {customer_order_id}: {err}", flush=True)
                    failed_orders.append({"customer_order_id": customer_order_id, "error": str(err)})
            else:
                errs = op.get("Request", {}).get("Errors", {})
                print(f"[JPD] ❌ 請求無效 {customer_order_id}: {errs}", flush=True)
                failed_orders.append({"customer_order_id": customer_order_id, "error": str(errs)})
        if not ok and not failed_orders:
            failed_orders.append({"customer_order_id": customer_order_id, "error": "JPD API 無回應"})

    conn.close()

    if created_orders and not failed_orders:
        nums = ', '.join([o["jpd_logis_num"] or o["customer_order_id"] for o in created_orders])
        return jsonify({
            "success": True,
            "created": created_orders,
            "message": f"已建立 {len(created_orders)} 張 JPD 運單：{nums}"
        })
    elif created_orders and failed_orders:
        return jsonify({
            "success": True,
            "created": created_orders,
            "failed": failed_orders,
            "message": f"成功 {len(created_orders)} 張、失敗 {len(failed_orders)} 張。失敗：" +
                       '、'.join([f["customer_order_id"] + '(' + f["error"] + ')' for f in failed_orders])
        })
    else:
        first_err = failed_orders[0]["error"] if failed_orders else "未知錯誤"
        return jsonify({"success": False, "error": f"JPD 建單全部失敗：{first_err}"})


# ============ 公告 API ============

@app.route("/api/announcements", methods=["GET"])
def get_announcements():
    """取得啟用中的公告"""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM announcements WHERE is_active=1 ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "announcements": [dict(r) for r in rows]})


@app.route("/api/admin/announcements", methods=["GET"])
def admin_get_announcements():
    """管理員取得所有公告（含未啟用的）。需登入。

    客戶端的公告牆走的是另一支 /api/announcements（只回 is_active=1），
    不受這裡影響 —— 客戶登入後照常看得到公告。
    這支多回了未啟用/已下架的公告內容與 agent_id，只給後台看。
    守門層級與同資源的 POST/PUT/DELETE 一致（is_super_admin），代理看不到。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    conn = get_db()
    rows = conn.execute("SELECT * FROM announcements ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify({"success": True, "announcements": [dict(r) for r in rows]})


@app.route("/api/admin/announcements", methods=["POST"])
def admin_create_announcement():
    """管理員新增公告（僅主管理員）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    title = (data.get("title") or "").strip()
    content = (data.get("content") or "").strip()
    if not title or not content:
        return jsonify({"success": False, "error": "標題和內容為必填"})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    cur = conn.execute(
        "INSERT INTO announcements (title, content, is_active, created_at) VALUES (?, ?, 1, ?)",
        (title, content, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": cur.lastrowid})


@app.route("/api/admin/announcements/<int:ann_id>", methods=["PUT"])
def admin_update_announcement(ann_id):
    """管理員更新公告（僅主管理員）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    conn = get_db()
    fields = {}
    for key in ["title", "content"]:
        if key in data:
            fields[key] = (data[key] or "").strip()
    if "is_active" in data:
        fields["is_active"] = 1 if data["is_active"] else 0
    if fields:
        sets = ", ".join(f"{k}=?" for k in fields)
        vals = list(fields.values()) + [ann_id]
        conn.execute(f"UPDATE announcements SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/announcements/<int:ann_id>", methods=["DELETE"])
def admin_delete_announcement(ann_id):
    """管理員刪除公告（僅主管理員）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    conn.execute("DELETE FROM announcements WHERE id=?", (ann_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 內部公告 API（後台老闆/員工專用，客戶端無此路由）============

@app.route("/api/admin/internal_announcements", methods=["GET"])
def admin_get_internal_announcements():
    """後台列出所有內部公告（老闆＋員工皆可看）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    conn = get_db()
    rows = conn.execute("SELECT * FROM internal_announcements ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify({"success": True, "announcements": [dict(r) for r in rows]})


@app.route("/api/admin/internal_announcements/latest", methods=["GET"])
def admin_latest_internal_announcement():
    """登入橫幅用：回最新一則啟用中的內部公告，並標記當前使用者是否已讀。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM internal_announcements WHERE is_active=1 ORDER BY id DESC LIMIT 1"
    ).fetchone()
    if not row:
        conn.close()
        return jsonify({"success": True, "announcement": None})
    me = current_operator()
    read = conn.execute(
        "SELECT 1 FROM internal_ann_reads WHERE username=? AND ann_id=?", (me, row["id"])
    ).fetchone()
    conn.close()
    d = dict(row)
    d["read"] = bool(read)
    return jsonify({"success": True, "announcement": d})


@app.route("/api/admin/internal_announcements/<int:ann_id>/read", methods=["POST"])
def admin_read_internal_announcement(ann_id):
    """當前使用者按「我知道了」→ 記已讀（該則對他不再跳）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "請先登入"}), 403
    conn = get_db()
    conn.execute(
        "INSERT OR IGNORE INTO internal_ann_reads (username, ann_id, read_at) VALUES (?, ?, ?)",
        (current_operator(), ann_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/internal_announcements", methods=["POST"])
def admin_create_internal_announcement():
    """發布內部公告（僅老闆）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆能發布內部公告"}), 403
    data = request.json or {}
    title = (data.get("title") or "").strip()
    content = (data.get("content") or "").strip()
    if not title or not content:
        return jsonify({"success": False, "error": "標題和內容為必填"})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    cur = conn.execute(
        "INSERT INTO internal_announcements (title, content, is_active, created_at) VALUES (?, ?, 1, ?)",
        (title, content, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": cur.lastrowid})


@app.route("/api/admin/internal_announcements/<int:ann_id>", methods=["PUT"])
def admin_update_internal_announcement(ann_id):
    """停用/啟用/修改內部公告（僅老闆）。"""
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆能管理內部公告"}), 403
    data = request.json or {}
    conn = get_db()
    fields = {}
    for key in ["title", "content"]:
        if key in data:
            fields[key] = (data[key] or "").strip()
    if "is_active" in data:
        fields["is_active"] = 1 if data["is_active"] else 0
    if fields:
        sets = ", ".join(f"{k}=?" for k in fields)
        conn.execute(f"UPDATE internal_announcements SET {sets} WHERE id=?", list(fields.values()) + [ann_id])
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/internal_announcements/<int:ann_id>", methods=["DELETE"])
def admin_delete_internal_announcement(ann_id):
    if not is_boss():
        return jsonify({"success": False, "error": "只有老闆能刪除內部公告"}), 403
    conn = get_db()
    conn.execute("DELETE FROM internal_announcements WHERE id=?", (ann_id,))
    conn.execute("DELETE FROM internal_ann_reads WHERE ann_id=?", (ann_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 地址簿 API ============

@app.route("/api/addresses", methods=["GET"])
def get_addresses():
    """取得客戶地址簿"""
    g_code = request.args.get("g_code", "").upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM addresses WHERE g_code=? ORDER BY is_default DESC, id DESC", (g_code,)
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "addresses": [dict(r) for r in rows]})


@app.route("/api/addresses", methods=["POST"])
def add_address():
    """新增地址"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    recipient = (data.get("recipient") or "").strip()
    phone = (data.get("phone") or "").strip()
    address = (data.get("address") or "").strip()
    label = (data.get("label") or "").strip()
    zipcode = (data.get("zipcode") or "").strip()
    is_default = 1 if data.get("is_default") else 0

    if not g_code or not recipient or not phone or not address:
        return jsonify({"success": False, "error": "收件人、電話、地址為必填"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    # 如果設為預設，先清除其他預設
    if is_default:
        conn.execute("UPDATE addresses SET is_default=0 WHERE g_code=?", (g_code,))
    # 如果是第一筆，自動設為預設
    count = conn.execute("SELECT COUNT(*) as c FROM addresses WHERE g_code=?", (g_code,)).fetchone()["c"]
    if count == 0:
        is_default = 1

    conn.execute(
        """INSERT INTO addresses (g_code, label, recipient, phone, zipcode, address, is_default, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (g_code, label, recipient, phone, zipcode, address, is_default, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "地址已新增"})


@app.route("/api/addresses/<int:addr_id>", methods=["PUT"])
def update_address(addr_id):
    """更新地址"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    # 驗證是本人的
    row = conn.execute("SELECT g_code FROM addresses WHERE id=?", (addr_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該地址"})

    fields = {}
    for key in ["label", "recipient", "phone", "zipcode", "address"]:
        if key in data:
            fields[key] = (data[key] or "").strip()
    if "is_default" in data and data["is_default"]:
        conn.execute("UPDATE addresses SET is_default=0 WHERE g_code=?", (g_code,))
        fields["is_default"] = 1

    if fields:
        sets = ", ".join(f"{k}=?" for k in fields)
        vals = list(fields.values()) + [addr_id]
        conn.execute(f"UPDATE addresses SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/addresses/<int:addr_id>", methods=["DELETE"])
def delete_address(addr_id):
    """刪除地址"""
    data = request.json or {}
    g_code = (data.get("g_code") or request.args.get("g_code", "")).strip().upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    row = conn.execute("SELECT g_code, is_default FROM addresses WHERE id=?", (addr_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該地址"})
    conn.execute("DELETE FROM addresses WHERE id=?", (addr_id,))
    # 如果刪的是預設，把第一筆設為預設
    if row["is_default"]:
        first = conn.execute("SELECT id FROM addresses WHERE g_code=? ORDER BY id LIMIT 1", (g_code,)).fetchone()
        if first:
            conn.execute("UPDATE addresses SET is_default=1 WHERE id=?", (first["id"],))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/addresses/<int:addr_id>/default", methods=["POST"])
def set_default_address(addr_id):
    """設為預設地址"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    row = conn.execute("SELECT g_code FROM addresses WHERE id=?", (addr_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該地址"})
    conn.execute("UPDATE addresses SET is_default=0 WHERE g_code=?", (g_code,))
    conn.execute("UPDATE addresses SET is_default=1 WHERE id=?", (addr_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 申報人 API（報單收貨人／納稅義務人，會員層級 1..N）============

DECLARANT_MAX = 10
DECLARANT_PHONE_ERR = ("申報人電話必須是台灣手機門號（09 開頭 10 碼），"
                       "且須為 EZ WAY 實名認證綁定的號碼")


def _normalize_tw_mobile(raw):
    """台灣手機門號正規化：去空白與 -、+886/886 前綴 → 09xxxxxxxx。
    不合法（市話、位數不足、含英文…）回空字串，由呼叫端擋下。"""
    p = re.sub(r"[\s\-()]", "", str(raw or ""))
    if p.startswith("+886"):
        p = "0" + p[4:]
    elif p.startswith("886"):
        p = "0" + p[3:]
    return p if re.fullmatch(r"09\d{8}", p) else ""


@app.route("/api/declarants", methods=["GET"])
def get_declarants():
    """取得客戶的申報人清單"""
    g_code = request.args.get("g_code", "").strip().upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM declarants WHERE g_code=? ORDER BY is_default DESC, id DESC", (g_code,)
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "declarants": [dict(r) for r in rows]})


@app.route("/api/declarants", methods=["POST"])
def add_declarant():
    """新增申報人（第一筆自動成為預設）"""
    data = request.json or {}
    g_code = (data.get("g_code") or "").strip().upper()
    name = (data.get("name") or "").strip()
    address = (data.get("address") or "").strip()
    is_default = 1 if data.get("is_default") else 0

    if not g_code or not name or not address or not (data.get("phone") or "").strip():
        return jsonify({"success": False, "error": "申報人姓名、EZ WAY 綁定手機、地址為必填"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    phone = _normalize_tw_mobile(data.get("phone"))
    if not phone:
        return jsonify({"success": False, "error": DECLARANT_PHONE_ERR})

    conn = get_db()
    count = conn.execute("SELECT COUNT(*) AS c FROM declarants WHERE g_code=?", (g_code,)).fetchone()["c"]
    if count >= DECLARANT_MAX:
        conn.close()
        return jsonify({"success": False, "error": f"申報人最多只能新增 {DECLARANT_MAX} 位"})
    # 一個門號只能對應一個 EZ WAY 帳號 → 同一會員底下不得重複
    if conn.execute("SELECT 1 FROM declarants WHERE g_code=? AND phone=?", (g_code, phone)).fetchone():
        conn.close()
        return jsonify({"success": False, "error": f"門號 {phone} 已經在你的申報人清單中"})

    if is_default:
        conn.execute("UPDATE declarants SET is_default=0 WHERE g_code=?", (g_code,))
    if count == 0:
        is_default = 1

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """INSERT INTO declarants (g_code, name, phone, address, is_default, created_at)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (g_code, name, phone, address, is_default, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "申報人已新增"})


@app.route("/api/declarants/<int:dec_id>", methods=["PUT"])
def update_declarant(dec_id):
    """更新申報人（先驗歸屬）"""
    data = request.json or {}
    g_code = (data.get("g_code") or "").strip().upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    row = conn.execute("SELECT g_code FROM declarants WHERE id=?", (dec_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該申報人"})

    fields = {}
    for key in ["name", "address"]:
        if key in data:
            v = (data[key] or "").strip()
            if not v:
                conn.close()
                return jsonify({"success": False, "error": "申報人姓名、地址不可留空"})
            fields[key] = v
    if "phone" in data:
        phone = _normalize_tw_mobile(data.get("phone"))
        if not phone:
            conn.close()
            return jsonify({"success": False, "error": DECLARANT_PHONE_ERR})
        dup = conn.execute(
            "SELECT 1 FROM declarants WHERE g_code=? AND phone=? AND id!=?", (g_code, phone, dec_id)
        ).fetchone()
        if dup:
            conn.close()
            return jsonify({"success": False, "error": f"門號 {phone} 已經在你的申報人清單中"})
        fields["phone"] = phone
    if data.get("is_default"):
        conn.execute("UPDATE declarants SET is_default=0 WHERE g_code=?", (g_code,))
        fields["is_default"] = 1

    if fields:
        sets = ", ".join(f"{k}=?" for k in fields)
        conn.execute(f"UPDATE declarants SET {sets} WHERE id=?", list(fields.values()) + [dec_id])
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/declarants/<int:dec_id>", methods=["DELETE"])
def delete_declarant(dec_id):
    """刪除申報人；刪掉預設時由最舊一筆遞補"""
    data = request.json or {}
    g_code = (data.get("g_code") or request.args.get("g_code", "")).strip().upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    row = conn.execute("SELECT g_code, is_default FROM declarants WHERE id=?", (dec_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該申報人"})
    conn.execute("DELETE FROM declarants WHERE id=?", (dec_id,))
    if row["is_default"]:
        first = conn.execute(
            "SELECT id FROM declarants WHERE g_code=? ORDER BY id LIMIT 1", (g_code,)).fetchone()
        if first:
            conn.execute("UPDATE declarants SET is_default=1 WHERE id=?", (first["id"],))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/declarants/<int:dec_id>/default", methods=["POST"])
def set_default_declarant(dec_id):
    """設為預設申報人"""
    data = request.json or {}
    g_code = (data.get("g_code") or "").strip().upper()
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    row = conn.execute("SELECT g_code FROM declarants WHERE id=?", (dec_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該申報人"})
    conn.execute("UPDATE declarants SET is_default=0 WHERE g_code=?", (g_code,))
    conn.execute("UPDATE declarants SET is_default=1 WHERE id=?", (dec_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 加值服務目錄 API ============

@app.route("/api/extra_services/catalog", methods=["GET"])
def public_extra_service_catalog():
    """客戶端用：
      • 預設（出貨申請可勾清單）→ 只回 sel=True 固定價項目
      • ?full=1（費用價目表）→ 回全部項目（含 sel 旗標）
    無需登入。"""
    cat = get_extra_service_catalog()
    full = request.args.get("full") in ("1", "true", "yes")
    src = cat if full else [c for c in cat if c.get("sel")]
    items = [
        {"id": c.get("id"), "name": c.get("name", ""), "cat": c.get("cat", ""),
         "desc": c.get("desc", ""), "price": int(c.get("price") or 0), "sel": bool(c.get("sel"))}
        for c in src
    ]
    return jsonify({"success": True, "services": items})


@app.route("/api/admin/extra_services/catalog", methods=["GET"])
def admin_get_extra_service_catalog():
    """後台管理用：回傳完整目錄（含 sel 旗標與變動價項目）。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "僅主管理員可管理加值服務目錄"}), 403
    return jsonify({"success": True, "services": get_extra_service_catalog()})


@app.route("/api/admin/extra_services/catalog", methods=["POST"])
def admin_save_extra_service_catalog():
    """後台管理用：整批覆寫目錄。"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "僅主管理員可管理加值服務目錄"}), 403
    data = request.json or {}
    raw = data.get("services", [])
    if not isinstance(raw, list):
        return jsonify({"success": False, "error": "資料格式錯誤"}), 400
    cleaned = []
    # 先蒐集所有明確 id（避免自動配號撞到後面才出現的明確 id）
    explicit_ids = set()
    for c in raw:
        if isinstance(c, dict):
            cid = (c.get("id") or "").strip()
            if cid:
                explicit_ids.add(cid)
    used_ids = set()
    _counter = [0]
    def _fresh_id():
        while True:
            _counter[0] += 1
            cand = f"es{_counter[0]:02d}"
            if cand not in explicit_ids and cand not in used_ids:
                return cand
    for c in raw:
        if not isinstance(c, dict):
            continue
        name = (c.get("name") or "").strip()
        if not name:
            continue
        try:
            price = int(float(c.get("price") or 0))
        except (ValueError, TypeError):
            price = 0
        cid = (c.get("id") or "").strip()
        if not cid or cid in used_ids:   # 空的或重複 → 配一個不衝突的新 id
            cid = _fresh_id()
        used_ids.add(cid)
        cleaned.append({
            "id": cid,
            "name": name,
            "cat": (c.get("cat") or "").strip(),
            "desc": (c.get("desc") or "").strip(),
            "price": max(price, 0),
            "sel": bool(c.get("sel")),
        })
    conn = get_db()
    conn.execute(
        "INSERT INTO admin_settings (key, value) VALUES ('extra_service_catalog', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (json.dumps(cleaned, ensure_ascii=False),)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "count": len(cleaned)})


# ============ 台灣配送貨況 API ============

@app.route("/api/admin/tracking/status", methods=["GET"])
def admin_tracking_status():
    if not is_super_admin():
        return jsonify({"success": False, "error": "僅主管理員可操作"}), 403
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) AS c FROM delivery_tracking").fetchone()["c"]
    conn.close()
    return jsonify({
        "success": True,
        "last_sync": _get_setting("tracking_last_sync", ""),
        "sheet_url": _get_setting("tracking_sheet_url", DEFAULT_TRACKING_SHEET_URL),
        "total": total,
    })


@app.route("/api/admin/tracking/sync", methods=["POST"])
def admin_tracking_sync():
    if not is_super_admin():
        return jsonify({"success": False, "error": "僅主管理員可操作"}), 403
    data = request.json or {}
    new_url = (data.get("sheet_url") or "").strip()
    if new_url:
        _set_setting("tracking_sheet_url", new_url)
    try:
        count = sync_delivery_tracking()
        return jsonify({"success": True, "count": count, "last_sync": _get_setting("tracking_last_sync", "")})
    except Exception as e:
        return jsonify({"success": False, "error": f"同步失敗：{e}"}), 500


# ============ 出貨申請 API ============

@app.route("/api/shipment_request", methods=["POST"])
def create_shipment_request():
    """客戶申請出貨"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    customer_name = data.get("customer_name", "")
    package_ids = data.get("package_ids", [])
    note = (data.get("note") or "").strip()
    # 收件地址
    ship_recipient = (data.get("ship_recipient") or "").strip()
    ship_phone = (data.get("ship_phone") or "").strip()
    ship_address = (data.get("ship_address") or "").strip()
    # 地址完整性守門：缺縣市/區的地址黑貓無法投遞。不完整且客戶未確認 → 擋下請補
    if ship_address and not tw_zip.is_address_complete(ship_address) and not data.get("address_confirmed"):
        return jsonify({
            "success": False,
            "need_address_confirm": True,
            "error": "收件地址似乎缺少縣市或區（例：嘉義市西區），台灣宅配可能無法投遞。請確認或補齊地址。"
        }), 400
    # 客戶勾選的加值服務（[{id, qty}]）→ 以伺服端目錄價驗證後存入（防前端竄改價格）
    sel_services = data.get("extra_services", []) or []

    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    if not package_ids:
        return jsonify({"success": False, "error": "請選擇要出貨的包裹"})
    if not ship_recipient or not ship_phone or not ship_address:
        return jsonify({"success": False, "error": "請選擇寄送地址"})

    # 申報人同意聲明：只在客戶自己指定申報人時要求（沒指定 = fallback 回收件人本人，不需聲明）。
    # 前端 checkbox 只是體驗，繞過前端直接打 API 也必須擋，且要留下時間戳與來源 IP 當證據。
    declarant_ids = []
    for d in (data.get("declarant_ids") or []):
        try:
            declarant_ids.append(int(d))
        except (ValueError, TypeError):
            pass
    # 收緊真值判斷：form-encoded 呼叫（代理端介面、LINE 表單）送過來全是字串，
    # 用 Python 真值判斷會讓 "false" 變成有效同意。這是存證欄位，只收白名單。
    declarant_consent = 1 if (declarant_ids and data.get("declarant_consent")
                              in (True, 1, "1", "true", "True")) else 0
    declarant_consent_at = ""
    declarant_consent_ip = ""
    if declarant_ids:
        if not declarant_consent:
            return jsonify({"success": False, "error": "請先確認已取得申報人同意"}), 400
        declarant_consent_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # TZ=Asia/Taipei
        declarant_consent_ip = _client_ip()

    # 組合包裹摘要
    conn = get_db()

    # 重複申請檢查：阻擋同一包裹同時出現在多筆「進行中」的出貨申請
    # 進行中 = 待處理 / 處理中（已出貨會把 packages.status 更新為 已出貨，不會被選到；已拒絕視為釋出）
    requested_set = set()
    for p in package_ids:
        try:
            requested_set.add(int(p))
        except (ValueError, TypeError):
            pass

    active_reqs = conn.execute(
        "SELECT id, package_ids FROM shipment_requests "
        "WHERE g_code=? AND status IN ('待處理', '處理中')",
        (g_code,)
    ).fetchall()
    already_pending = set()
    for r in active_reqs:
        ids_str = r["package_ids"] or ""
        for pid_str in ids_str.split(","):
            try:
                already_pending.add(int(pid_str.strip()))
            except (ValueError, AttributeError):
                pass

    duplicates = requested_set & already_pending
    if duplicates:
        # 把重複的包裹 id 對應到 product_name 顯示得更友善
        dup_rows = conn.execute(
            f"SELECT id, product_name, logis_num FROM packages "
            f"WHERE id IN ({','.join(['?']*len(duplicates))})",
            list(duplicates)
        ).fetchall()
        dup_names = []
        for d in dup_rows:
            name = d["product_name"] or "未命名"
            logis = d["logis_num"] or ""
            dup_names.append(f"{name}（末四碼 {logis}）" if logis and logis != "-" else name)
        conn.close()
        return jsonify({
            "success": False,
            "error": f"以下包裹已在進行中的出貨申請中，無法重複申請：\n• " + "\n• ".join(dup_names) +
                     "\n\n請等管理員處理完成後再申請新出貨，或聯繫客服取消舊申請。"
        })

    placeholders = ",".join(["?"] * len(package_ids))
    rows = conn.execute(
        f"SELECT id, logis_num, product_name, weight FROM packages WHERE id IN ({placeholders}) AND g_code=?",
        package_ids + [g_code]
    ).fetchall()

    if not rows:
        conn.close()
        return jsonify({"success": False, "error": "找不到對應的包裹"})

    summary_parts = []
    total_weight = 0
    for idx, r in enumerate(rows, 1):
        r = dict(r)
        name = r["product_name"] or "商品"
        logis = r["logis_num"] or ""
        w = r["weight"] or ""
        line = f"{idx}. {name}"
        if w:
            line += f" / {w} kg"
        if logis and logis != "-":
            line += f" / {logis}"
        summary_parts.append(line)
        try:
            total_weight += float(w) if w else 0
        except:
            pass
    summary = "\n".join(summary_parts)
    if total_weight > 0:
        summary += f"\n合計約 {total_weight:.1f} kg"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ids_str = ",".join(str(i) for i in package_ids)
    sr_agent_id = get_agent_id_for_g_code(g_code)

    # 依伺服端目錄驗證客戶勾選：只收 sel=True 的項目、價格一律用目錄價、數量下限 1
    catalog = {c["id"]: c for c in get_extra_service_catalog() if c.get("sel")}
    customer_extras = []
    for s in sel_services:
        if not isinstance(s, dict):
            continue
        c = catalog.get(s.get("id"))
        if not c:
            continue
        try:
            qty = int(s.get("qty", 1))
        except (ValueError, TypeError):
            qty = 1
        if qty < 1:
            qty = 1
        price = int(c.get("price") or 0)
        # 客戶勾選時一律以 0 元存入的「意願型」加值服務（實際費用由現場/系統計算）：
        #  ・「合箱」：合箱費由系統依箱數計算（consolidation_fee）
        #  ・品名含「每件」：現場人員依實際包裝件數計算，勾選只表達需求
        if c["name"] == "合箱" or ("每件" in c["name"]):
            price = 0
        customer_extras.append({
            "id": c["id"], "name": c["name"], "qty": qty,
            "price": price, "subtotal": price * qty,
            "src": "customer",  # 客戶申請（管理員請款時可增刪，最終以帳單為準）
        })
    extra_services_json = json.dumps(customer_extras, ensure_ascii=False)

    # 申報人：只用 id 去 DB 撈（不信任前端傳的姓名/電話），且必須屬於本會員。
    # 撈不到或沒傳 → 四欄留空，出檔案時由 vendors 的三層 fallback 回收件人（維持現況）。
    declarants = []
    if declarant_ids:
        ph = ",".join(["?"] * len(declarant_ids))
        drows = conn.execute(
            f"SELECT name, phone, address, is_default FROM declarants "
            f"WHERE id IN ({ph}) AND g_code=? ORDER BY is_default DESC, id",
            declarant_ids + [g_code]
        ).fetchall()
        declarants = [{"name": d["name"], "phone": d["phone"], "address": d["address"]} for d in drows]
    declarants_json = json.dumps(declarants, ensure_ascii=False) if declarants else ""
    # 主申報人＝清單中的預設者（上面已 is_default DESC 排序），無則第一筆
    dec_main = declarants[0] if declarants else {}
    declarant_name = dec_main.get("name", "")
    declarant_phone = dec_main.get("phone", "")
    declarant_address = dec_main.get("address", "")

    conn.execute(
        """INSERT INTO shipment_requests (g_code, customer_name, package_ids, package_summary, status, note, ship_recipient, ship_phone, ship_address, extra_services, created_at, agent_id,
                                          declarant_name, declarant_phone, declarant_address, declarants_json,
                                          declarant_consent, declarant_consent_at, declarant_consent_ip)
           VALUES (?, ?, ?, ?, '待處理', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (g_code, customer_name, ids_str, summary, note, ship_recipient, ship_phone, ship_address, extra_services_json, now, sr_agent_id,
         declarant_name, declarant_phone, declarant_address, declarants_json,
         declarant_consent, declarant_consent_at, declarant_consent_ip)
    )
    conn.commit()
    conn.close()

    return jsonify({"success": True, "message": "出貨申請已送出，管理員會盡快處理！"})


@app.route("/api/shipment_requests", methods=["GET"])
def get_my_shipment_requests():
    """客戶查看自己的出貨申請"""
    g_code = request.args.get("g_code", "").upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM shipment_requests WHERE g_code=? ORDER BY id DESC", (g_code,)
    ).fetchall()

    # 台灣配送貨況：每筆出貨單用「存的 export_code ＋ 現算 {g_code}-{MMDD}」多候選比對，
    # 讓 export_code 上線前的舊單也能對到（MMDD 取 updated_at＝標記已出貨那天，其次 created_at）
    def _mmdd(v):
        try:
            return datetime.strptime(str(v)[:10], "%Y-%m-%d").strftime("%m%d")
        except (ValueError, TypeError):
            return None

    req_candidates = {}
    all_codes = set()
    for r in rows:
        cands, seen = [], set()
        for c in ([r["export_code"]] +
                  [f"{r['g_code']}-{mm}" for mm in (_mmdd(r["updated_at"]), _mmdd(r["created_at"])) if mm]):
            if c and c not in seen:
                seen.add(c); cands.append(c)
        req_candidates[r["id"]] = cands
        all_codes.update(cands)

    tmap = {}
    if all_codes:
        codes_list = list(all_codes)
        ph = ",".join(["?"] * len(codes_list))
        for t in conn.execute(
            f"SELECT customer_code, carrier, tracking_num FROM delivery_tracking WHERE customer_code IN ({ph})",
            codes_list
        ).fetchall():
            tmap[t["customer_code"]] = t
    conn.close()

    result = []
    for r in rows:
        d = dict(r)
        for c in req_candidates.get(r["id"], []):
            t = tmap.get(c)
            if t and t["tracking_num"]:
                d["delivery_carrier"] = t["carrier"]
                d["delivery_tracking"] = t["tracking_num"]
                d["delivery_url"] = delivery_tracking_url(t["carrier"], t["tracking_num"])
                break
        result.append(d)
    return jsonify({"success": True, "requests": result})


@app.route("/api/shipment_requests/<int:req_id>/payment", methods=["POST"])
def submit_payment_info(req_id):
    """客戶回報匯款後五碼"""
    data = request.json
    last5 = (data.get("last5") or "").strip()
    g_code = (data.get("g_code") or "").strip().upper()

    if not last5 or len(last5) != 5:
        return jsonify({"success": False, "error": "請輸入帳號後五碼（5位數字）"})
    if not last5.isdigit():
        return jsonify({"success": False, "error": "請輸入數字"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    last5 = last5.zfill(5)   # 純數字補滿5位，保住前導0（00000/00123 不被截）

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    # 確認是該客戶的申請
    row = conn.execute("SELECT g_code FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該申請"})

    conn.execute(
        "UPDATE shipment_requests SET payment_last5=?, payment_at=? WHERE id=?",
        (last5, now, req_id)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "匯款回報成功！"})


@app.route("/api/admin/old_packages", methods=["GET"])
def admin_old_packages():
    """查詢倉庫滯留超過 N 天的未出貨包裹（預設 30 天）"""
    try:
        days = int(request.args.get("days", 30))
    except (ValueError, TypeError):
        days = 30
    cutoff_date = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    aid = get_current_agent_id()
    conn = get_db()
    if aid > 0:
        rows = conn.execute(
            """SELECT * FROM packages
               WHERE status != '已出貨'
                 AND agent_id = ?
                 AND COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) <= ?
               ORDER BY COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) ASC""",
            (aid, cutoff_date)
        ).fetchall()
    else:
        rows = conn.execute(
            """SELECT * FROM packages
               WHERE status != '已出貨'
                 AND COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) <= ?
               ORDER BY COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) ASC""",
            (cutoff_date,)
        ).fetchall()
    conn.close()
    today = datetime.now().date()
    result = []
    for r in rows:
        d = dict(r)
        ref_date_str = d.get("in_date") or (d.get("created_at") or "")[:10]
        try:
            ref_date = datetime.strptime(ref_date_str, "%Y-%m-%d").date()
            age_days = (today - ref_date).days
        except (ValueError, TypeError):
            age_days = 0
        d["age_days"] = age_days
        d["ref_date"] = ref_date_str
        result.append(d)
    return jsonify({
        "success": True,
        "count": len(result),
        "days": days,
        "packages": result
    })


@app.route("/api/admin/customer_unpaid/<g_code>", methods=["GET"])
def admin_customer_unpaid(g_code):
    """查詢客戶未付款的已出貨筆數和金額（用於出貨警告）"""
    # 代理只能查自己的客戶
    aid = get_current_agent_id()
    if aid > 0 and get_agent_id_for_g_code(g_code.upper()) != aid:
        return jsonify({"success": True, "count": 0, "total": 0, "latest": "", "ids": []})
    conn = get_db()
    rows = conn.execute(
        "SELECT id, total_fee, updated_at FROM shipment_requests "
        "WHERE g_code=? AND status='已出貨' AND total_fee > 0 "
        "AND (payment_last5 IS NULL OR payment_last5='') "
        "ORDER BY updated_at DESC",
        (g_code.upper(),)
    ).fetchall()
    conn.close()
    items = [dict(r) for r in rows]
    return jsonify({
        "success": True,
        "count": len(items),
        "total": sum(int(r.get("total_fee") or 0) for r in items),
        "latest": items[0]["updated_at"] if items else "",
        "ids": [r["id"] for r in items]
    })


@app.route("/api/admin/shipment_requests/<int:req_id>/confirm_payment", methods=["POST"])
def admin_confirm_payment(req_id):
    """管理員確認匯款已收到（可填後五碼、LINE Pay、現金等任意備註）"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    note = (data.get("last5") or "").strip()

    if len(note) > 20:
        return jsonify({"success": False, "error": "備註請勿超過 20 字"})
    if not note:
        note = "管確認"
    elif note.isdigit():
        note = note.zfill(5)   # 純數字末五碼補滿5位，保住前導0（00000 等）

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    conn.execute(
        "UPDATE shipment_requests SET payment_last5=?, payment_at=? WHERE id=?",
        (note, now, req_id)
    )
    conn.commit()
    conn.close()
    log_op("帳單確認付款", f"出貨單#{req_id}", f"後五碼 {note}")
    # 只影響快照的 paid_kg（費率用的是 total_kg，與收款無關）
    _resync_snapshot(_row.get("g_code"), (_row.get("updated_at") or "")[:7])
    return jsonify({"success": True, "message": "已確認匯款"})


@app.route("/api/admin/shipment_requests/<int:req_id>/unconfirm_payment", methods=["POST"])
def admin_unconfirm_payment(req_id):
    """管理員取消已確認的匯款（誤按時用）"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    conn.execute(
        "UPDATE shipment_requests SET payment_last5='', payment_at='' WHERE id=?",
        (req_id,)
    )
    conn.commit()
    conn.close()
    # 取消付款確認 = 解除金額鎖定（見 _paid_lock_diff），一定要留下誰、何時、解了哪張單
    log_op("帳單取消付款確認", f"出貨單#{req_id}",
           f"原後五碼 {_row.get('payment_last5') or ''}；金額鎖定解除")
    _resync_snapshot(_row.get("g_code"), (_row.get("updated_at") or "")[:7])
    return jsonify({"success": True, "message": "已取消匯款確認"})


# ============ 出檔案給廠商（Nigel / JpD ） ============

@app.route("/api/admin/vendors", methods=["GET"])
def admin_list_vendors():
    """前端 UI 廠商下拉選單用"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    try:
        return jsonify({"success": True, "vendors": vendor_templates.list_vendors()})
    except Exception as e:
        import traceback
        print(f"[vendors] 💥 例外:\n{traceback.format_exc()}", flush=True)
        return jsonify({"success": False, "error": f"{type(e).__name__}: {e}"}), 500


@app.route("/api/admin/customer_vendor_codes", methods=["GET"])
def admin_get_vendor_codes():
    """查詢一批客戶的廠商編號（FWT0001 等）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    g_codes_str = request.args.get("g_codes", "")
    g_codes = [s.strip().upper() for s in g_codes_str.split(",") if s.strip()]
    conn = get_db()
    if g_codes:
        placeholders = ",".join(["?"] * len(g_codes))
        rows = conn.execute(
            f"SELECT g_code, vendor, code FROM customer_vendor_codes WHERE g_code IN ({placeholders})",
            g_codes
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT g_code, vendor, code FROM customer_vendor_codes ORDER BY g_code, vendor"
        ).fetchall()
    conn.close()
    # 回傳 nested dict: {g_code: {vendor: code}}
    result = {}
    for r in rows:
        result.setdefault(r["g_code"], {})[r["vendor"]] = r["code"]
    return jsonify({"success": True, "codes": result})


@app.route("/api/admin/customer_vendor_codes", methods=["POST"])
def admin_set_vendor_code():
    """設定／更新某客戶的廠商編號"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    g_code = (data.get("g_code") or "").strip().upper()
    vendor = (data.get("vendor") or "").strip().lower()
    code = (data.get("code") or "").strip()
    if not g_code or not vendor:
        return jsonify({"success": False, "error": "缺少 g_code 或 vendor"})
    if vendor not in vendor_templates.VENDORS:
        return jsonify({"success": False, "error": f"未知廠商：{vendor}"})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    if code:
        # INSERT OR REPLACE
        conn.execute(
            "INSERT INTO customer_vendor_codes (g_code, vendor, code, updated_at) VALUES (?,?,?,?) "
            "ON CONFLICT(g_code, vendor) DO UPDATE SET code=excluded.code, updated_at=excluded.updated_at",
            (g_code, vendor, code, now)
        )
    else:
        # 空字串 = 刪除
        conn.execute(
            "DELETE FROM customer_vendor_codes WHERE g_code=? AND vendor=?",
            (g_code, vendor)
        )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "已儲存"})


@app.route("/api/admin/exports/pending", methods=["GET"])
def admin_exports_pending():
    """列出已付款但未匯出給廠商的出貨單"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    try:
        return _admin_exports_pending_impl()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(f"[exports/pending] 💥 例外:\n{tb}", flush=True)
        # 回 JSON 而不是 HTML 500，讓前端能解析
        return jsonify({
            "success": False,
            "error": f"後端錯誤: {type(e).__name__}: {e}",
            "traceback_excerpt": tb.splitlines()[-1] if tb else "",
        }), 500


def _admin_exports_pending_impl():
    conn = get_db()
    rows = conn.execute("""
        SELECT * FROM shipment_requests
        WHERE status='已出貨'
          AND payment_last5 IS NOT NULL AND payment_last5 != ''
          AND (exported_at IS NULL OR exported_at = '')
        ORDER BY payment_at ASC, id ASC
    """).fetchall()

    # 撈所有相關 packages
    pkg_ids = set()
    for r in rows:
        pkg_ids.update(_parse_pkg_ids(r["package_ids"]))

    pkg_map = {}
    if pkg_ids:
        placeholders = ",".join(["?"] * len(pkg_ids))
        pkg_rows = conn.execute(
            f"SELECT id, g_code, logis_num, product_name, weight FROM packages WHERE id IN ({placeholders})",
            list(pkg_ids)
        ).fetchall()
        for p in pkg_rows:
            pkg_map[p["id"]] = dict(p)

    # 撈 vendor codes
    g_codes = list({r["g_code"] for r in rows})
    codes_map = {}
    if g_codes:
        placeholders = ",".join(["?"] * len(g_codes))
        code_rows = conn.execute(
            f"SELECT g_code, vendor, code FROM customer_vendor_codes WHERE g_code IN ({placeholders})",
            g_codes
        ).fetchall()
        for c in code_rows:
            codes_map.setdefault(c["g_code"], {})[c["vendor"]] = c["code"]

    # Fallback 資料來源（與 /generate 一致）：members 表 + Shopify cache
    members_map = {}
    if g_codes:
        placeholders = ",".join(["?"] * len(g_codes))
        for m in conn.execute(
            f"SELECT g_code, name, phone, address FROM members WHERE g_code IN ({placeholders})",
            g_codes
        ).fetchall():
            members_map[m["g_code"]] = {"name": m["name"], "phone": m["phone"], "address": m["address"]}
    # Shopify 客戶 fallback：M2 起讀本地 members_shopify（毫秒級，不會阻塞請求）
    shopify_map = {}
    try:
        for c in get_all_goyoutati_customers():
            if c.get("g_code") in g_codes:
                shopify_map[c["g_code"]] = {"name": c.get("name", ""), "phone": c.get("phone", ""), "address": c.get("address", "")}
    except Exception as e:
        print(f"[export-pending] 會員資料讀取失敗（不致命）: {e}", flush=True)
    conn.close()

    items = []
    for r in rows:
        rd = dict(r)
        pids = _parse_pkg_ids(rd.get("package_ids"))
        # ship_* 為空時用 fallback 在 UI 也能看到正確資料
        ship_recipient = _safe_str(rd.get("ship_recipient"))
        ship_phone     = _safe_str(rd.get("ship_phone"))
        ship_address   = _safe_str(rd.get("ship_address"))
        if not (ship_recipient and ship_phone and ship_address):
            fb = members_map.get(rd["g_code"]) or shopify_map.get(rd["g_code"]) or {}
            if not ship_recipient: ship_recipient = _safe_str(fb.get("name")) or _safe_str(rd.get("customer_name"))
            if not ship_phone:     ship_phone     = _safe_str(fb.get("phone"))
            if not ship_address:   ship_address   = _safe_str(fb.get("address"))

        items.append({
            "id":               rd["id"],
            "g_code":           rd["g_code"],
            "customer_name":    rd.get("customer_name") or "",
            "ship_recipient":   ship_recipient,
            "ship_phone":       ship_phone,
            "ship_address":     ship_address,
            "declarant_name":   rd.get("declarant_name") or "",
            "declarant_phone":  rd.get("declarant_phone") or "",
            "declarant_address": rd.get("declarant_address") or "",
            "declarants_json":  rd.get("declarants_json") or "",
            "declarant_consent":    rd.get("declarant_consent") or 0,
            "declarant_consent_at": rd.get("declarant_consent_at") or "",
            "billed_weight":    rd.get("billed_weight") or 0,
            "total_fee":        rd.get("total_fee") or 0,
            "payment_at":       rd.get("payment_at") or "",
            "payment_last5":    rd.get("payment_last5") or "",
            "updated_at":       rd.get("updated_at") or "",
            "package_count":    len(pids),
            "packages":         [pkg_map[i] for i in pids if i in pkg_map],
            "vendor_codes":     codes_map.get(rd["g_code"], {}),
        })
    return jsonify({"success": True, "items": items})


@app.route("/api/admin/exports/generate", methods=["POST"])
def admin_exports_generate():
    """匯出選定的出貨單為廠商 Excel，並標記 exported_*"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    try:
        return _admin_exports_generate_impl()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(f"[exports/generate] 💥 例外:\n{tb}", flush=True)
        return jsonify({
            "success": False,
            "error": f"後端錯誤: {type(e).__name__}: {e}",
            "traceback_excerpt": tb.splitlines()[-1] if tb else "",
        }), 500


def _admin_exports_generate_impl():
    data = request.json or {}
    vendor_id = (data.get("vendor") or "").strip().lower()
    ids = data.get("ids") or []

    vendor = vendor_templates.get_vendor(vendor_id)
    if not vendor:
        return jsonify({"success": False, "error": f"未知廠商：{vendor_id}"}), 400
    if not ids:
        return jsonify({"success": False, "error": "請至少選一筆"}), 400
    ids = [int(x) for x in ids if str(x).isdigit()]
    if not ids:
        return jsonify({"success": False, "error": "無有效 ID"}), 400

    conn = get_db()
    placeholders = ",".join(["?"] * len(ids))
    rows = conn.execute(
        f"""SELECT * FROM shipment_requests
            WHERE id IN ({placeholders})
              AND status='已出貨'
              AND payment_last5 IS NOT NULL AND payment_last5 != ''
              AND (exported_at IS NULL OR exported_at = '')
        """,
        ids
    ).fetchall()
    if not rows:
        conn.close()
        return jsonify({"success": False, "error": "選定的單都已匯出或狀態不符（可能被別人剛剛搶先匯出了）"}), 400

    # 撈所有相關 packages
    all_pkg_ids = set()
    for r in rows:
        all_pkg_ids.update(_parse_pkg_ids(r["package_ids"]))
    pkg_map = {}
    if all_pkg_ids:
        ph = ",".join(["?"] * len(all_pkg_ids))
        for p in conn.execute(
            f"SELECT id, g_code, logis_num, product_name, weight FROM packages WHERE id IN ({ph})",
            list(all_pkg_ids)
        ).fetchall():
            pkg_map[p["id"]] = dict(p)

    # 撈會員資料做 fallback（舊出貨單 ship_* 欄位可能空白）
    g_codes_needed = list({r["g_code"] for r in rows})
    members_map = {}  # g_code → {name, phone, address}
    if g_codes_needed:
        ph = ",".join(["?"] * len(g_codes_needed))
        for m in conn.execute(
            f"SELECT g_code, name, phone, address FROM members WHERE g_code IN ({ph})",
            g_codes_needed
        ).fetchall():
            members_map[m["g_code"]] = {"name": m["name"], "phone": m["phone"], "address": m["address"]}
    # Shopify 客戶 fallback：M2 起讀本地 members_shopify（毫秒級，不會阻塞請求）
    shopify_map = {}
    try:
        for c in get_all_goyoutati_customers():
            if c.get("g_code") in g_codes_needed:
                shopify_map[c["g_code"]] = {"name": c.get("name", ""), "phone": c.get("phone", ""), "address": c.get("address", "")}
    except Exception as ex:
        print(f"[export] Shopify fallback 失敗（不致命）: {ex}", flush=True)

    # 組 shipments list
    shipments = []
    fallback_updates = []  # (id, ship_recipient, ship_phone, ship_address) - 把 fallback 後的值寫回 DB
    missing_pkg_count = 0
    for r in rows:
        rd = dict(r)
        pids = _parse_pkg_ids(rd.get("package_ids"))
        # 包裹資料：找到的用真實值、找不到的用 stub
        # vendor 範本實際上只用 package_id 當隨機種子，不用 logis_num/weight 等具體欄位
        # 所以孤立資料（migration 後 packages 表沒對應）也能撐過 Excel 產出
        pkgs = []
        for i in pids:
            if i in pkg_map:
                pkgs.append(pkg_map[i])
            else:
                pkgs.append({"id": i, "g_code": rd["g_code"], "logis_num": "", "product_name": "", "weight": 0})
                missing_pkg_count += 1
        # 若 package_ids 字串完全解析不出任何整數 → 真的沒辦法產，跳過
        # （容錯解析已支援 "5.0" 型髒資料；走到這代表原始值真的空/全壞）
        if not pkgs:
            print(f"[export] ⚠️ shipment id={rd.get('id')} g_code={rd.get('g_code')} 的 "
                  f"package_ids 解析後為空，跳過（原始值={rd.get('package_ids')!r}）", flush=True)
            continue

        # Fallback 順序：shipment_requests.ship_* → members 表 → Shopify cache → customer_name
        ship_recipient = _safe_str(rd.get("ship_recipient"))
        ship_phone     = _safe_str(rd.get("ship_phone"))
        ship_address   = _safe_str(rd.get("ship_address"))
        if not (ship_recipient and ship_phone and ship_address):
            g_code = rd["g_code"]
            fallback_src = members_map.get(g_code) or shopify_map.get(g_code) or {}
            if not ship_recipient: ship_recipient = _safe_str(fallback_src.get("name")) or _safe_str(rd.get("customer_name"))
            if not ship_phone:     ship_phone     = _safe_str(fallback_src.get("phone"))
            if not ship_address:   ship_address   = _safe_str(fallback_src.get("address"))
            # 如果填到任何值，順手寫回 DB（下次匯出不用再 fallback）
            if ship_recipient or ship_phone or ship_address:
                fallback_updates.append((ship_recipient, ship_phone, ship_address, rd["id"]))

        shipments.append({
            "id":                   rd["id"],
            "g_code":               rd["g_code"],
            "ship_recipient":       ship_recipient,
            "ship_phone":           ship_phone,
            "ship_address":         ship_address,
            "billed_weight":        rd.get("billed_weight") or 0,
            "total_fee":            rd.get("total_fee") or 0,
            # 出貨追蹤號碼（多箱換行）→ Nigel 填「清關號碼」/ JpD 填「JpD包裹ID」
            "tracking_num":         _safe_str(rd.get("tracking_num")),
            # 打包日期來源：admin 標記出貨時的 updated_at（fallback 到客戶申請的 created_at）
            "updated_at":           rd.get("updated_at") or "",
            "created_at":           rd.get("created_at") or "",
            "packages":             pkgs,
            "boxes":                _parse_boxes(rd.get("boxes_json")),
            # 主申報人（每箱未指定時的預設值）；空 → vendors 內 fallback 回收件人
            "declarant_name":       _safe_str(rd.get("declarant_name")),
            "declarant_phone":      _safe_str(rd.get("declarant_phone")),
            "declarant_address":    _safe_str(rd.get("declarant_address")),
            # 同意聲明只是隨行資料，vendors 的欄位表是固定的，不會進到給廠商的檔案
            "declarant_consent":    rd.get("declarant_consent") or 0,
            "declarant_consent_at": _safe_str(rd.get("declarant_consent_at")),
        })

    if not shipments:
        conn.close()
        return jsonify({"success": False, "error": "選定的單沒有包裹資料"}), 400

    # ── 申報人單日報關件數上限檢查 ──
    # 報關單位是「出檔案批次」，所以硬檢查點在這裡。本批 + 同日已出檔案的批次一起算，
    # 否則同一天出兩批、各自沒超過、加起來就爆了。
    over_limit = _check_declarant_box_limit(conn, shipments)
    if over_limit and not data.get("confirm_over_limit"):
        conn.close()
        return jsonify({
            "success": False,
            "need_confirm": True,
            "error": "有申報人超過單日報關件數上限，請重新分箱或確認後繼續",
            "limit": _get_box_limit(),
            "over_limit": over_limit,
        }), 409

    # 產生 Excel
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    headers, table_rows = vendor_templates.build_rows(vendor_id, shipments)

    wb = Workbook()
    ws = wb.active
    ws.title = vendor["display_name"]

    # 標頭樣式
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    hdr_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))

    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = thin

    for row_idx, row_data in enumerate(table_rows, start=2):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin
            # 多箱追蹤號碼是換行字串 → 開自動換行才不會擠成一行
            if isinstance(val, str) and "\n" in val:
                cell.alignment = Alignment(wrap_text=True, vertical="center")

    # 欄寬自動
    for col_idx, h in enumerate(headers, start=1):
        max_len = len(str(h))
        for row_data in table_rows:
            v = row_data[col_idx - 1] if col_idx - 1 < len(row_data) else ""
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 4, 40)

    # 標記 exported
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch_id = f"{vendor_id}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    shipment_ids_actually_used = [s["id"] for s in shipments]
    ph = ",".join(["?"] * len(shipment_ids_actually_used))
    conn.execute(
        f"""UPDATE shipment_requests
            SET exported_at=?, exported_vendor=?, exported_batch_id=?
            WHERE id IN ({ph})
        """,
        [now, vendor_id, batch_id] + shipment_ids_actually_used
    )
    # 存 export_code（{g_code}-{MMDD}，與 Excel 客戶編號一致）供台灣配送貨況比對
    conn.executemany(
        "UPDATE shipment_requests SET export_code=? WHERE id=?",
        [(vendor_templates.export_code_for(s), s["id"]) for s in shipments]
    )
    # 順手把 fallback 出來的 ship_* 值寫回（下次匯出不用再算）
    if fallback_updates:
        conn.executemany(
            "UPDATE shipment_requests SET ship_recipient=?, ship_phone=?, ship_address=? WHERE id=?",
            fallback_updates
        )
        print(f"[export] fallback 補回 ship_* 欄位 {len(fallback_updates)} 筆", flush=True)
    conn.commit()
    conn.close()

    # 強制略過上限檢查一定要留紀錄（可能有正當例外，但要查得到是誰在什麼時候放行的）
    if over_limit:
        detail = "；".join(
            f"{o['declarant_name']}({o['declarant_phone']}) 本批{o['batch_boxes']}"
            f"+同日{o['same_day_boxes']}={o['total']}>上限{o['limit']}"
            for o in over_limit
        )
        log_op("超出申報人件數上限出檔案", f"批次{batch_id}", detail)
        print(f"[export] ⚠️ 超出申報人單日件數上限仍出檔案：{detail}", flush=True)

    # 輸出檔案
    filename = vendor_templates.filename_for(vendor_id)
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    from urllib.parse import quote
    response = make_response(bio.read())
    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    response.headers["Content-Disposition"] = f'attachment; filename="{quote(filename)}"; filename*=UTF-8\'\'{quote(filename)}'
    response.headers["X-Batch-Id"] = batch_id
    response.headers["X-Shipments-Exported"] = str(len(shipments))
    if missing_pkg_count:
        response.headers["X-Missing-Packages"] = str(missing_pkg_count)
        print(f"[export] ⚠️ 本批有 {missing_pkg_count} 個包裹資料缺失（用 stub 撐過），shipments={len(shipments)} 筆", flush=True)
    return response


@app.route("/api/admin/exports/history", methods=["GET"])
def admin_exports_history():
    """檢視歷史批次（最近 50 批）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    rows = conn.execute("""
        SELECT exported_batch_id, exported_vendor, MIN(exported_at) AS exported_at,
               COUNT(*) AS shipment_count,
               SUM(billed_weight) AS total_kg,
               SUM(total_fee) AS total_fee
        FROM shipment_requests
        WHERE exported_at IS NOT NULL AND exported_at != ''
        GROUP BY exported_batch_id, exported_vendor
        ORDER BY MIN(exported_at) DESC
        LIMIT 50
    """).fetchall()
    conn.close()
    return jsonify({
        "success": True,
        "batches": [dict(r) for r in rows]
    })


@app.route("/api/admin/shipment_requests", methods=["GET"])
def admin_get_shipment_requests():
    """管理員查看所有出貨申請（含對應客戶的待處理預報資料）"""
    maybe_auto_sync()  # 後台有人活動時，距上次同步>24h 就背景同步台灣配送貨況
    status = request.args.get("status", "")
    pay = (request.args.get("pay") or "").strip()   # 帳單付款狀態：unpaid / paid / all
    q = (request.args.get("q") or "").strip()
    date_from = (request.args.get("date_from") or "").strip()
    date_to = (request.args.get("date_to") or "").strip()
    date_field = (request.args.get("date_field") or "auto").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except (ValueError, TypeError):
        page = 1
    try:
        limit = min(200, max(1, int(request.args.get("limit", 50))))
    except (ValueError, TypeError):
        limit = 50
    offset = (page - 1) * limit
    aid = get_current_agent_id()
    try:
        conn = get_db()
        where, params = [], []
        # 帳單付款狀態（優先於 status；供帳單管理分頁用）
        if pay == "unpaid":
            where.append("status='已出貨' AND total_fee>0 AND (payment_last5 IS NULL OR payment_last5='')")
            order = "ORDER BY id DESC"
        elif pay == "paid":
            where.append("status='已出貨' AND payment_last5 IS NOT NULL AND payment_last5!=''")
            order = "ORDER BY payment_at DESC, id DESC"
        elif pay == "all":
            where.append("status='已出貨'")
            order = "ORDER BY id DESC"
        # 狀態（出貨申請用）
        elif status == "已付款":
            where.append("status='已出貨' AND payment_last5 != '' AND payment_last5 IS NOT NULL")
            order = "ORDER BY payment_at DESC, id DESC"
        elif status and status != "recent":
            where.append("status=?"); params.append(status)
            order = "ORDER BY id DESC"
        else:
            order = "ORDER BY id DESC"
        # 代理過濾
        if aid > 0:
            where.append("agent_id=?"); params.append(aid)
        # 關鍵字（後端查，不再前端全撈）
        if q:
            like = f"%{q}%"
            fields = ["g_code", "customer_name", "ship_recipient", "ship_phone",
                      "ship_address", "note", "tracking_num", "package_summary", "payment_last5"]
            where.append("(" + " OR ".join(f"{f} LIKE ?" for f in fields) + ")")
            params += [like] * len(fields)
        # 日期區間（auto：已出貨/已付款→出貨日 updated_at，其他→申請日 created_at）
        if date_from or date_to:
            df = date_field
            if df == "auto":
                df = "updated_at" if status in ("已出貨", "已付款") else "created_at"
            col = f"date(replace(substr(COALESCE(NULLIF({df},''), created_at),1,10),'/','-'))"
            if date_from:
                where.append(f"{col} >= date(?)"); params.append(date_from)
            if date_to:
                where.append(f"{col} <= date(?)"); params.append(date_to)

        wsql = (" WHERE " + " AND ".join(where)) if where else ""
        total = conn.execute(f"SELECT COUNT(*) AS c FROM shipment_requests{wsql}", params).fetchone()["c"]
        # 帳單管理需要「全部符合」的加總（列印摘要用），一次算好
        sum_total = sum_weight = 0
        if pay:
            srow = conn.execute(
                f"SELECT COALESCE(SUM(total_fee),0) AS st, COALESCE(SUM(billed_weight),0) AS sw FROM shipment_requests{wsql}",
                params
            ).fetchone()
            sum_total = round(srow["st"] or 0)
            sum_weight = round((srow["sw"] or 0), 1)
        rows = conn.execute(
            f"SELECT * FROM shipment_requests{wsql} {order} LIMIT ? OFFSET ?",
            params + [limit, offset]
        ).fetchall()

        # 一次撈出涉及到的客戶的待處理預報（避免 N+1 查詢）
        g_codes = list({r["g_code"] for r in rows if r["g_code"]})
        forecast_map = {}
        if g_codes:
            placeholders = ",".join(["?"] * len(g_codes))
            fc_rows = conn.execute(
                f"SELECT * FROM forecasts WHERE g_code IN ({placeholders}) AND status='待處理' ORDER BY id",
                g_codes
            ).fetchall()
            for fc in fc_rows:
                try:
                    items = json.loads(fc["items_json"] or "[]")
                except (TypeError, ValueError, json.JSONDecodeError):
                    items = []
                forecast_map.setdefault(fc["g_code"], []).append({
                    "id": fc["id"],
                    "note": fc["note"] or "",
                    "created_at": fc["created_at"] or "",
                    "items": items
                })

        # 每筆出貨單的「信件」件數 → 帳單自動帶入信件費（件數 × NT$20）
        req_pids = {}
        all_pids = set()
        for r in rows:
            pids = _parse_pkg_ids(r["package_ids"])
            req_pids[r["id"]] = pids
            all_pids.update(pids)
        letter_ids = set()
        if all_pids:
            ph = ",".join(["?"] * len(all_pids))
            trows = conn.execute(
                f"SELECT id FROM packages WHERE id IN ({ph}) AND pkg_type='信件'",
                list(all_pids)
            ).fetchall()
            letter_ids = {t["id"] for t in trows}

        conn.close()

        result = []
        for r in rows:
            d = dict(r)
            d["pending_forecasts"] = forecast_map.get(r["g_code"], [])
            d["letter_count"] = sum(1 for pid in req_pids.get(r["id"], []) if pid in letter_ids)
            result.append(d)
        return jsonify({"success": True, "requests": result,
                        "total": total, "page": page, "limit": limit,
                        "has_more": offset + len(rows) < total,
                        "sum_total": sum_total, "sum_weight": sum_weight})
    except Exception as e:
        return jsonify({"success": False, "error": str(e), "requests": []})


# ============ 已付款帳單金額鎖定 ============
# 老闆決定：payment_last5 非空（= 已收款）之後，訂單金額不能再被改變。
# 2026-09 的事故：計費 modal 每次開啟都用現行規則重算，員工開已付款的舊單看一眼、
# 手滑按儲存，total_fee 就從 211 變 422，系統憑空多出一筆欠款。
# 前端擋只是體驗，這裡才是防線（前端規則在某台電腦上沒生效就是這次的根因）。
# 解鎖只有一條路：既有的「取消付款確認」→ payment_last5 清空 → 改完再確認付款。

PAID_LOCK_FIELDS = ("billed_weight", "rate_per_kg", "shipping_fee", "handling_fee",
                    "consolidation_fee", "letter_fee", "total_fee")
PAID_LOCK_BOX_FIELDS = ("actual_weight", "length", "width", "height", "billed_weight")
PAID_LOCK_TOL = 0.01
PAID_LOCK_ERROR = "此單已收款，金額不可修改。如需調整，請先在帳單列表點「取消付款確認」。"


def _is_paid(row):
    """已付款定義：payment_last5 非空字串（客戶自報與管理員確認寫同一欄位，取較嚴格認定）。"""
    return bool(str((row or {}).get("payment_last5") or "").strip())


def _lock_num(v):
    """金額欄位一律轉 float 比較；空字串 / None / 解析失敗 → 0。"""
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _lock_extras_canon(v):
    """extra_services 正規化：只看 name / qty / price / subtotal（stored 可能多帶 src 等欄位），
    數字四捨五入到分、依內容排序 → 順序不同不算改金額。"""
    if isinstance(v, str):
        v = _parse_boxes(v)          # 同樣的「壞 JSON 回空 list」語意
    if not isinstance(v, list):
        v = []
    out = []
    for e in v:
        if not isinstance(e, dict):
            continue
        out.append((str(e.get("name") or "").strip(),
                    round(_lock_num(e.get("qty")), 2),
                    round(_lock_num(e.get("price")), 2),
                    round(_lock_num(e.get("subtotal")), 2)))
    return sorted(out)


def _paid_lock_diff(stored, data):
    """比對送進來的金額欄位與現存值，回 [{欄位, 現存值, 送來的值}]；空 list = 沒動到金額。"""
    diff = []
    for f in PAID_LOCK_FIELDS:
        cur, new = _lock_num(stored.get(f)), _lock_num(data.get(f))
        if abs(cur - new) > PAID_LOCK_TOL:
            diff.append({"欄位": f, "現存值": stored.get(f), "送來的值": data.get(f)})

    cur_ex, new_ex = _lock_extras_canon(stored.get("extra_services")), _lock_extras_canon(data.get("extra_services", []))
    if cur_ex != new_ex:
        diff.append({"欄位": "extra_services", "現存值": stored.get("extra_services"),
                     "送來的值": data.get("extra_services", [])})

    cur_boxes = [b for b in _parse_boxes(stored.get("boxes_json")) if isinstance(b, dict)]
    new_boxes = [b for b in (data.get("boxes") or []) if isinstance(b, dict)]
    if len(cur_boxes) != len(new_boxes):
        diff.append({"欄位": "boxes.count", "現存值": len(cur_boxes), "送來的值": len(new_boxes)})
    else:
        for i, (cb, nb) in enumerate(zip(cur_boxes, new_boxes)):
            for f in PAID_LOCK_BOX_FIELDS:
                if abs(_lock_num(cb.get(f)) - _lock_num(nb.get(f))) > PAID_LOCK_TOL:
                    diff.append({"欄位": f"boxes[{i}].{f}", "現存值": cb.get(f), "送來的值": nb.get(f)})
    return diff


@app.route("/api/admin/shipment_requests/<int:req_id>", methods=["PUT"])
def admin_update_shipment_request(req_id):
    """管理員更新出貨申請狀態（含帳單資訊）"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    status = data.get("status", "")
    admin_note = data.get("admin_note", "")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db()

    # 帳單欄位（出貨時填寫）
    billed_weight = data.get("billed_weight", 0)
    rate_per_kg = data.get("rate_per_kg", 0)
    shipping_fee = data.get("shipping_fee", 0)
    handling_fee = data.get("handling_fee", 0)
    consolidation_fee = data.get("consolidation_fee", 0)
    letter_fee = data.get("letter_fee", 0)
    total_fee = data.get("total_fee", 0)
    tracking_num = data.get("tracking_num", "")
    extra_services = json.dumps(data.get("extra_services", []), ensure_ascii=False)
    boxes_json = json.dumps(data.get("boxes", []), ensure_ascii=False)  # 多箱明細（空=單箱舊制）

    # 代理可自由設定費率，無下限（min_rate 僅為 UI 預設值參考）

    if status == "已出貨" and billed_weight:
        # 已收款 → 金額欄位一個都不准動；整筆拒絕、不做部分寫入。
        # 追蹤號 / 備註 / 每箱申報人與追蹤號不在鎖定範圍（金額原封不動送回就照常通過）。
        if _is_paid(_row):
            lock_diff = _paid_lock_diff(_row, data)
            if lock_diff:
                conn.close()
                return jsonify({"success": False, "locked": True,
                                "error": PAID_LOCK_ERROR, "diff": lock_diff}), 409
        conn.execute(
            """UPDATE shipment_requests 
               SET status=?, admin_note=?, updated_at=?,
                   billed_weight=?, rate_per_kg=?, shipping_fee=?, handling_fee=?, consolidation_fee=?, letter_fee=?, total_fee=?,
                   tracking_num=?, extra_services=?, boxes_json=?
               WHERE id=?""",
            (status, admin_note, now, billed_weight, rate_per_kg, shipping_fee, handling_fee, consolidation_fee, letter_fee, total_fee, tracking_num, extra_services, boxes_json, req_id)
        )
    else:
        conn.execute(
            "UPDATE shipment_requests SET status=?, admin_note=?, updated_at=? WHERE id=?",
            (status, admin_note, now, req_id)
        )

    # 如果管理員標記為「已出貨」，同步更新包裹狀態 + 預報標為已處理
    g_code_val = ""
    customer_name_val = ""
    if status == "已出貨":
        req = conn.execute("SELECT package_ids, g_code, customer_name FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
        if req:
            g_code_val = req["g_code"]
            customer_name_val = req["customer_name"] or ""
            pkg_ids = [int(x.strip()) for x in req["package_ids"].split(",") if x.strip()]
            if pkg_ids:
                placeholders = ",".join(["?"] * len(pkg_ids))
                conn.execute(
                    f"UPDATE packages SET status='已出貨' WHERE id IN ({placeholders})", pkg_ids
                )
            # 自動把該客戶的待處理預報標為已處理
            conn.execute(
                "UPDATE forecasts SET status='已處理' WHERE g_code=? AND status='待處理'",
                (g_code_val,)
            )
    else:
        req = conn.execute("SELECT g_code, customer_name FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
        if req:
            g_code_val = req["g_code"]
            customer_name_val = req["customer_name"] or ""

    conn.commit()
    conn.close()
    rate_override = None
    if status == "已出貨":
        log_op("出貨處理", g_code_val, f"出貨單#{req_id}")
        # ── 階梯費率覆寫留痕 ──
        # 伺服器只提供「預設值」，管理員仍可在帳單視窗改（既有行為，刻意保留）。
        # 但改了要留得下來：報表要能篩出「被人工覆寫」的單。
        # 放在 conn.close() 之後：_effective_rate_for() 會自己開連線並可能寫入凍結值，
        # 與上面那條寫入交疊會讓 SQLite 卡在 write lock。
        if billed_weight:
            rate_override = _log_rate_override(req_id, g_code_val, rate_per_kg, now)
        # 這張單改變了該月累計 → 快照要跟上，否則下月費率會用到舊的 total_kg
        _resync_snapshot(g_code_val, now[:7])
    return jsonify({"success": True, "g_code": g_code_val,
                    "customer_name": customer_name_val,
                    "rate_override": rate_override})


def _log_rate_override(req_id, g_code, actual_rate, when):
    """比對「階梯應有費率」與「實際套用費率」，不同就記一筆操作紀錄。

    回 dict（有覆寫）或 None（沒覆寫 / 無法判定）。
    刻意不擋、不改值 —— 管理員有權覆寫，這裡只負責留痕。

    兩種情況不比對，因為那時「階梯應有費率」這個概念不存在：
      ・方案B（flat）模式
      ・該 g_code 不是團主（散客照原費率出貨，比對只會產生假的覆寫紀錄）"""
    try:
        actual = float(actual_rate or 0)
        if actual <= 0 or not g_code:
            return None
        info = _effective_rate_for(g_code, (when or "")[:7] or _current_ym())
        if info.get("mode") != "tier" or not info.get("is_tenant"):
            return None
        expected = float(info.get("rate") or 0)
        if expected <= 0 or abs(expected - actual) < 1e-6:
            return None
        detail = (f"出貨單#{req_id} 階梯應為 NT${expected:g}/kg"
                  f"（依上月 {info.get('basis_kg', 0):g}kg），實際套用 NT${actual:g}/kg")
        log_op("費率人工覆寫", g_code, detail)
        return {"expected_rate": expected, "actual_rate": actual,
                "basis_kg": info.get("basis_kg", 0), "ym": info.get("ym", "")}
    except Exception as e:
        print(f"[rate_override] 出貨單#{req_id} 比對失敗: {e}", flush=True)
        return None


@app.route("/api/admin/shipment_requests/<int:req_id>/revert", methods=["POST"])
def admin_revert_shipment_request(req_id):
    """還原出貨申請：狀態回到待處理，包裹回到已到貨，清空帳單"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    req = conn.execute("SELECT * FROM shipment_requests WHERE id=?", (req_id,)).fetchone()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """UPDATE shipment_requests 
           SET status='待處理', updated_at=?,
               billed_weight=0, rate_per_kg=0, shipping_fee=0, handling_fee=0,
               consolidation_fee=0, letter_fee=0, total_fee=0,
               tracking_num='', payment_last5='', payment_at='', extra_services=''
           WHERE id=?""",
        (now, req_id)
    )

    # 包裹狀態還原為「已到貨」
    pkg_ids_str = req["package_ids"]
    if pkg_ids_str:
        pkg_ids = [int(x.strip()) for x in pkg_ids_str.split(",") if x.strip()]
        if pkg_ids:
            placeholders = ",".join(["?"] * len(pkg_ids))
            conn.execute(
                f"UPDATE packages SET status='已到貨' WHERE id IN ({placeholders})", pkg_ids
            )

    conn.commit()
    conn.close()
    # 還原會把這張單從原本歸屬的月份拿掉（狀態變待處理，且 updated_at 被改成今天）。
    # 要重算「原本那個月」，不是只算今天這個月 —— 否則舊月快照會一直留著這張單的重量，
    # 下個月的階梯就踩在不存在的貨量上。
    _resync_snapshot(req["g_code"], (req["updated_at"] or req["created_at"] or "")[:7], now[:7])
    return jsonify({"success": True})


# ============ 預報包裹 API（本地存檔，不連 JPD）============

@app.route("/api/forecast_simple", methods=["POST"])
def create_forecast_simple():
    """客戶提交預報（存到本地 DB）"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    customer_name = data.get("customer_name", "")
    items = data.get("items", [])
    note = (data.get("note") or "").strip()

    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    if not items:
        return jsonify({"success": False, "error": "請至少填寫一個商品"})

    # 過濾空的
    valid_items = [i for i in items if (i.get("name") or "").strip()]
    if not valid_items:
        return jsonify({"success": False, "error": "請至少填寫一個商品名稱"})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fc_agent_id = get_agent_id_for_g_code(g_code)
    conn = get_db()
    conn.execute(
        """INSERT INTO forecasts (g_code, customer_name, items_json, status, note, created_at, agent_id)
           VALUES (?, ?, ?, '待處理', ?, ?, ?)""",
        (g_code, customer_name, json.dumps(valid_items, ensure_ascii=False), note, now, fc_agent_id)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "預報已送出！我們收到後會盡快處理。"})


@app.route("/api/my_forecasts", methods=["GET"])
def get_my_forecasts():
    """客戶查看自己的預報"""
    g_code = request.args.get("g_code", "").upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    ok, resp = _require_customer(g_code)
    if not ok:
        return resp
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM forecasts WHERE g_code=? ORDER BY id DESC LIMIT 20", (g_code,)
    ).fetchall()
    conn.close()
    results = []
    for r in rows:
        row = dict(r)
        try:
            row["items"] = json.loads(row.get("items_json") or "[]")
        except:
            row["items"] = []
        results.append(row)
    return jsonify({"success": True, "forecasts": results})


@app.route("/api/admin/forecasts", methods=["GET"])
def admin_get_forecasts():
    """管理員查看所有預報"""
    status = request.args.get("status", "")
    g_code = request.args.get("g_code", "").upper()
    aid = get_current_agent_id()
    af = " AND agent_id=?" if aid > 0 else ""
    aparams = (aid,) if aid > 0 else ()
    conn = get_db()
    if g_code and status:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE g_code=? AND status=?" + af + " ORDER BY id DESC",
            (g_code, status) + aparams
        ).fetchall()
    elif g_code:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE g_code=?" + af + " ORDER BY id DESC", (g_code,) + aparams
        ).fetchall()
    elif status:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE status=?" + af + " ORDER BY id DESC", (status,) + aparams
        ).fetchall()
    else:
        if aid > 0:
            rows = conn.execute("SELECT * FROM forecasts WHERE agent_id=? ORDER BY id DESC", (aid,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM forecasts ORDER BY id DESC").fetchall()
    conn.close()
    results = []
    for r in rows:
        row = dict(r)
        try:
            row["items"] = json.loads(row.get("items_json") or "[]")
        except:
            row["items"] = []
        results.append(row)
    return jsonify({"success": True, "forecasts": results})


@app.route("/api/admin/forecasts/<int:fc_id>", methods=["PUT"])
def admin_update_forecast(fc_id):
    """管理員更新預報狀態"""
    ok, _row = check_record_ownership("forecasts", fc_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該預報"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    status = data.get("status", "")
    conn = get_db()
    conn.execute("UPDATE forecasts SET status=? WHERE id=?", (status, fc_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/forecasts/<int:fc_id>/excel")
def admin_download_forecast_excel(fc_id):
    """下載單筆預報的 JPD Excel"""
    ok, _row = check_record_ownership("forecasts", fc_id)
    if not _row:
        return "Not found", 404
    if not ok:
        return "Forbidden", 403
    conn = get_db()
    row = conn.execute("SELECT * FROM forecasts WHERE id=?", (fc_id,)).fetchone()
    conn.close()
    if not row:
        return "Not found", 404
    row = dict(row)
    try:
        items = json.loads(row.get("items_json") or "[]")
    except:
        items = []

    g_code = row["g_code"]
    today_str = datetime.now().strftime("%m%d")
    customer_order_id = f"{g_code}-{today_str}"

    wb = Workbook()
    ws = wb.active
    ws.title = "預報資料"

    # 標頭
    headers = [
        "客戶運單號", "JpD包裹ID", "運單ID", "包裹特殊服務",
        "收件人", "收件人身份證ID", "收件人詳細地址", "收件人电话号码",
        "備註", "特殊服务", "渠道ID",
        "申報人", "申報人身份證ID", "申報人詳細地址", "申報人电话号码",
        "品名", "数量", "金额", "材質", "產地", "URL/JanCode"
    ]
    hfill = PatternFill("solid", fgColor="1F4E79")
    hfont = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin

    # 資料
    for row_idx, item in enumerate(items, 2):
        data_row = [
            customer_order_id, "", "", "",
            "", "", "", "",
            row.get("note", ""), "", "40",
            "", "", "", "",
            item.get("name", ""),
            item.get("quantity", 1),
            item.get("price", 0),
            "", "Japan",
            item.get("url", "")
        ]
        for col_idx, val in enumerate(data_row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.border = thin
            cell.alignment = Alignment(vertical="center")

    # 欄寬
    col_widths = {1:16, 2:14, 5:12, 7:20, 8:16, 9:12, 11:8, 16:20, 17:8, 18:10, 21:30}
    for col, w in col_widths.items():
        ws.column_dimensions[chr(64+col) if col<=26 else 'A'].width = w
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"{g_code}_{today_str}_forecast.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║       客人集運預報系統                                      ║
    ║       御用達 × JPD 雲倉                                     ║
    ╚═══════════════════════════════════════════════════════════╝
    🌐 服務啟動於 Port: {port}
    💱 TWD → JPY 匯率: {TWD_TO_JPY_RATE}
    """)
    app.run(host="0.0.0.0", port=port, debug=debug)

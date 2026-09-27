#!/usr/bin/env python3
"""
一次性工具：跑 Shopify OAuth 授權碼流程，換到 Admin API 的 offline access token（shpat_ 開頭、不會自己過期）。

★ 本機執行一次即可，伺服器不需要。只用標準函式庫，不加進 requirements.txt。
★ token 只印在終端機，不寫進任何檔案、不寫 log。client_secret 只從環境變數或互動輸入取得。

背景：2026/01/01 起 Shopify 停用後台「Develop apps」的舊版自訂應用程式流程；Dev Dashboard 建的 app
不會在任何介面顯示永久 token，必須跑一次 OAuth 授權碼流程換 offline token。
（不要用 client credentials 流程——那種 token 24 小時就過期。）

事前準備（Dev Dashboard）：
  • app 已建好並安裝到商店
  • Redirect URL 已登記：http://localhost:3456/callback（必須與本腳本送出的完全一致，含 http、port、路徑）

執行（Windows PowerShell）：
  $env:SHOPIFY_SHOP = "xxx.myshopify.com"
  $env:SHOPIFY_CLIENT_ID = "..."
  $env:SHOPIFY_CLIENT_SECRET = "..."          # 或不設，執行時會用 getpass 問（不回顯）
  $env:SHOPIFY_SCOPES = "read_orders,read_products,write_products"
  python tools\get_shopify_token.py
沒設的會用 input() 互動詢問。也可用參數：--shop / --client-id / --scopes（secret 不接受參數，避免留在命令歷史）。

流程：本機起 http://127.0.0.1:3456/callback → 自動開瀏覽器到 Shopify 授權頁 → 按「安裝」→ 導回本機
→ 驗 state（防 CSRF）→ 驗 HMAC（用 client_secret）→ POST /admin/oauth/access_token → 印出 token 與 scope。
★ scope 清單是整份覆蓋：漏列即失去該權限。
"""

import argparse
import getpass
import hashlib
import html as html_mod
import hmac
import json
import os
import secrets
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, urlencode, urlparse
from urllib.request import Request, urlopen

DEFAULT_PORT = 3456
CALLBACK_PATH = "/callback"
WAIT_TIMEOUT_SEC = 600      # 等瀏覽器授權最多 10 分鐘
# 測試用：指到假的 Shopify server（例如 http://127.0.0.1:port）。正式使用不用設。
OAUTH_BASE_OVERRIDE = os.environ.get("SHOPIFY_OAUTH_BASE", "").rstrip("/")


class ConfigError(Exception):
    """參數缺漏或格式錯（印訊息、不印 traceback）。"""


class AuthError(Exception):
    """授權流程中止（state / hmac / 換 token 失敗）。"""


# ── 純函式（可單測）──

def normalize_shop(shop):
    s = (shop or "").strip().lower()
    s = s.replace("https://", "").replace("http://", "").rstrip("/")
    if not s:
        raise ConfigError("缺少 SHOP（商店網域，例如 xxx.myshopify.com）")
    if "." not in s:
        s += ".myshopify.com"
    if not s.endswith(".myshopify.com"):
        raise ConfigError(f"SHOP 格式錯：{s}（必須是 xxx.myshopify.com，不是自訂網域）")
    return s


def normalize_scopes(scopes):
    parts = [p.strip() for p in (scopes or "").replace(" ", ",").split(",") if p.strip()]
    if not parts:
        raise ConfigError("缺少 SCOPES（例如 read_orders,read_products；整份覆蓋、漏列即失去該權限）")
    return ",".join(dict.fromkeys(parts))


def oauth_base(shop):
    return OAUTH_BASE_OVERRIDE or f"https://{shop}"


def redirect_uri(port=DEFAULT_PORT):
    return f"http://localhost:{port}{CALLBACK_PATH}"


def build_auth_url(shop, client_id, scopes, state, port=DEFAULT_PORT):
    q = urlencode({
        "client_id": client_id,
        "scope": scopes,
        "redirect_uri": redirect_uri(port),
        "state": state,
    })
    return f"{oauth_base(shop)}/admin/oauth/authorize?{q}"


def verify_hmac(query_params, client_secret):
    """Shopify 規則：去掉 hmac（與舊的 signature），其餘依 key 排序後以 k=v&k=v 串接，
    HMAC-SHA256(client_secret) 的 hex 必須等於 query 裡的 hmac。"""
    given = (query_params.get("hmac") or "").strip()
    if not given:
        return False
    items = sorted((k, v) for k, v in query_params.items() if k not in ("hmac", "signature"))
    message = "&".join(f"{k}={v}" for k, v in items)
    digest = hmac.new(client_secret.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).hexdigest()
    return hmac.compare_digest(digest, given.lower())


def validate_callback(query_params, expected_state, client_secret):
    """回 code；state 或 hmac 不符就拋 AuthError（此時絕不去換 token）。"""
    if query_params.get("error"):
        raise AuthError(f"Shopify 回錯誤：{query_params.get('error')} {query_params.get('error_description', '')}".strip())
    got_state = query_params.get("state") or ""
    if not hmac.compare_digest(got_state, expected_state):
        raise AuthError("state 不符（可能是 CSRF 或不是這次流程的回呼），中止；未發出 token 交換請求")
    if not verify_hmac(query_params, client_secret):
        raise AuthError("HMAC 驗證失敗（回呼不是 Shopify 用這個 client_secret 簽的，或 client_secret 打錯），中止；未發出 token 交換請求")
    code = query_params.get("code") or ""
    if not code:
        raise AuthError("回呼缺少 code，中止")
    return code


def exchange_token(shop, client_id, client_secret, code):
    """POST /admin/oauth/access_token → {access_token, scope}。"""
    url = f"{oauth_base(shop)}/admin/oauth/access_token"
    body = json.dumps({"client_id": client_id, "client_secret": client_secret, "code": code}).encode("utf-8")
    req = Request(url, data=body, method="POST",
                  headers={"Content-Type": "application/json", "Accept": "application/json"})
    try:
        with urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", "replace")
    except HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300] if e.fp else ""
        raise AuthError(f"換 token 失敗 HTTP {e.code}：{detail}（code 只能用一次、幾分鐘內有效；重跑一次即可）")
    except URLError as e:
        raise AuthError(f"換 token 失敗（連線）：{e.reason}")
    try:
        data = json.loads(raw)
        token = data["access_token"]
    except (ValueError, KeyError, TypeError):
        raise AuthError(f"換 token 回應格式錯：{raw[:300]}")
    return {"access_token": token, "scope": data.get("scope", "")}


# ── 本機回呼 server ──

HTML_OK = """<!doctype html><html><head><meta charset="utf-8"><title>授權完成</title></head>
<body style="font-family:system-ui;padding:40px;"><h2>✅ 授權完成，可以關閉此分頁</h2>
<p>token 已印在終端機。</p></body></html>"""
HTML_FAIL = """<!doctype html><html><head><meta charset="utf-8"><title>授權失敗</title></head>
<body style="font-family:system-ui;padding:40px;"><h2>❌ 授權中止</h2><p>{msg}</p><p>請看終端機。</p></body></html>"""


class CallbackServer(HTTPServer):
    """收一次 /callback 就結束。結果放 self.result（{"code": ...} 或 {"error": ...}）。"""
    # POSIX：允許重跑時馬上重用 port（TIME_WAIT）。Windows：SO_REUSEADDR 會讓兩個程式綁同一個 port
    # 都成功、回呼落到別人手上，所以關掉。
    allow_reuse_address = os.name != "nt"

    def __init__(self, addr, expected_state, client_secret):
        super().__init__(addr, CallbackHandler)
        self.expected_state = expected_state
        self.client_secret = client_secret
        self.result = None
        self.done = threading.Event()


class CallbackHandler(BaseHTTPRequestHandler):
    def log_message(self, *a):        # 不印 access log（query 裡有 code）
        pass

    def _html(self, code, html):
        body = html.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path != CALLBACK_PATH:
            return self._html(404, "<h2>404</h2>")
        if self.server.done.is_set():
            return self._html(410, HTML_FAIL.format(msg="這次流程已結束，請重跑腳本。"))
        params = dict(parse_qsl(u.query, keep_blank_values=True))
        try:
            code = validate_callback(params, self.server.expected_state, self.server.client_secret)
            self.server.result = {"code": code, "shop": params.get("shop", "")}
            self._html(200, HTML_OK)
        except AuthError as e:
            self.server.result = {"error": str(e)}
            self._html(400, HTML_FAIL.format(msg=html_mod.escape(str(e))))
        finally:
            self.server.done.set()


def wait_for_callback(server, timeout=WAIT_TIMEOUT_SEC):
    """處理請求直到收到 /callback（favicon 之類的雜訊會被忽略）或逾時。"""
    server.timeout = 1
    deadline = time.time() + timeout
    while not server.done.is_set():
        if time.time() > deadline:
            raise AuthError(f"等待瀏覽器授權逾時（{timeout} 秒），中止")
        server.handle_request()
    return server.result


# ── 參數 ──

def _env(*names):
    for n in names:
        v = os.environ.get(n)
        if v and v.strip():
            return v.strip()
    return ""


def _ask(prompt, secret=False):
    try:
        v = getpass.getpass(prompt) if secret else input(prompt)
    except (EOFError, KeyboardInterrupt):
        raise ConfigError(f"未提供：{prompt.strip().rstrip('：:')}")
    v = (v or "").strip()
    if not v:
        raise ConfigError(f"未提供：{prompt.strip().rstrip('：:')}")
    return v


def load_config(argv=None, interactive=True):
    ap = argparse.ArgumentParser(description="Shopify OAuth 取 offline access token（本機跑一次）", add_help=True)
    ap.add_argument("--shop", help="xxx.myshopify.com（或環境變數 SHOPIFY_SHOP / SHOP）")
    ap.add_argument("--client-id", help="Dev Dashboard 的 Client ID（或 SHOPIFY_CLIENT_ID / CLIENT_ID）")
    ap.add_argument("--scopes", help="逗號分隔，整份覆蓋（或 SHOPIFY_SCOPES / SCOPES）")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"本機回呼 port（預設 {DEFAULT_PORT}，須與登記的 Redirect URL 一致）")
    ap.add_argument("--no-browser", action="store_true", help="不自動開瀏覽器，只印授權網址")
    ap.add_argument("--timeout", type=int, default=WAIT_TIMEOUT_SEC, help="等待授權秒數")
    ap.add_argument("--non-interactive", action="store_true", help="缺參數時不互動詢問、直接報錯（腳本化用）")
    a = ap.parse_args(argv)
    interactive = interactive and not a.non_interactive

    shop = a.shop or _env("SHOPIFY_SHOP", "SHOP")
    client_id = a.client_id or _env("SHOPIFY_CLIENT_ID", "CLIENT_ID")
    client_secret = _env("SHOPIFY_CLIENT_SECRET", "CLIENT_SECRET")     # 只從環境變數或互動輸入
    scopes = a.scopes or _env("SHOPIFY_SCOPES", "SCOPES")

    if interactive and sys.stdin and sys.stdin.isatty():
        # 互動詢問；輸入空白或 EOF（Windows 上 stdin 接 NUL 時 isatty 也會回 True）就當沒提供，
        # 交給下面統一列出缺哪些。
        try:
            if not shop:
                shop = _ask("商店網域（xxx.myshopify.com）：")
            if not client_id:
                client_id = _ask("Client ID：")
            if not client_secret:
                client_secret = _ask("Client Secret（不回顯）：", secret=True)
            if not scopes:
                scopes = _ask("Scopes（逗號分隔，整份覆蓋）：")
        except ConfigError:
            pass

    missing = [n for n, v in (("SHOP", shop), ("CLIENT_ID", client_id),
                              ("CLIENT_SECRET", client_secret), ("SCOPES", scopes)) if not v]
    if missing:
        raise ConfigError("缺少必要參數：" + "、".join(missing)
                          + "。請設環境變數 SHOPIFY_SHOP / SHOPIFY_CLIENT_ID / SHOPIFY_CLIENT_SECRET / SHOPIFY_SCOPES，"
                            "或在互動終端機執行讓腳本詢問。")
    return {
        "shop": normalize_shop(shop),
        "client_id": client_id,
        "client_secret": client_secret,
        "scopes": normalize_scopes(scopes),
        "port": a.port,
        "open_browser": not a.no_browser,
        "timeout": a.timeout,
    }


# ── 主流程 ──

def run_flow(cfg, out=None):
    """起 server → 開瀏覽器 → 等回呼 → 驗證 → 換 token。回 {access_token, scope, shop}。失敗拋 AuthError。"""
    out = out or sys.stdout
    state = secrets.token_urlsafe(32)
    try:
        server = CallbackServer(("127.0.0.1", cfg["port"]), state, cfg["client_secret"])
    except OSError as e:
        raise AuthError(f"無法監聽 127.0.0.1:{cfg['port']}（{e}）。port 被占用？關掉占用的程式或改 --port（Redirect URL 也要跟著改）")
    try:
        url = build_auth_url(cfg["shop"], cfg["client_id"], cfg["scopes"], state, cfg["port"])
        print(f"redirect_uri = {redirect_uri(cfg['port'])}（必須與 Dev Dashboard 登記的完全一致）", file=out)
        print("授權網址：", file=out)
        print(url, file=out)
        if cfg.get("open_browser", True):
            try:
                webbrowser.open(url, new=2)
                print("已開啟瀏覽器；若沒有自動開啟，請手動複製上面網址。", file=out)
            except Exception as e:      # noqa: BLE001
                print(f"開瀏覽器失敗（{e}），請手動複製上面網址。", file=out)
        print("等待你在瀏覽器按「安裝」…", file=out)
        result = wait_for_callback(server, cfg.get("timeout", WAIT_TIMEOUT_SEC))
        if not result or result.get("error"):
            raise AuthError((result or {}).get("error") or "未收到回呼")
        print("回呼 state / HMAC 驗證通過，換 token 中…", file=out)
        tok = exchange_token(cfg["shop"], cfg["client_id"], cfg["client_secret"], result["code"])
        tok["shop"] = cfg["shop"]
        return tok
    finally:
        server.server_close()


def main(argv=None):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:       # noqa: BLE001
        pass
    try:
        cfg = load_config(argv)
        tok = run_flow(cfg)
    except ConfigError as e:
        print(f"❌ {e}", file=sys.stderr)
        return 2
    except AuthError as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("已取消", file=sys.stderr)
        return 130
    print()
    print("=" * 64)
    print(f"商店：{tok['shop']}")
    print(f"實際取得的 scope：{tok.get('scope') or '（Shopify 未回傳）'}")
    print("SHOPIFY_ACCESS_TOKEN =")
    print(tok["access_token"])
    print("=" * 64)
    print("這個 token 不會過期（除非 app 被移除或重新安裝）。貼進 Zeabur 環境變數 SHOPIFY_ACCESS_TOKEN。")
    print("本腳本沒有把它存在任何地方；貼完請關掉這個終端機視窗（或清掉捲動紀錄）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

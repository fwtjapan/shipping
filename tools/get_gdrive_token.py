#!/usr/bin/env python3
"""
一次性工具：跑 Google OAuth 授權碼流程，換到 Google Drive 的 refresh token（給 SQLite 自動備份用）。

★ 本機執行一次即可，伺服器不需要。只用標準函式庫，不加進 requirements.txt。
★ token 只印在終端機，不寫進任何檔案、不寫 log。client_secret 只從環境變數或互動輸入取得
  （不接受命令列參數，避免留在命令歷史）。

★ 不要用 service account：Drive 的 service account 沒有儲存空間配額，上傳必回 403 storageQuotaExceeded
  （分享資料夾給它也一樣，Google 平台限制）。已查證。

事前準備（Google Cloud Console，用要放備份的那個 Gmail；專案 goyoutati-backup 已建好）：
  • 已啟用 Google Drive API
  • OAuth 同意畫面已「發布」為實際運作中（測試模式的 refresh token 7 天就失效）
  • OAuth 用戶端類型＝「桌面應用程式」（桌面型用戶端 Google 允許任何 http://localhost:<port> 回呼，不必登記）

執行（Windows PowerShell）：
  $env:GDRIVE_CLIENT_ID = "xxxx.apps.googleusercontent.com"
  $env:GDRIVE_CLIENT_SECRET = "GOCSPX-xxxx"          # 或不設，執行時會用 getpass 問（不回顯）
  python tools\get_gdrive_token.py
沒設的會用 input() 互動詢問。也可用參數：--client-id（secret 不接受參數）。

流程：本機起 http://127.0.0.1:3457/callback（用 3457，避開 tools/get_shopify_token.py 的 3456，兩支同時開也不撞）
→ 自動開瀏覽器到 Google 授權頁 → 選帳號、按「允許」→ 導回本機 → 驗 state（防 CSRF）
→ POST https://oauth2.googleapis.com/token 換 refresh token
→ ★ 立刻拿 refresh token 換一次 access token 打 GET /drive/v3/about?fields=user，印出帳號 email，
  讓你確認授權到的是正確的 Google 帳號（避免辛苦設定完才發現授權到別的帳號）。

授權網址兩個參數缺一不可：access_type=offline、prompt=consent。缺了 Google 只回 access token、不給 refresh token
（同一帳號第二次授權若不帶 prompt=consent，Google 也不會再給）。
scope 只用 drive.file：程式只能存取自己建立的檔案（helpshipping-backups 資料夾與裡面的備份），
token 萬一外洩，波及範圍限於備份檔，不是整個雲端硬碟。不要改成完整的 drive scope。
"""

import argparse
import getpass
import hmac
import html as html_mod
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

DEFAULT_PORT = 3457
CALLBACK_PATH = "/callback"
SCOPE = "https://www.googleapis.com/auth/drive.file"
WAIT_TIMEOUT_SEC = 600      # 等瀏覽器授權最多 10 分鐘
# 測試用：三個端點都可指到假的 Google server（例如 http://127.0.0.1:port/...）。正式使用不用設。
# TOKEN_URL / API_BASE 的環境變數名稱與 gdrive_backup.py 相同。
AUTH_URL = os.environ.get("GDRIVE_AUTH_URL", "https://accounts.google.com/o/oauth2/v2/auth")
TOKEN_URL = os.environ.get("GDRIVE_TOKEN_URL", "https://oauth2.googleapis.com/token")
API_BASE = os.environ.get("GDRIVE_API_BASE", "https://www.googleapis.com").rstrip("/")


class ConfigError(Exception):
    """參數缺漏或格式錯（印訊息、不印 traceback）。"""


class AuthError(Exception):
    """授權流程中止（state / 換 token / 驗證失敗）。"""


# ── 純函式（可單測）──

def normalize_client_id(client_id):
    s = (client_id or "").strip()
    if not s:
        raise ConfigError("缺少 CLIENT_ID")
    if not s.endswith(".apps.googleusercontent.com"):
        raise ConfigError(f"CLIENT_ID 格式錯：{s}（應以 .apps.googleusercontent.com 結尾，請到 Cloud Console 憑證頁複製）")
    return s


def redirect_uri(port=DEFAULT_PORT):
    return f"http://localhost:{port}{CALLBACK_PATH}"


def build_auth_url(client_id, state, port=DEFAULT_PORT):
    q = urlencode({
        "client_id": client_id,
        "redirect_uri": redirect_uri(port),
        "response_type": "code",
        "scope": SCOPE,
        "access_type": "offline",       # 缺了不給 refresh token
        "prompt": "consent",            # 缺了同帳號第二次授權不給 refresh token
        "state": state,
    })
    return f"{AUTH_URL}?{q}"


def validate_callback(query_params, expected_state):
    """回 code；state 不符就拋 AuthError（此時絕不去換 token）。"""
    if query_params.get("error"):
        raise AuthError(f"Google 回錯誤：{query_params.get('error')} {query_params.get('error_description', '')}".strip())
    got_state = query_params.get("state") or ""
    if not hmac.compare_digest(got_state, expected_state):
        raise AuthError("state 不符（可能是 CSRF 或不是這次流程的回呼），中止；未發出 token 交換請求")
    code = query_params.get("code") or ""
    if not code:
        raise AuthError("回呼缺少 code，中止")
    return code


def _post_token(form, what):
    """POST TOKEN_URL（x-www-form-urlencoded）→ dict。HTTP / 連線 / 格式錯都轉成 AuthError。"""
    body = urlencode(form).encode("utf-8")
    req = Request(TOKEN_URL, data=body, method="POST",
                  headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"})
    try:
        with urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", "replace")
    except HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300] if e.fp else ""
        hint = ""
        if e.code in (400, 401):
            hint = "（常見原因：client_secret 打錯、code 已用過或過期、redirect_uri 不一致；重跑一次即可）"
        raise AuthError(f"{what}失敗 HTTP {e.code}：{detail}{hint}")
    except URLError as e:
        raise AuthError(f"{what}失敗（連線）：{e.reason}")
    try:
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise TypeError
    except (ValueError, TypeError):
        raise AuthError(f"{what}回應格式錯：{raw[:300]}")
    return data


def exchange_token(client_id, client_secret, code, port=DEFAULT_PORT):
    """POST token endpoint（grant_type=authorization_code）→ {refresh_token, access_token, scope}。
    缺 refresh_token 時給明確錯誤訊息與可能原因，不是 traceback。"""
    data = _post_token({
        "client_id": client_id,
        "client_secret": client_secret,
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": redirect_uri(port),
    }, "換 token ")
    refresh = data.get("refresh_token")
    if not refresh:
        raise AuthError(
            "Google 回應裡沒有 refresh_token（只給了 access token）。可能原因：\n"
            "  1. 授權網址缺 access_type=offline 或 prompt=consent（本腳本兩個都有帶；若你手動改過網址請確認）\n"
            "  2. 這個帳號先前已授權過此 app 且 Google 不再重發：到 https://myaccount.google.com/permissions\n"
            "     移除這個 app 的存取權後再跑一次\n"
            f"  Google 回應欄位：{', '.join(sorted(data.keys())) or '（空）'}")
    return {"refresh_token": refresh, "access_token": data.get("access_token", ""), "scope": data.get("scope", "")}


def verify_refresh_token(client_id, client_secret, refresh_token):
    """用 refresh token 換 access token → GET /drive/v3/about?fields=user → 回帳號 email。
    這一步證明 token 真的能用，且讓使用者確認授權到的是哪個 Google 帳號。"""
    data = _post_token({
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }, "用 refresh token 換 access token ")
    access = data.get("access_token")
    if not access:
        raise AuthError(f"用 refresh token 換 access token 的回應缺 access_token：{', '.join(sorted(data.keys()))}")
    req = Request(f"{API_BASE}/drive/v3/about?fields=user", headers={"Authorization": f"Bearer {access}"})
    try:
        with urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", "replace")
    except HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300] if e.fp else ""
        hint = "（403 多半是 Cloud Console 沒啟用 Google Drive API）" if e.code == 403 else ""
        raise AuthError(f"驗證 token 失敗 HTTP {e.code}：{detail}{hint}")
    except URLError as e:
        raise AuthError(f"驗證 token 失敗（連線）：{e.reason}")
    try:
        user = json.loads(raw).get("user") or {}
    except (ValueError, AttributeError):
        raise AuthError(f"驗證 token 回應格式錯：{raw[:300]}")
    email = (user.get("emailAddress") or "").strip()
    if not email:
        raise AuthError(f"驗證 token 回應缺 user.emailAddress：{raw[:300]}")
    return email


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

    def __init__(self, addr, expected_state):
        super().__init__(addr, CallbackHandler)
        self.expected_state = expected_state
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
            code = validate_callback(params, self.server.expected_state)
            self.server.result = {"code": code}
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
    ap = argparse.ArgumentParser(description="Google Drive OAuth 取 refresh token（本機跑一次）", add_help=True)
    ap.add_argument("--client-id", help="Cloud Console 的 OAuth 用戶端 ID（或環境變數 GDRIVE_CLIENT_ID / CLIENT_ID）")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"本機回呼 port（預設 {DEFAULT_PORT}）")
    ap.add_argument("--no-browser", action="store_true", help="不自動開瀏覽器，只印授權網址")
    ap.add_argument("--timeout", type=int, default=WAIT_TIMEOUT_SEC, help="等待授權秒數")
    ap.add_argument("--non-interactive", action="store_true", help="缺參數時不互動詢問、直接報錯（腳本化用）")
    a = ap.parse_args(argv)
    interactive = interactive and not a.non_interactive

    client_id = a.client_id or _env("GDRIVE_CLIENT_ID", "CLIENT_ID")
    client_secret = _env("GDRIVE_CLIENT_SECRET", "CLIENT_SECRET")     # 只從環境變數或互動輸入

    if interactive and sys.stdin and sys.stdin.isatty():
        # 互動詢問；輸入空白或 EOF（Windows 上 stdin 接 NUL 時 isatty 也會回 True）就當沒提供，
        # 交給下面統一列出缺哪些。
        try:
            if not client_id:
                client_id = _ask("Client ID（xxxx.apps.googleusercontent.com）：")
            if not client_secret:
                client_secret = _ask("Client Secret（不回顯）：", secret=True)
        except ConfigError:
            pass

    missing = [n for n, v in (("CLIENT_ID", client_id), ("CLIENT_SECRET", client_secret)) if not v]
    if missing:
        raise ConfigError("缺少必要參數：" + "、".join(missing)
                          + "。請設環境變數 GDRIVE_CLIENT_ID / GDRIVE_CLIENT_SECRET，"
                            "或在互動終端機執行讓腳本詢問。")
    return {
        "client_id": normalize_client_id(client_id),
        "client_secret": client_secret,
        "port": a.port,
        "open_browser": not a.no_browser,
        "timeout": a.timeout,
    }


# ── 主流程 ──

def run_flow(cfg, out=None):
    """起 server → 開瀏覽器 → 等回呼 → 驗 state → 換 token → 驗證 token 並取 email。
    回 {refresh_token, scope, email}。失敗拋 AuthError。"""
    out = out or sys.stdout
    state = secrets.token_urlsafe(32)
    try:
        server = CallbackServer(("127.0.0.1", cfg["port"]), state)
    except OSError as e:
        raise AuthError(f"無法監聽 127.0.0.1:{cfg['port']}（{e}）。port 被占用？關掉占用的程式或改 --port")
    try:
        url = build_auth_url(cfg["client_id"], state, cfg["port"])
        print(f"redirect_uri = {redirect_uri(cfg['port'])}", file=out)
        print("授權網址：", file=out)
        print(url, file=out)
        if cfg.get("open_browser", True):
            try:
                webbrowser.open(url, new=2)
                print("已開啟瀏覽器；若沒有自動開啟，請手動複製上面網址。", file=out)
            except Exception as e:      # noqa: BLE001
                print(f"開瀏覽器失敗（{e}），請手動複製上面網址。", file=out)
        print("等待你在瀏覽器選帳號（要放備份的那個 Gmail）並按「允許」…", file=out)
        result = wait_for_callback(server, cfg.get("timeout", WAIT_TIMEOUT_SEC))
        if not result or result.get("error"):
            raise AuthError((result or {}).get("error") or "未收到回呼")
        print("回呼 state 驗證通過，換 token 中…", file=out)
        tok = exchange_token(cfg["client_id"], cfg["client_secret"], result["code"], cfg["port"])
        print("已取得 refresh token，驗證中（換 access token → 查 Drive 帳號）…", file=out)
        tok["email"] = verify_refresh_token(cfg["client_id"], cfg["client_secret"], tok["refresh_token"])
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
    print(f"授權到的 Google 帳號：{tok['email']}   ← 請確認是要放備份的那個帳號")
    print(f"實際取得的 scope：{tok.get('scope') or '（Google 未回傳）'}")
    print("GDRIVE_REFRESH_TOKEN =")
    print(tok["refresh_token"])
    print("=" * 64)
    print("帳號不對就到 https://myaccount.google.com/permissions 移除此 app 存取權，重跑並選正確帳號。")
    print("帳號正確就把上面這串貼進 Zeabur 環境變數 GDRIVE_REFRESH_TOKEN（連同 GDRIVE_CLIENT_ID / GDRIVE_CLIENT_SECRET），redeploy。")
    print("同意畫面已發布為正式版，這個 token 不會 7 天過期（除非你在 Google 帳號移除存取權、或半年沒用）。")
    print("本腳本沒有把它存在任何地方；貼完請關掉這個終端機視窗（或清掉捲動紀錄）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

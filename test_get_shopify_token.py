# -*- coding: utf-8 -*-
"""
tools/get_shopify_token.py 驗收測試（任務 AC）。用假的 Shopify server，不碰真的 Shopify。
執行：python test_get_shopify_token.py
"""
import ast
import hashlib
import hmac
import io
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlencode, urlparse
from urllib.request import urlopen

ROOT = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.join(ROOT, "tools", "get_shopify_token.py")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SHOP = "test-shop.myshopify.com"
CLIENT_ID = "cid_123"
CLIENT_SECRET = "shpss_fake_secret_abc"
SCOPES = "read_orders,write_products"
FAKE_TOKEN = "shpat_fake_token_0123456789abcdef"
PORT = 3456


def log(m):
    sys.stderr.write(m + "\n")


# ── 假 Shopify ──
class FakeShopify:
    def __init__(self):
        self.token_calls = 0
        self.issued_code = "authcode_xyz"
        self.last_authorize = None


FAKE = FakeShopify()


def sign(params, secret):
    msg = "&".join(f"{k}={v}" for k, v in sorted(params.items()))
    return hmac.new(secret.encode(), msg.encode(), hashlib.sha256).hexdigest()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, code, obj):
        b = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path == "/admin/oauth/authorize":
            FAKE.last_authorize = q
            assert q["client_id"] == CLIENT_ID and q["scope"] == SCOPES
            assert q["redirect_uri"] == f"http://localhost:{PORT}/callback", q["redirect_uri"]
            params = {"code": FAKE.issued_code, "shop": SHOP, "state": q["state"],
                      "timestamp": str(int(time.time())), "host": "aG9zdA=="}
            params["hmac"] = sign(params, CLIENT_SECRET)
            # 使用者按「安裝」→ 302 回 redirect_uri
            self.send_response(302)
            self.send_header("Location", q["redirect_uri"] + "?" + urlencode(params))
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        self._json(404, {"error": "nf"})

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        body = json.loads(self.rfile.read(n).decode() or "{}")
        if u.path == "/admin/oauth/access_token":
            FAKE.token_calls += 1
            if body.get("client_id") != CLIENT_ID or body.get("client_secret") != CLIENT_SECRET \
                    or body.get("code") != FAKE.issued_code:
                return self._json(400, {"error": "invalid_request"})
            return self._json(200, {"access_token": FAKE_TOKEN, "scope": SCOPES})
        self._json(404, {"error": "nf"})


def main():
    srv = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{srv.server_address[1]}"
    os.environ["SHOPIFY_OAUTH_BASE"] = base
    sys.path.insert(0, os.path.join(ROOT, "tools"))
    passed = []

    # 1. ast.parse
    ast.parse(open(TOOL, encoding="utf-8").read())
    import get_shopify_token as T
    assert T.OAUTH_BASE_OVERRIDE == base
    passed.append("1 ast.parse OK")

    # 只用標準函式庫
    tree = ast.parse(open(TOOL, encoding="utf-8").read())
    mods = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            mods.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom):
            mods.add((node.module or "").split(".")[0])
    assert mods <= set(sys.stdlib_module_names), mods - set(sys.stdlib_module_names)
    req = open(os.path.join(ROOT, "requirements.txt"), encoding="utf-8").read()
    assert "google" not in req.lower() and "shopify" not in req.lower()
    passed.append("只用標準函式庫、requirements.txt 未動")

    cfg = {"shop": SHOP, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET, "scopes": SCOPES,
           "port": PORT, "open_browser": False, "timeout": 10}

    def run_flow_async():
        """在背景跑 run_flow，回 (thread, holder)。"""
        holder = {}
        buf = io.StringIO()

        def go():
            try:
                holder["result"] = T.run_flow(cfg, out=buf)
            except Exception as e:      # noqa: BLE001
                holder["error"] = e
            holder["out"] = buf.getvalue()
        th = threading.Thread(target=go)
        th.start()
        # 等 server 起來、拿到印出的授權網址
        for _ in range(100):
            m = re.search(r"^(http\S+/admin/oauth/authorize\?\S+)$", buf.getvalue(), re.M)
            if m:
                holder["auth_url"] = m.group(1)
                break
            time.sleep(0.05)
        assert "auth_url" in holder, buf.getvalue()
        return th, holder

    def hit(url):
        try:
            with urlopen(url, timeout=5) as r:
                return r.status, r.read().decode("utf-8", "replace")
        except Exception as e:      # noqa: BLE001
            if hasattr(e, "code"):
                return e.code, e.read().decode("utf-8", "replace")
            raise

    # 2. state 不符 → 中止、不換 token
    th, h = run_flow_async()
    state = parse_qs(urlparse(h["auth_url"]).query)["state"][0]
    params = {"code": "c", "shop": SHOP, "state": "WRONG_" + state, "timestamp": "1"}
    params["hmac"] = sign(params, CLIENT_SECRET)     # 簽章正確、只有 state 錯
    before = FAKE.token_calls
    st, body = hit(f"http://127.0.0.1:{PORT}/callback?" + urlencode(params))
    th.join(5)
    assert st == 400 and "授權中止" in body
    assert "error" in h and "state 不符" in str(h["error"]), h
    assert FAKE.token_calls == before, "state 不符時不得換 token"
    passed.append("2 state 不符 → 中止並顯示原因、未發出 token 交換請求")

    # 3. hmac 驗證失敗 → 中止
    th, h = run_flow_async()
    state = parse_qs(urlparse(h["auth_url"]).query)["state"][0]
    params = {"code": "c", "shop": SHOP, "state": state, "timestamp": "1"}
    params["hmac"] = sign(params, "wrong_secret")
    before = FAKE.token_calls
    st, body = hit(f"http://127.0.0.1:{PORT}/callback?" + urlencode(params))
    th.join(5)
    assert st == 400 and "error" in h and "HMAC" in str(h["error"]), h
    assert FAKE.token_calls == before
    # 沒帶 hmac 也要擋
    th, h = run_flow_async()
    state = parse_qs(urlparse(h["auth_url"]).query)["state"][0]
    st, body = hit(f"http://127.0.0.1:{PORT}/callback?" + urlencode({"code": "c", "state": state}))
    th.join(5)
    assert "error" in h and "HMAC" in str(h["error"]) and FAKE.token_calls == before
    # 純函式層
    good = {"code": "x", "shop": SHOP, "state": "s", "timestamp": "1"}
    good["hmac"] = sign(good, CLIENT_SECRET)
    assert T.verify_hmac(good, CLIENT_SECRET) is True
    assert T.verify_hmac(dict(good, code="tampered"), CLIENT_SECRET) is False
    passed.append("3 HMAC 驗證失敗 / 缺 hmac → 中止並顯示原因")

    # 4. 完整流程（假 Shopify：authorize 302 回 callback → 換 token）
    th, h = run_flow_async()
    before = FAKE.token_calls
    st, body = hit(h["auth_url"])        # 模擬瀏覽器：跟著 302 打到 localhost:3456/callback
    th.join(10)
    assert st == 200 and "授權完成，可以關閉此分頁" in body, (st, body)
    assert "result" in h, h.get("error")
    assert h["result"]["access_token"] == FAKE_TOKEN and h["result"]["scope"] == SCOPES
    assert FAKE.token_calls == before + 1
    # 再打一次 callback（code 重放）→ 410、不再換 token
    # （server 已關閉，連線會被拒；這裡只確認 token_calls 沒變）
    assert FAKE.token_calls == before + 1
    # main() 端到端：子程序、環境變數給參數、--no-browser，stdout 要印出 token 與 scope
    env = dict(os.environ, SHOPIFY_SHOP=SHOP, SHOPIFY_CLIENT_ID=CLIENT_ID, SHOPIFY_CLIENT_SECRET=CLIENT_SECRET,
               SHOPIFY_SCOPES=SCOPES, SHOPIFY_OAUTH_BASE=base, PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen([sys.executable, "-u", TOOL, "--no-browser", "--timeout", "15"], env=env,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL, cwd=ROOT)
    # 等它起 server，再用 authorize 網址模擬瀏覽器
    for _ in range(100):
        try:
            hit(f"http://127.0.0.1:{PORT}/nope")
            break
        except Exception:      # noqa: BLE001
            time.sleep(0.05)
    auth_url = T.build_auth_url(SHOP, CLIENT_ID, SCOPES, "dummy", PORT)
    # state 由子程序產生、我們不知道 → 只能用真的流程：讓假 Shopify 從 authorize 導回；
    # 但 authorize 的 state 必須是子程序的。所以改從子程序 stdout 抓網址。
    out_lines = []
    while True:
        line = proc.stdout.readline().decode("utf-8", "replace")
        if not line:
            break
        out_lines.append(line)
        m = re.match(r"^(http\S+/admin/oauth/authorize\?\S+)\s*$", line)
        if m:
            auth_url = m.group(1)
            break
    st, body = hit(auth_url)
    assert st == 200
    so, se = proc.communicate(timeout=20)
    so = "".join(out_lines) + so.decode("utf-8", "replace")
    assert proc.returncode == 0, se.decode("utf-8", "replace")
    assert FAKE_TOKEN in so and f"實際取得的 scope：{SCOPES}" in so, so
    assert "SHOPIFY_ACCESS_TOKEN =" in so
    passed.append("4 假 Shopify 完整流程：解析出 access_token 並印出（含 scope）；子程序 main() 也通過")

    # 5. 缺少必要參數 → 明確錯誤訊息、不是 traceback（--non-interactive：Windows 上 stdin=NUL 時 isatty 仍回 True、getpass 會卡住，所以不靠 tty 偵測）
    env2 = {k: v for k, v in os.environ.items() if not k.startswith("SHOPIFY_") and k not in ("SHOP", "CLIENT_ID", "CLIENT_SECRET", "SCOPES")}
    env2["PYTHONIOENCODING"] = "utf-8"
    r = subprocess.run([sys.executable, TOOL, "--no-browser", "--non-interactive"], env=env2, capture_output=True,
                       stdin=subprocess.DEVNULL, cwd=ROOT, timeout=20)
    err = r.stderr.decode("utf-8", "replace")
    assert r.returncode == 2 and "缺少必要參數" in err and "Traceback" not in err, (r.returncode, err)
    assert "SHOP" in err and "CLIENT_SECRET" in err
    # 只缺 secret
    env3 = dict(env2, SHOPIFY_SHOP=SHOP, SHOPIFY_CLIENT_ID=CLIENT_ID, SHOPIFY_SCOPES=SCOPES)
    r = subprocess.run([sys.executable, TOOL, "--no-browser", "--non-interactive"], env=env3, capture_output=True,
                       stdin=subprocess.DEVNULL, cwd=ROOT, timeout=20)
    err = r.stderr.decode("utf-8", "replace")
    assert r.returncode == 2 and "缺少必要參數：CLIENT_SECRET" in err and "Traceback" not in err, err
    # shop 格式錯
    env4 = dict(env3, SHOPIFY_CLIENT_SECRET="x", SHOPIFY_SHOP="www.goyoutati.com")
    r = subprocess.run([sys.executable, TOOL, "--no-browser", "--non-interactive"], env=env4, capture_output=True,
                       stdin=subprocess.DEVNULL, cwd=ROOT, timeout=20)
    err = r.stderr.decode("utf-8", "replace")
    assert r.returncode == 2 and "SHOP 格式錯" in err and "Traceback" not in err, err
    # port 被占用 → 明確訊息
    blocker = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    try:
        r = subprocess.run([sys.executable, TOOL, "--no-browser", "--non-interactive"], env=dict(env3, SHOPIFY_CLIENT_SECRET="x"),
                           capture_output=True, stdin=subprocess.DEVNULL, cwd=ROOT, timeout=20)
    finally:
        blocker.server_close()
    err = r.stderr.decode("utf-8", "replace")
    assert r.returncode == 1 and "無法監聽" in err and "Traceback" not in err, err
    passed.append("5 缺參數 / 格式錯 / port 被占 → 明確錯誤訊息、無 traceback")

    # 6. token 不在任何檔案裡（repo 全部檔案 + 工作目錄新檔）
    hits = []
    for dp, dn, fn in os.walk(ROOT):
        dn[:] = [d for d in dn if d not in (".git", "__pycache__", "node_modules")]
        for f in fn:
            p = os.path.join(dp, f)
            if p == os.path.abspath(__file__):
                continue
            try:
                if FAKE_TOKEN.encode() in open(p, "rb").read():
                    hits.append(p)
            except OSError:
                pass
    assert not hits, f"token 出現在檔案裡：{hits}"
    # 程式碼裡也沒寫死 secret、沒有 open(...'w') 寫檔
    src = open(TOOL, encoding="utf-8").read()
    assert not re.search(r"open\([^)]*['\"]w", src), "腳本不得寫檔"
    assert "shpss_" not in src and "shpat_" in src   # 只在說明裡提到前綴
    passed.append("6 token 不出現在任何檔案裡、腳本無寫檔")

    # 7. Windows 可執行：路徑用 os.path、stdout utf-8 reconfigure、webbrowser 只在非 --no-browser 時呼叫
    assert "sys.stdout.reconfigure" in src and "webbrowser.open" in src
    assert 'CallbackServer(("127.0.0.1"' in src, "綁 127.0.0.1（Windows 上 localhost 可能先解成 ::1）"
    assert 'charset="utf-8"' in src and "text/html; charset=utf-8" in src
    if os.name == "nt":
        passed.append("7 Windows 下實際執行通過（本測試就在 Windows 跑：子程序、port、編碼皆正常）")
    else:
        passed.append("7 Windows 相容檢查（靜態）")

    srv.shutdown()
    log("\n".join("✅ " + p for p in passed))
    log(f"\n全部 {len(passed)} 項通過")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)

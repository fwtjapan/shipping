# -*- coding: utf-8 -*-
"""
SQLite → Google Drive 自動備份 驗收測試（任務 AA）。
用假的 Drive API server 攔截所有請求，不碰真的 Google。

執行：python test_backup.py
（合成資料，不含真實客戶資料。）
"""
import ast
import gzip
import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def log(msg):
    sys.stderr.write(msg + "\n")
    sys.stderr.flush()


# ═══════════════════════════════════════════════════════════════
# 假的 Google Drive / OAuth server
# ═══════════════════════════════════════════════════════════════
class FakeDrive:
    def __init__(self):
        self.files = {}          # id → {name, mimeType, parents, size, trashed, data}
        self.seq = 0
        self.fail_mode = None    # None | token401 | upload403 | upload_timeout
        self.upload_delay = 0.0
        self.uploads = 0
        self.token_calls = 0
        self.lock = threading.Lock()

    def new_id(self, prefix="f"):
        with self.lock:
            self.seq += 1
            return f"{prefix}{self.seq:04d}"

    def children(self, parent):
        return [dict(id=i, **{k: v for k, v in f.items() if k != "data"})
                for i, f in self.files.items() if parent in f["parents"] and not f["trashed"]]

    def folders_named(self, name):
        return [dict(id=i, name=f["name"]) for i, f in self.files.items()
                if f["name"] == name and f["mimeType"].endswith("folder") and not f["trashed"]]


FAKE = FakeDrive()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _auth_ok(self):
        return self.headers.get("Authorization") == "Bearer fake-access-token"

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(n) if n else b""
        if u.path == "/token":
            FAKE.token_calls += 1
            q = parse_qs(body.decode())
            if FAKE.fail_mode == "token401" or q.get("grant_type") != ["refresh_token"] \
                    or q.get("refresh_token") != ["fake-refresh"]:
                return self._json(401, {"error": "invalid_grant", "error_description": "Bad refresh token"})
            return self._json(200, {"access_token": "fake-access-token", "expires_in": 3600, "token_type": "Bearer"})
        if not self._auth_ok():
            return self._json(401, {"error": {"code": 401, "message": "Invalid Credentials"}})
        if u.path == "/drive/v3/files":
            meta = json.loads(body.decode())
            fid = FAKE.new_id("d")
            FAKE.files[fid] = {"name": meta["name"], "mimeType": meta.get("mimeType", ""),
                               "parents": meta.get("parents") or [], "size": "0", "trashed": False, "data": b""}
            return self._json(200, {"id": fid})
        if u.path == "/upload/drive/v3/files":
            if FAKE.fail_mode == "upload403":
                return self._json(403, {"error": {"code": 403, "message": "storageQuotaExceeded"}})
            if FAKE.fail_mode == "upload_timeout":
                time.sleep(4)
                try:
                    return self._json(500, {"error": "too late"})
                except OSError:      # client 早就逾時斷線了
                    return
            if FAKE.upload_delay:
                time.sleep(FAKE.upload_delay)
            ctype = self.headers.get("Content-Type", "")
            m = re.search(r'boundary=([^;]+)', ctype)
            assert "multipart/related" in ctype and m, ctype
            boundary = m.group(1).encode()
            parts = body.split(b"--" + boundary)
            parts = [p for p in parts if p.strip() not in (b"", b"--")]
            assert len(parts) == 2, len(parts)
            meta = json.loads(parts[0].split(b"\r\n\r\n", 1)[1].rstrip(b"\r\n").decode())
            data = parts[1].split(b"\r\n\r\n", 1)[1]
            if data.endswith(b"\r\n"):
                data = data[:-2]
            fid = FAKE.new_id("u")
            FAKE.files[fid] = {"name": meta["name"], "mimeType": "application/gzip",
                               "parents": meta.get("parents") or [], "size": str(len(data)),
                               "trashed": False, "data": data}
            FAKE.uploads += 1
            return self._json(200, {"id": fid, "name": meta["name"], "size": str(len(data))})
        return self._json(404, {"error": "no route"})

    def do_GET(self):
        u = urlparse(self.path)
        if not self._auth_ok():
            return self._json(401, {"error": {"code": 401, "message": "Invalid Credentials"}})
        q = parse_qs(u.query)
        if u.path == "/drive/v3/files":
            qq = (q.get("q") or [""])[0]
            m = re.search(r"'([^']+)' in parents", qq)
            if m:
                files = FAKE.children(m.group(1))
                return self._json(200, {"files": [{"id": f["id"], "name": f["name"], "size": f["size"],
                                                   "createdTime": ""} for f in files]})
            m = re.search(r"name = '([^']+)'", qq)
            if m:
                return self._json(200, {"files": FAKE.folders_named(m.group(1))})
            return self._json(400, {"error": "bad q"})
        m = re.match(r"^/drive/v3/files/([^/]+)$", u.path)
        if m:
            f = FAKE.files.get(m.group(1))
            if not f:
                return self._json(404, {"error": {"code": 404, "message": "File not found"}})
            return self._json(200, {"id": m.group(1), "name": f["name"], "mimeType": f["mimeType"],
                                    "trashed": f["trashed"], "size": f["size"]})
        return self._json(404, {"error": "no route"})

    def do_DELETE(self):
        if not self._auth_ok():
            return self._json(401, {"error": "auth"})
        m = re.match(r"^/drive/v3/files/([^/]+)$", urlparse(self.path).path)
        if m and m.group(1) in FAKE.files:
            del FAKE.files[m.group(1)]
            self.send_response(204)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        return self._json(404, {"error": "nf"})


def start_fake_server():
    srv = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, f"http://127.0.0.1:{srv.server_address[1]}"


# ═══════════════════════════════════════════════════════════════
# 子程序：三個環境變數皆空 → 程式正常啟動、功能正常、警告只印一次
# ═══════════════════════════════════════════════════════════════
def child_disabled():
    os.environ["GDRIVE_CLIENT_ID"] = ""
    os.environ["GDRIVE_CLIENT_SECRET"] = ""
    os.environ["GDRIVE_REFRESH_TOKEN"] = ""
    import app as A
    c = A.app.test_client()
    r = c.get("/api/config")
    assert r.status_code == 200, r.status_code
    with c.session_transaction() as s:
        s["user_type"] = "admin"; s["role"] = "super"; s["user_id"] = 1; s["agent_id"] = 0; s["username"] = "Admin"
    st = c.get("/api/admin/maintenance/backup_status").get_json()
    assert st["success"] and st["enabled"] is False and st["stale"] is True, st
    r = c.post("/api/admin/maintenance/backup_now", json={})
    assert r.status_code == 400 and r.get_json().get("disabled"), (r.status_code, r.get_json())
    r = c.get("/admin")
    assert r.status_code == 200 and b"backupCard" in r.data
    rep = A.run_backup_job("child")
    assert rep.get("disabled") and not rep["success"]
    assert not any(t.name == "GDriveBackup" for t in threading.enumerate())
    print("CHILD_OK", flush=True)


# ═══════════════════════════════════════════════════════════════
# 主測試
# ═══════════════════════════════════════════════════════════════
def main():
    tmp = tempfile.mkdtemp(prefix="hsbk_")
    db_path = os.path.join(tmp, "packages.db")
    local_dir = os.path.join(tmp, "backups")
    srv, base = start_fake_server()

    env = {
        "DB_PATH": db_path, "SHOPIFY_STORE": "", "MEMBER_SYNC_AUTO": "0", "BACKUP_AUTO": "0",
        "BACKUP_LOCAL_DIR": local_dir,
        "GDRIVE_CLIENT_ID": "fake-client", "GDRIVE_CLIENT_SECRET": "fake-secret", "GDRIVE_REFRESH_TOKEN": "fake-refresh",
        "GDRIVE_TOKEN_URL": base + "/token", "GDRIVE_API_BASE": base,
    }
    os.environ.update(env)
    passed = []

    # ── 1. 語法 / node --check / Jinja render ──
    ast.parse(open(os.path.join(ROOT, "app.py"), encoding="utf-8").read())
    ast.parse(open(os.path.join(ROOT, "gdrive_backup.py"), encoding="utf-8").read())
    ast.parse(open(os.path.join(ROOT, "tools", "get_gdrive_token.py"), encoding="utf-8").read())
    html = open(os.path.join(ROOT, "templates", "admin.html"), encoding="utf-8").read()
    scripts = re.findall(r"<script>(.*?)</script>", html, re.S)
    assert len(scripts) >= 2, len(scripts)
    jsf = os.path.join(tmp, "admin_all.js")
    open(jsf, "w", encoding="utf-8").write("\n;\n".join(scripts))
    r = subprocess.run(["node", "--check", jsf], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    import app as A
    import gdrive_backup as G
    c = A.app.test_client()
    r = c.get("/admin")
    assert r.status_code == 200 and "backupCard" in r.get_data(as_text=True)
    passed.append("1 語法/JS/Jinja OK")

    # ── 2. 三個環境變數皆空（子程序）──
    cenv = dict(os.environ)
    for k in ("GDRIVE_CLIENT_ID", "GDRIVE_CLIENT_SECRET", "GDRIVE_REFRESH_TOKEN"):
        cenv[k] = ""
    cenv["DB_PATH"] = os.path.join(tmp, "child.db")
    cenv["PYTHONIOENCODING"] = "utf-8"
    r = subprocess.run([sys.executable, os.path.abspath(__file__), "--child-disabled"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", env=cenv, cwd=ROOT)
    assert r.returncode == 0, r.stderr[-3000:]
    assert "CHILD_OK" in r.stdout, r.stdout[-2000:]
    n_warn = r.stdout.count("[backup] ⚠️ 自動備份停用")
    assert n_warn == 1, f"警告應只印一次，實際 {n_warn}\n{r.stdout[-2000:]}"
    passed.append("2 憑證為空：正常啟動、功能正常、警告只印一次")

    # ── 準備資料 ──
    def boss():
        with c.session_transaction() as s:
            s["user_type"] = "admin"; s["role"] = "super"; s["user_id"] = 1; s["agent_id"] = 0; s["username"] = "Admin"

    def staff():
        with c.session_transaction() as s:
            s["user_type"] = "admin"; s["role"] = "staff"; s["user_id"] = 2; s["agent_id"] = 0; s["username"] = "琤茵"

    def anon():
        with c.session_transaction() as s:
            s.clear()

    conn = A.get_db()
    for i in range(20):
        conn.execute("INSERT INTO packages (g_code, logis_num, product_name, weight, status, in_date, created_at) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?)",
                     (f"G{i:04d}", f"LN{i}", "測試品", "1.0", "在倉", "2026-09-20", "2026-09-20 00:00:00"))
    conn.commit()
    conn.close()
    src_rows = G.table_row_count(db_path)
    assert src_rows > 20

    def unz_rows(data):
        p = os.path.join(tmp, f"chk_{time.time_ns()}.db")
        open(p, "wb").write(gzip.decompress(data))
        conn = sqlite3.connect(p)
        try:
            assert conn.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
            n = sum(conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0] for (t,) in
                    conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"))
            npk = conn.execute("SELECT COUNT(*) FROM packages").fetchone()[0]
        finally:
            conn.close()
        return n, npk

    # ── 3. 端到端 ──
    rep = A.run_backup_job("test-e2e")
    assert rep["success"], rep
    assert rep["filename"].startswith("packages-") and rep["filename"].endswith(".db.gz")
    assert re.match(r"^packages-\d{8}_\d{6}\.db\.gz$", rep["filename"]), rep["filename"]
    # 備份流程自己會先寫 admin_settings（租約、attempt 時間戳），快照列數 ≥ 事前算的
    assert src_rows <= rep["rows"] <= src_rows + 3, (src_rows, rep["rows"])
    folder = [f for f in FAKE.files.values() if f["name"] == "helpshipping-backups"]
    assert len(folder) == 1, "應建立 helpshipping-backups 資料夾（且只建一個）"
    assert A._get_setting("gdrive_backup_folder_id", "") == rep["folder_id"] != ""
    up = [FAKE.files[rep["drive_file_id"]]]
    assert up[0]["name"] == rep["filename"] and up[0]["parents"] == [rep["folder_id"]]
    n, npk = unz_rows(up[0]["data"])
    # 備份過程本身會寫 admin_settings（attempt 時間戳），快照時點是 attempt 之後、report 之前 → packages 列數必須一樣
    assert npk == 20 and n == rep["rows"], (n, npk, rep["rows"])
    assert os.path.exists(os.path.join(local_dir, rep["filename"]))
    assert open(os.path.join(local_dir, rep["filename"]), "rb").read() == up[0]["data"]
    assert not [x for x in os.listdir(local_dir) if x.startswith(".")], ("暫存檔應清乾淨", os.listdir(local_dir))
    assert rep["size"] == int(up[0]["size"]) and rep["size"] < os.path.getsize(db_path), "gzip 應比原檔小"
    boss()
    st = c.get("/api/admin/maintenance/backup_status").get_json()
    assert st["enabled"] and st["last_ok_at"] == rep["at"] and st["stale"] is False
    assert st["drive_count"] == 1 and st["latest_size"] == rep["size"] and st["local_count"] == 1
    # 第二次：重用 folder id，不再建新資料夾
    rep2 = A.run_backup_job("test-e2e-2")
    assert rep2["success"] and rep2["folder_id"] == rep["folder_id"] and not rep2["folder_created"]
    assert len([f for f in FAKE.files.values() if f["name"] == "helpshipping-backups"]) == 1
    passed.append("3 端到端：快照→gzip→驗證→上傳→本機留一份，解壓列數一致、資料夾重用")

    # ── 4. WAL：另一條連線寫入不 checkpoint → copy2 缺資料、backup API 完整 ──
    wal_db = os.path.join(tmp, "wal.db")
    w = sqlite3.connect(wal_db)
    w.execute("PRAGMA journal_mode=WAL")
    w.execute("CREATE TABLE t (id INTEGER PRIMARY KEY, v TEXT)")
    w.commit()
    w.execute("PRAGMA wal_checkpoint(TRUNCATE)")   # 先把建表清到主檔，之後的寫入只留在 WAL
    w.execute("PRAGMA wal_autocheckpoint=0")
    for i in range(500):
        w.execute("INSERT INTO t (v) VALUES (?)", ("x" * 100,))
    w.commit()                                     # 連線保持開著、不 checkpoint
    assert os.path.getsize(wal_db + "-wal") > 0
    # 反例：shutil.copy2 只複製主檔 → 打得開、integrity ok、但 500 列不見（正是這種備份最危險）
    bad = os.path.join(tmp, "copy2.db")
    shutil.copy2(wal_db, bad)
    bc = sqlite3.connect(bad)
    bad_rows = bc.execute("SELECT COUNT(*) FROM t").fetchone()[0]
    assert bc.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    bc.close()
    assert bad_rows == 0, f"copy2 的反例應該缺資料（測試才有鑑別力），實際 {bad_rows}"
    # 正解：backup API
    good = os.path.join(tmp, "backupapi.db")
    G.snapshot_db(wal_db, good)
    gc = sqlite3.connect(good)
    good_rows = gc.execute("SELECT COUNT(*) FROM t").fetchone()[0]
    gc.close()
    assert good_rows == 500, good_rows
    # 同樣情境走整條備份流程（app 的 DB 是 WAL）：保持寫入連線、不 checkpoint、再備份
    holder = sqlite3.connect(db_path)
    holder.execute("PRAGMA wal_autocheckpoint=0")
    for i in range(30):
        holder.execute("INSERT INTO packages (g_code, logis_num, product_name, weight, status, in_date, created_at) "
                       "VALUES (?, ?, ?, ?, ?, ?, ?)",
                       (f"W{i:04d}", f"WAL{i}", "WAL品", "1.0", "在倉", "2026-09-20", "2026-09-20 00:00:00"))
    holder.commit()
    rep = A.run_backup_job("test-wal")
    assert rep["success"], rep
    up = FAKE.files[rep["drive_file_id"]]      # 同一秒內連跑兩次檔名會相同，用 Drive 回的 id 抓
    n, npk = unz_rows(up["data"])
    assert npk == 50, f"備份必須含 WAL 裡最新交易（期望 50 列 packages，實際 {npk}）"
    holder.close()
    w.close()
    passed.append("4 WAL：copy2 反例掉 500 列（鑑別力 OK）、backup API 完整；整條流程含最新交易")

    # ── 5. integrity_check 失敗 → 不上傳、記為失敗、狀態端點看得到原因 ──
    corrupt = os.path.join(tmp, "corrupt.db")
    cc = sqlite3.connect(corrupt)
    cc.execute("CREATE TABLE t (id INTEGER PRIMARY KEY, v TEXT)")
    cc.executemany("INSERT INTO t (v) VALUES (?)", [("y" * 200,)] * 300)
    cc.commit(); cc.close()
    raw = bytearray(open(corrupt, "rb").read())
    page = 4096 if len(raw) >= 4096 * 3 else 1024
    raw[page * 2 + 8: page * 2 + 8 + 200] = b"\xff" * 200   # 砸爛第 3 頁
    open(corrupt, "wb").write(bytes(raw))
    ok, reason, _ = G.verify_snapshot(corrupt)
    assert not ok and "integrity_check" in reason, reason
    orig_snapshot = G.snapshot_db
    G.snapshot_db = lambda src, dst: shutil.copyfile(corrupt, dst)
    try:
        before_uploads = FAKE.uploads
        before_local = sorted(os.listdir(local_dir))
        rep = A.run_backup_job("test-corrupt")
    finally:
        G.snapshot_db = orig_snapshot
    assert not rep["success"] and rep["stage"] == "verify" and "integrity_check" in rep["error"], rep
    assert FAKE.uploads == before_uploads, "驗證失敗不得上傳"
    assert sorted(os.listdir(local_dir)) == before_local, "驗證失敗不得留本機檔、暫存檔要清掉"
    st = c.get("/api/admin/maintenance/backup_status").get_json()
    assert st["last_fail_at"] == rep["at"] and "integrity_check" in st["last_fail_reason"], st
    assert st["stale"] is False, "上一次成功還在 48 小時內，不算逾時"
    passed.append("5 integrity_check 失敗：不上傳、記失敗、狀態端點有原因")

    # ── 6. Drive API 錯誤（401 / 403 / 逾時）→ 不拋例外、不影響其他功能、記錄原因 ──
    G_UP_TO = G.UPLOAD_TIMEOUT
    G.UPLOAD_TIMEOUT = 1.5
    try:
        for mode, needle in (("token401", "401"), ("upload403", "403"), ("upload_timeout", "Timeout")):
            FAKE.fail_mode = mode
            before_uploads = FAKE.uploads
            before_local = sorted(os.listdir(local_dir))
            rep = A.run_backup_job(f"test-{mode}")
            assert not rep["success"] and needle in rep["error"], (mode, rep)
            assert FAKE.uploads == before_uploads
            assert sorted(os.listdir(local_dir)) == before_local, "失敗不得留本機檔"
            assert c.get("/api/config").status_code == 200, "其他功能必須正常"
            st = c.get("/api/admin/maintenance/backup_status").get_json()
            assert needle in st["last_fail_reason"], st["last_fail_reason"]
            assert st["running"] is False, "租約必須釋放"
    finally:
        FAKE.fail_mode = None
        G.UPLOAD_TIMEOUT = G_UP_TO
    passed.append("6 Drive 401/403/逾時：不拋例外、不影響其他功能、原因可見、租約釋放")

    # ── 7. 清理：35 份假備份 → Drive 剩 30、本機剩 7，刪的是最舊的 ──
    folder_id = A._get_setting("gdrive_backup_folder_id", "")
    for f in [k for k, v in FAKE.files.items() if v["name"].startswith("packages-")]:
        del FAKE.files[f]
    for x in os.listdir(local_dir):
        os.remove(os.path.join(local_dir, x))
    base_t = datetime(2026, 1, 1, 3, 0, 0)
    fake_names = [f"packages-{(base_t + timedelta(days=i)).strftime('%Y%m%d_%H%M%S')}.db.gz" for i in range(35)]
    for nme in fake_names:
        FAKE.files[FAKE.new_id("o")] = {"name": nme, "mimeType": "application/gzip", "parents": [folder_id],
                                        "size": "10", "trashed": False, "data": b"0123456789"}
        open(os.path.join(local_dir, nme), "wb").write(b"0123456789")
    rep = A.run_backup_job("test-prune")
    assert rep["success"], rep
    drive_names = sorted(f["name"] for f in FAKE.files.values() if f["name"].startswith("packages-"))
    local_names = sorted(os.listdir(local_dir))
    assert len(drive_names) == 30, len(drive_names)
    assert len(local_names) == 7, len(local_names)
    expected_drive = sorted(fake_names + [rep["filename"]])[-30:]
    expected_local = sorted(fake_names + [rep["filename"]])[-7:]
    assert drive_names == expected_drive, "Drive 應刪最舊的"
    assert local_names == expected_local, "本機應刪最舊的"
    assert rep["filename"] in drive_names and rep["filename"] in local_names
    assert rep["drive_count"] == 30 and rep["local_count"] == 7
    assert len(rep["pruned_drive"]) == 6 and len(rep["pruned_local"]) == 29
    passed.append("7 清理：Drive 30 / 本機 7，刪最舊")

    # ── 8. 租約：兩個 worker 同時觸發 → 只有一個實際執行 ──
    FAKE.upload_delay = 1.5
    before_uploads = FAKE.uploads
    results = []

    def worker(tag):
        results.append((tag, A.run_backup_job(f"worker-{tag}")))

    ths = [threading.Thread(target=worker, args=(i,)) for i in range(2)]
    for t in ths:
        t.start()
    for t in ths:
        t.join()
    FAKE.upload_delay = 0
    okc = [r for _, r in results if r.get("success")]
    skc = [r for _, r in results if r.get("skipped")]
    assert len(okc) == 1 and len(skc) == 1, results
    assert FAKE.uploads == before_uploads + 1
    assert A._backup_is_running() is False
    # 租約值存在 admin_settings（不是記憶體變數）
    assert A._get_setting("backup_lease", None) == "0"
    passed.append("8 租約：兩個 worker 同時觸發只跑一個，租約在 admin_settings")

    # ── 9. 權限 ──
    staff()
    assert c.get("/api/admin/maintenance/backup_status").status_code == 200
    assert c.post("/api/admin/maintenance/backup_now", json={}).status_code == 403
    assert c.put("/api/admin/settings/backup_hour", json={"backup_hour": 5}).status_code == 403
    anon()
    assert c.get("/api/admin/maintenance/backup_status").status_code == 403
    assert c.post("/api/admin/maintenance/backup_now", json={}).status_code == 403
    assert c.put("/api/admin/settings/backup_hour", json={"backup_hour": 5}).status_code == 403
    boss()
    r = c.post("/api/admin/maintenance/backup_now", json={})
    assert r.status_code == 200 and r.get_json()["success"], r.get_json()
    r = c.put("/api/admin/settings/backup_hour", json={"backup_hour": 5})
    assert r.status_code == 200 and r.get_json()["backup_hour"] == 5 and r.get_json()["old"] == 3
    assert c.put("/api/admin/settings/backup_hour", json={"backup_hour": 24}).status_code == 400
    assert c.get("/api/admin/maintenance/backup_status").get_json()["backup_hour"] == 5
    conn = A.get_db()
    ops = [dict(r) for r in conn.execute("SELECT * FROM operation_logs WHERE action IN ('手動備份','修改備份時間')")]
    conn.close()
    assert len(ops) >= 2
    passed.append("9 權限：員工可看不可觸發、未登入 403、老闆可觸發/改時間、有操作紀錄")

    # ── 排程到期判斷 ──
    A._set_setting("backup_hour", "3")
    A._set_setting("backup_last_attempt_at", "2026-01-01 00:00:00")
    A._set_setting("backup_last_ok_at", "2026-09-19 03:00:12")
    assert A._backup_due(datetime(2026, 9, 20, 2, 59)) is False, "還沒到 3 點"
    assert A._backup_due(datetime(2026, 9, 20, 3, 0)) is True, "3 點到了、今天還沒成功"
    A._set_setting("backup_last_ok_at", "2026-09-20 03:00:12")
    assert A._backup_due(datetime(2026, 9, 20, 23, 0)) is False, "今天成功過就不再跑"
    A._set_setting("backup_last_ok_at", "2026-09-19 03:00:12")
    A._set_setting("backup_last_attempt_at", (datetime.now() - timedelta(minutes=5)).strftime("%Y-%m-%d %H:%M:%S"))
    assert A._backup_due(datetime(2026, 9, 20, 10, 0)) is False, "5 分鐘前才試過（失敗退避 30 分）"
    A._set_setting("backup_last_attempt_at", (datetime.now() - timedelta(minutes=31)).strftime("%Y-%m-%d %H:%M:%S"))
    assert A._backup_due(datetime(2026, 9, 20, 10, 0)) is True
    passed.append("排程：到期判斷（時間、當日一次、失敗退避）")

    # ── 10. 後台顯示：最後備份時間正確；50 小時前 → stale + 標紅警告 ──
    m = re.search(r"function renderBackupStatus\(d\) \{.*?\n\}\n", html, re.S)
    assert m, "admin.html 缺 renderBackupStatus"
    fn_src = m.group(0)
    now = datetime.now().replace(microsecond=0)
    fresh = now - timedelta(hours=1)
    old = now - timedelta(hours=50)
    A._set_setting("backup_last_ok_at", fresh.strftime("%Y-%m-%d %H:%M:%S"))
    st_fresh = c.get("/api/admin/maintenance/backup_status").get_json()
    assert st_fresh["stale"] is False and st_fresh["last_ok_at"] == fresh.strftime("%Y-%m-%d %H:%M:%S")
    A._set_setting("backup_last_ok_at", old.strftime("%Y-%m-%d %H:%M:%S"))
    st_old = c.get("/api/admin/maintenance/backup_status").get_json()
    assert st_old["stale"] is True and 49.9 <= st_old["hours_since_ok"] <= 50.1, st_old["hours_since_ok"]
    A._set_setting("backup_last_ok_at", "")
    st_never = c.get("/api/admin/maintenance/backup_status").get_json()
    assert st_never["stale"] is True and st_never["last_ok_at"] == ""
    node_js = fn_src + "\nconst cases = " + json.dumps({"fresh": st_fresh, "old": st_old, "never": st_never,
                                                        "disabled": dict(st_fresh, enabled=False)}) + ";\n" \
        + "const out = {}; for (const k in cases) out[k] = renderBackupStatus(cases[k]);\n" \
        + "process.stdout.write(JSON.stringify(out));\n"
    jsp = os.path.join(tmp, "render.js")
    open(jsp, "w", encoding="utf-8").write(node_js)
    r = subprocess.run(["node", jsp], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    out = json.loads(r.stdout)
    exp_fresh = fresh.strftime("%Y-%m-%d %H:%M")
    assert out["fresh"]["stale"] is False and "backup-warn" not in out["fresh"]["html"]
    assert f'color:#27ae60;">{exp_fresh}</strong>' in out["fresh"]["html"], out["fresh"]["html"]
    assert out["old"]["stale"] is True and "backup-warn" in out["old"]["html"]
    assert f'color:#ff7675;">{old.strftime("%Y-%m-%d %H:%M")}</strong>' in out["old"]["html"], out["old"]["html"]
    assert "已超過 50 小時沒有成功備份" in out["old"]["html"], out["old"]["html"]
    assert out["never"]["stale"] is True and "從未成功備份過" in out["never"]["html"] and "尚未備份" in out["never"]["html"]
    assert out["disabled"]["stale"] is True and "未啟用" in out["disabled"]["html"]
    passed.append("10 後台顯示：最後備份時間 YYYY-MM-DD HH:MM；50 小時前 → 標紅 + 警告；從未/未設定也警告")

    srv.shutdown()
    log("\n".join("✅ " + p for p in passed))
    log(f"\n全部 {len(passed)} 項通過")


if __name__ == "__main__":
    if "--child-disabled" in sys.argv:
        child_disabled()
    else:
        try:
            main()
        except Exception:
            traceback.print_exc()
            sys.exit(1)

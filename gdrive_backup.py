"""
SQLite 自動備份 → Google Drive（OAuth2 refresh token）

背景：2026-09-18 Zeabur volume 被誤刪，/data/packages.db 全部遺失，磁碟層面無法復原。
      前一晚的手動備份沒留住 → 問題在「備份靠人」。這支讓備份自動離開那台機器。

設計要點（不要改）：
  • 授權走 OAuth2 refresh token（一般 Gmail 帳號）。service account 沒有儲存配額，
    上傳必回 403 storageQuotaExceeded，即使分享資料夾給它也一樣——Google 平台限制。
  • scope 只用 drive.file（只能碰本程式自己建立的檔案），token 外洩也波及不到整個雲端硬碟。
    所以備份資料夾也由程式自己建（helpshipping-backups），id 由呼叫端存進 admin_settings 重用。
  • 不引入 google-api-python-client，用 requests 直接打兩支 REST endpoint。
  • 快照用 SQLite 官方 backup API（WAL 安全）。★不可用 shutil.copy2：DB 是 WAL 模式，
    copy2 只複製主檔，會得到缺最近交易、但照樣打得開的備份，很難察覺。
  • 驗證（integrity_check=ok 且列數>0）不過就不上傳、記為失敗。壞掉的備份比沒備份更危險。
  • 本模組不依賴 Flask / app.py；租約、排程、端點在 app.py。所有對外函式不拋例外，回 report dict。

環境變數：GDRIVE_CLIENT_ID / GDRIVE_CLIENT_SECRET / GDRIVE_REFRESH_TOKEN（任一為空 → 備份停用）。
測試用：GDRIVE_TOKEN_URL / GDRIVE_API_BASE 可指向假的 Drive server。
"""

import gzip
import hashlib
import json
import os
import sqlite3
import time
import uuid
from datetime import datetime

import requests

TOKEN_URL = os.environ.get("GDRIVE_TOKEN_URL", "https://oauth2.googleapis.com/token")
API_BASE = os.environ.get("GDRIVE_API_BASE", "https://www.googleapis.com").rstrip("/")
FOLDER_NAME = "helpshipping-backups"
FOLDER_MIME = "application/vnd.google-apps.folder"
FILE_PREFIX = "packages-"          # 檔名 packages-YYYYMMDD_HHMMSS.db.gz
FILE_SUFFIX = ".db.gz"
DRIVE_KEEP = 30                    # Drive 保留份數
LOCAL_KEEP = 7                     # 本機保留份數（同機不可靠，只當次要／恢復最快）
HTTP_TIMEOUT = 60                  # 秒；上傳另外給 UPLOAD_TIMEOUT
UPLOAD_TIMEOUT = 300


class GDriveError(Exception):
    """Drive / OAuth 呼叫失敗（HTTP 狀態碼、逾時、回應格式錯）。訊息會直接進 report.error。"""


def get_env_credentials():
    """回 (client_id, client_secret, refresh_token)；程式只讀不寫。"""
    return (
        (os.environ.get("GDRIVE_CLIENT_ID") or "").strip(),
        (os.environ.get("GDRIVE_CLIENT_SECRET") or "").strip(),
        (os.environ.get("GDRIVE_REFRESH_TOKEN") or "").strip(),
    )


def is_configured():
    return all(get_env_credentials())


def _short(resp_text, n=200):
    return (resp_text or "").strip().replace("\n", " ")[:n]


class GDriveClient:
    """最小 Drive v3 client：換 access token、找/建資料夾、multipart 上傳、列表、刪除。"""

    def __init__(self, client_id, client_secret, refresh_token, session=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self.http = session or requests.Session()
        self._access_token = None
        self._token_exp = 0

    # ── OAuth2 ──
    def access_token(self):
        if self._access_token and time.time() < self._token_exp - 60:
            return self._access_token
        try:
            r = self.http.post(TOKEN_URL, data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "refresh_token": self.refresh_token,
                "grant_type": "refresh_token",
            }, timeout=HTTP_TIMEOUT)
        except requests.RequestException as e:
            raise GDriveError(f"取 access token 失敗（連線）：{e.__class__.__name__}: {e}")
        if r.status_code != 200:
            raise GDriveError(f"取 access token 失敗 HTTP {r.status_code}：{_short(r.text)}")
        try:
            body = r.json()
            tok = body["access_token"]
        except (ValueError, KeyError, TypeError):
            raise GDriveError(f"取 access token 回應格式錯：{_short(r.text)}")
        self._access_token = tok
        self._token_exp = time.time() + int(body.get("expires_in") or 3600)
        return tok

    def _headers(self, extra=None):
        h = {"Authorization": f"Bearer {self.access_token()}"}
        if extra:
            h.update(extra)
        return h

    def _request(self, method, path, what, timeout=HTTP_TIMEOUT, **kw):
        url = path if path.startswith("http") else f"{API_BASE}{path}"
        try:
            r = self.http.request(method, url, headers=self._headers(kw.pop("headers", None)),
                                  timeout=timeout, **kw)
        except requests.RequestException as e:
            raise GDriveError(f"{what} 失敗（連線）：{e.__class__.__name__}: {e}")
        if r.status_code >= 300:
            raise GDriveError(f"{what} 失敗 HTTP {r.status_code}：{_short(r.text)}")
        return r

    # ── 資料夾 ──
    def get_file(self, file_id, fields="id,name,mimeType,trashed,size"):
        """存在且未進垃圾桶回 metadata，否則回 None（404 視為不存在，其他錯誤拋出）。"""
        url = f"{API_BASE}/drive/v3/files/{file_id}"
        try:
            r = self.http.get(url, headers=self._headers(), params={"fields": fields}, timeout=HTTP_TIMEOUT)
        except requests.RequestException as e:
            raise GDriveError(f"查檔案失敗（連線）：{e.__class__.__name__}: {e}")
        if r.status_code == 404:
            return None
        if r.status_code >= 300:
            raise GDriveError(f"查檔案失敗 HTTP {r.status_code}：{_short(r.text)}")
        meta = r.json()
        if meta.get("trashed"):
            return None
        return meta

    def find_folder(self, name=FOLDER_NAME):
        q = f"name = '{name}' and mimeType = '{FOLDER_MIME}' and trashed = false"
        r = self._request("GET", "/drive/v3/files", "找資料夾",
                          params={"q": q, "fields": "files(id,name)", "pageSize": 10, "spaces": "drive"})
        files = (r.json() or {}).get("files") or []
        return files[0]["id"] if files else None

    def create_folder(self, name=FOLDER_NAME):
        r = self._request("POST", "/drive/v3/files", "建資料夾",
                          params={"fields": "id"},
                          json={"name": name, "mimeType": FOLDER_MIME})
        return r.json()["id"]

    def ensure_folder(self, stored_id=None, name=FOLDER_NAME):
        """優先重用存起來的 id（確認還在且未被丟進垃圾桶），否則找同名（drive.file 只看得到自己建的），再不然就建。"""
        if stored_id:
            if self.get_file(stored_id):
                return stored_id
        fid = self.find_folder(name)
        if fid:
            return fid
        return self.create_folder(name)

    # ── 檔案 ──
    def upload(self, name, data, parent_id, mime="application/gzip"):
        """multipart/related 上傳（metadata JSON + 二進位）。回 {id,name,size}。"""
        boundary = "hs_backup_" + uuid.uuid4().hex
        meta = json.dumps({"name": name, "parents": [parent_id]}).encode("utf-8")
        body = b"".join([
            b"--", boundary.encode(), b"\r\n",
            b"Content-Type: application/json; charset=UTF-8\r\n\r\n", meta, b"\r\n",
            b"--", boundary.encode(), b"\r\n",
            b"Content-Type: ", mime.encode(), b"\r\n\r\n", data, b"\r\n",
            b"--", boundary.encode(), b"--\r\n",
        ])
        r = self._request("POST", "/upload/drive/v3/files", "上傳", timeout=UPLOAD_TIMEOUT,
                          params={"uploadType": "multipart", "fields": "id,name,size"},
                          headers={"Content-Type": f"multipart/related; boundary={boundary}",
                                   "Content-Length": str(len(body))},
                          data=body)
        return r.json()

    def list_backups(self, parent_id):
        """資料夾內所有 packages-*.db.gz，依檔名（＝時間）由新到舊。"""
        q = f"'{parent_id}' in parents and trashed = false"
        out, token = [], None
        while True:
            params = {"q": q, "fields": "nextPageToken,files(id,name,size,createdTime)",
                      "pageSize": 100, "spaces": "drive"}
            if token:
                params["pageToken"] = token
            body = self._request("GET", "/drive/v3/files", "列備份", params=params).json() or {}
            out.extend(body.get("files") or [])
            token = body.get("nextPageToken")
            if not token:
                break
        out = [f for f in out if str(f.get("name", "")).startswith(FILE_PREFIX)
               and str(f.get("name", "")).endswith(FILE_SUFFIX)]
        out.sort(key=lambda f: f["name"], reverse=True)
        return out

    def delete(self, file_id):
        self._request("DELETE", f"/drive/v3/files/{file_id}", "刪備份")


# ============ 快照 / 驗證 / 壓縮 ============

def snapshot_db(src_path, dst_path):
    """SQLite 官方 backup API：一致性快照，含 WAL 裡尚未 checkpoint 的交易。"""
    src = sqlite3.connect(src_path)
    try:
        src.execute("PRAGMA busy_timeout=10000")
        dst = sqlite3.connect(dst_path)
        try:
            src.backup(dst)
            # 快照會繼承來源的 WAL 標記；改回 DELETE 讓備份檔是單一自足檔案
            # （不會一打開就長出 -wal/-shm；還原時 app 啟動會自己再切 WAL）。
            dst.execute("PRAGMA journal_mode=DELETE")
        finally:
            dst.close()
    finally:
        src.close()


def verify_snapshot(path):
    """回 (ok, reason, rows)。ok 條件：integrity_check == 'ok' 且所有資料表列數合計 > 0。"""
    try:
        conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    except sqlite3.Error as e:
        return False, f"開啟快照失敗：{e}", 0
    try:
        res = conn.execute("PRAGMA integrity_check").fetchall()
        msgs = [r[0] for r in res]
        if msgs != ["ok"]:
            return False, "integrity_check：" + "; ".join(msgs)[:300], 0
        tables = [r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
        rows = 0
        for t in tables:
            rows += conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
        if not tables:
            return False, "快照沒有任何資料表", 0
        if rows <= 0:
            return False, "快照所有資料表皆為空（列數 0）", 0
        return True, "ok", rows
    except sqlite3.Error as e:
        # 壞得夠嚴重時 integrity_check 本身就會拋（database disk image is malformed）
        return False, f"integrity_check 執行失敗（SQLite 錯誤）：{e}", 0
    finally:
        conn.close()


def table_row_count(path):
    """所有資料表列數合計（測試比對用）。"""
    ok, _, rows = verify_snapshot(path)
    return rows if ok else -1


def gzip_file(src_path, dst_path):
    """gzip 壓縮並回傳 (原始 sha256, 壓縮後 bytes 長度)；解壓回讀比對，確保壓縮檔本身沒壞。"""
    h = hashlib.sha256()
    with open(src_path, "rb") as fi, gzip.open(dst_path, "wb", compresslevel=6) as fo:
        while True:
            chunk = fi.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
            fo.write(chunk)
    h2 = hashlib.sha256()
    with gzip.open(dst_path, "rb") as fi:
        while True:
            chunk = fi.read(1024 * 1024)
            if not chunk:
                break
            h2.update(chunk)
    if h.hexdigest() != h2.hexdigest():
        raise GDriveError("gzip 回讀比對失敗（壓縮檔內容與快照不符）")
    return h.hexdigest(), os.path.getsize(dst_path)


def prune_local(local_dir, keep=LOCAL_KEEP):
    """本機只留最新 keep 份（依檔名排序＝時間）。回刪掉的檔名清單。"""
    try:
        names = sorted(
            n for n in os.listdir(local_dir)
            if n.startswith(FILE_PREFIX) and n.endswith(FILE_SUFFIX)
        )
    except OSError:
        return []
    doomed = names[:-keep] if keep > 0 else names
    removed = []
    for n in doomed:
        try:
            os.remove(os.path.join(local_dir, n))
            removed.append(n)
        except OSError:
            pass
    return removed


def prune_drive(client, folder_id, keep=DRIVE_KEEP):
    """Drive 只留最新 keep 份。回 (刪掉的檔名清單, 刪後剩餘清單)。"""
    files = client.list_backups(folder_id)     # 新→舊
    doomed = files[keep:] if keep > 0 else files
    removed = []
    for f in doomed:
        client.delete(f["id"])
        removed.append(f["name"])
    return removed, files[:keep]


# ============ 主流程 ============

def run_backup(db_path, local_dir, client, stored_folder_id=None, now=None,
               keep_drive=DRIVE_KEEP, keep_local=LOCAL_KEEP):
    """
    快照 → 驗證 → gzip → 上傳 Drive → 本機留一份 → 清理。
    永不拋例外；回 report dict：
      success, at, stage, error, filename, size, rows, sha256, drive_file_id,
      drive_count, drive_latest_size, local_count, folder_id, folder_created, elapsed_sec
    stage 走到哪就停在哪（snapshot / verify / gzip / folder / upload / local / prune_drive / prune_local / done）。
    stored_folder_id：呼叫端存在 admin_settings 的資料夾 id；report.folder_id 若不同，呼叫端要存回去。
    """
    t0 = time.time()
    now = now or datetime.now()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"{FILE_PREFIX}{stamp}{FILE_SUFFIX}"
    report = {
        "success": False, "at": now.strftime("%Y-%m-%d %H:%M:%S"), "stage": "start",
        "error": "", "filename": filename, "size": 0, "rows": 0, "sha256": "",
        "drive_file_id": "", "drive_count": None, "drive_latest_size": None,
        "local_count": None, "folder_id": stored_folder_id or "", "folder_created": False,
        "pruned_drive": [], "pruned_local": [], "elapsed_sec": 0,
    }
    tmp_db = tmp_gz = None
    try:
        os.makedirs(local_dir, exist_ok=True)
        tmp_db = os.path.join(local_dir, f".snapshot-{stamp}-{os.getpid()}.db")
        tmp_gz = os.path.join(local_dir, f".{filename}.part")

        report["stage"] = "snapshot"
        if not os.path.exists(db_path):
            raise GDriveError(f"資料庫檔不存在：{db_path}")
        snapshot_db(db_path, tmp_db)

        report["stage"] = "verify"
        ok, reason, rows = verify_snapshot(tmp_db)
        report["rows"] = rows
        if not ok:
            raise GDriveError(f"快照驗證失敗，不上傳：{reason}")

        report["stage"] = "gzip"
        sha, size = gzip_file(tmp_db, tmp_gz)
        report["sha256"], report["size"] = sha, size

        report["stage"] = "folder"
        folder_id = client.ensure_folder(stored_folder_id)
        report["folder_id"] = folder_id
        report["folder_created"] = (folder_id != (stored_folder_id or ""))

        report["stage"] = "upload"
        with open(tmp_gz, "rb") as f:
            data = f.read()
        up = client.upload(filename, data, folder_id)
        report["drive_file_id"] = up.get("id", "")
        try:
            up_size = int(up.get("size") or 0)
        except (ValueError, TypeError):
            up_size = 0
        if up_size and up_size != size:
            raise GDriveError(f"上傳大小不符：本機 {size} bytes、Drive 回 {up_size} bytes")

        # 上傳成功後才把本機檔案轉正（失敗的嘗試不留半成品）
        report["stage"] = "local"
        final_local = os.path.join(local_dir, filename)
        os.replace(tmp_gz, final_local)
        tmp_gz = None
        report["success"] = True      # 主要目標（離開機器）已達成；清理失敗只記 warning

        report["stage"] = "prune_drive"
        try:
            removed, remain = prune_drive(client, folder_id, keep_drive)
            report["pruned_drive"] = removed
            report["drive_count"] = len(remain)
            if remain:
                try:
                    report["drive_latest_size"] = int(remain[0].get("size") or 0)
                except (ValueError, TypeError):
                    report["drive_latest_size"] = None
        except Exception as e:      # noqa: BLE001 — 清理失敗不影響備份成功
            report["warning"] = f"Drive 清理失敗：{e}"

        report["stage"] = "prune_local"
        try:
            report["pruned_local"] = prune_local(local_dir, keep_local)
            report["local_count"] = len([n for n in os.listdir(local_dir)
                                         if n.startswith(FILE_PREFIX) and n.endswith(FILE_SUFFIX)])
        except Exception as e:      # noqa: BLE001
            report["warning"] = (report.get("warning", "") + f" 本機清理失敗：{e}").strip()
        report["stage"] = "done"
    except Exception as e:      # noqa: BLE001 — 任何失敗都只記錄，不能影響主程式
        report["success"] = False
        report["error"] = f"[{report['stage']}] {e}"
    finally:
        for p in (tmp_db, tmp_gz):
            if p:
                for suffix in ("", "-wal", "-shm", "-journal"):
                    try:
                        os.remove(p + suffix)
                    except OSError:
                        pass
        report["elapsed_sec"] = round(time.time() - t0, 2)
    return report

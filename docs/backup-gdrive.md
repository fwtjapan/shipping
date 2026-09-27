# SQLite 每日自動備份 → Google Drive

2026-09-18 Zeabur volume 被誤刪，`/data/packages.db` 全部遺失、無法復原。
前一晚有手動下載備份但沒留住——問題不在忘記，在於備份靠人。
這套機制讓備份自動、每天、離開那台機器。**存在同一顆磁碟上的備份已證明沒用。**

## 一次性設定（本機做，約 10 分鐘）

1. Google Cloud Console（用要放備份的那個 Gmail）→ 新專案 → 啟用 **Google Drive API**。
2. OAuth 同意畫面：External。測試模式的 refresh token **7 天會失效**，所以要按「發布」
   （只申請 `drive.file` 這種非敏感 scope，發布不需要 Google 審查）。
3. 憑證 → OAuth 用戶端 ID → **電腦版應用程式** → 記下 client id / secret。
4. 本機跑一次（只用標準函式庫、不用 pip 裝東西；伺服器不需要）：
   ```powershell
   $env:GDRIVE_CLIENT_ID = "xxxx.apps.googleusercontent.com"
   $env:GDRIVE_CLIENT_SECRET = "GOCSPX-xxxx"     # 或不設，執行時會問（不回顯）
   python tools\get_gdrive_token.py
   ```
   本機監聽 `127.0.0.1:3457/callback`（避開 Shopify 那支的 3456）。瀏覽器選帳號、同意後，
   腳本會**立刻用拿到的 refresh token 打一次 Drive API 印出帳號 email**——先確認是 你的 Google 帳號
   再貼到 Zeabur；不對就到 https://myaccount.google.com/permissions 移除存取權重跑。
   token 只印在終端機，不會寫進任何檔案。
5. Zeabur 環境變數填三個：`GDRIVE_CLIENT_ID` / `GDRIVE_CLIENT_SECRET` / `GDRIVE_REFRESH_TOKEN`，redeploy。
6. 後台首頁「💾 資料庫自動備份」卡 → 按「立即備份」→ 看到「最後備份：時間」變綠就完成。
   Google Drive 會多一個 `helpshipping-backups` 資料夾（程式自己建的，id 存在 admin_settings）。

三個變數任一為空 → 備份停用、啟動時印一次警告、其他功能完全不受影響；後台首頁那張卡會紅字提醒。

## 為什麼這樣做

| 決定 | 原因 |
|---|---|
| OAuth2 refresh token，不用 service account | service account **沒有儲存配額**，上傳必回 403 `storageQuotaExceeded`，分享資料夾給它也一樣（Google 平台限制）。一般 Gmail 沒有共用雲端硬碟可用。 |
| scope 只用 `drive.file` | 只能碰程式自己建的檔案。token 外洩損失範圍＝備份檔，不是整個雲端硬碟。 |
| 不裝 google-api-python-client | 只需要兩支 REST endpoint（換 token、multipart 上傳），`requests` 直接打。少一個相依、少一個會壞的地方。 |
| SQLite backup API，不用 `shutil.copy2` | DB 是 WAL 模式，copy2 只複份主檔，最近的交易還在 `-wal` 裡→備份缺資料但照樣開得起來。`test_backup.py` 第 4 項有反例證明。 |
| 驗證不過就不上傳 | `PRAGMA integrity_check` 必須 `ok` 且列數 > 0。壞掉的備份比沒備份危險：你會以為自己有備份。 |
| 多 worker 用 admin_settings 租約 | gunicorn `--workers 2`，記憶體變數擋不住；與會員同步同一套條件式 UPDATE。 |

## 運作

- 每日一次，`admin_settings.backup_hour`（預設 3，**容器本地時間**；容器若是 UTC 就是台灣 11:00）。後台卡片老闆可改。
- 失敗當天每 30 分鐘重試；成功後當天不再跑。容器重啟不會重置（看的是 DB 裡的時間戳）。
- 流程：backup API 快照 → integrity_check → gzip → 上傳 Drive（`packages-YYYYMMDD_HHMMSS.db.gz`）→ `/data/backups/` 留一份 → Drive 留 30 份、本機留 7 份。
- 後台首頁：「最後備份：YYYY-MM-DD HH:MM」。**超過 48 小時沒成功 → 整張卡標紅 + 警告 + 最後失敗原因**。
  備份最常見的失敗不是沒設，是設了之後某天悄悄壞掉；沒有這個顯示等於沒做。
- 端點：`POST /api/admin/maintenance/backup_now`（老闆）、`GET /api/admin/maintenance/backup_status`（老闆+員工）、`PUT /api/admin/settings/backup_hour`（老闆）。

## 還原（手動，刻意不做成按鈕）

還原需要人判斷「要回到哪個時間點、現在的資料要不要保留」，不該自動化。

1. Drive `helpshipping-backups` 下載要還原的 `packages-….db.gz`（或 `/data/backups/` 裡的）。
2. 解壓：`python -c "import gzip,shutil;shutil.copyfileobj(gzip.open('packages-….db.gz','rb'),open('packages.db','wb'))"`
3. 本機先開起來看：`sqlite3 packages.db "PRAGMA integrity_check; SELECT COUNT(*) FROM packages;"`
4. 停服務 → 把現在的 `/data/packages.db`（含 `-wal`/`-shm`）改名留著 → 放進還原檔 → 啟動（app 會自己切回 WAL）。

## 測試

```
python test_backup.py
```
假的 Drive server 攔所有請求，不碰真 Google。涵蓋：憑證為空、端到端、WAL 鑑別、integrity 失敗、401/403/逾時、清理 30/7、兩 worker 租約、權限、後台 48 小時標紅。

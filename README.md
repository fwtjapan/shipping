# FWT JAPAN 雲端集運系統

由 helpshipping 改寫的 FWT JAPAN 集運程式（Flask + SQLite）。

## 快速開始（本機）

1. `pip install -r requirements.txt`
2. 複製 `.env.example` 成 `.env`，至少填 `ADMIN_PASSWORD`、`SESSION_SECRET`
3. `python app.py` → 會員頁 http://localhost:5001/ 、後台 http://localhost:5001/admin

## 品牌與設定

- 所有對客人顯示的品牌資訊（名稱、倉庫地址、電話、LINE、匯款帳號、保險連結）集中在 `brand.py`，全部用環境變數覆寫，清單見 `.env.example`。
- 會員編號存在 Shopify 客戶 metafield `custom.<MEMBER_METAFIELD_KEY>`（預設 `shipping_member_id`）。
- 代理商模組（分潤、代理品牌頁）預設關閉；要開啟設 `ENABLE_AGENTS=1`。

---

# 銷帳檔對帳模組（helpshipping）

每月把銀行的銷帳檔丟進來，自動跟系統帳單對帳，抓出末五碼登錄錯誤、短收、溢收、查無入帳。

## 檔案

```
recon/
  parser.py    銷帳檔（固定長度）解析
  matcher.py   三輪比對引擎
  db.py        從 SQLite 撈帳單  ← 只有這支要照實際 schema 改
  report.py    文字報表 + Excel 匯出
  routes.py    Flask blueprint（上傳頁 /recon）
  cli.py       命令列 / 排程用
templates/
  recon.html
```

## 接線（三個步驟）

**1. 確認資料表結構**

```bash
HELPSHIPPING_DB=helpshipping.db python -m recon.cli --schema
```

**2. 改 `recon/db.py` 的 `BILL_SQL`**

目前假設是 `bills` 資料表，欄位 `customer_code / customer_name / ship_date / total / pay_mark / paid_at`。
只要 `AS` 後面的別名維持不變，SQL 本體可以隨便改（join、view 都行）。
`pay_mark` 對應帳單畫面上「匯款」欄的原始值 —— 末五碼或「現金 / 後付 / 管確認 / 收日幣現金」。

**3. app.py 註冊 blueprint**

```python
from recon.routes import bp as recon_bp
app.register_blueprint(recon_bp)
app.config["HELPSHIPPING_DB"] = "helpshipping.db"
```

`routes.py` 的 `_login_required()` 目前吃 `session['admin']` 或 `session['user_id']`，
請改成 helpshipping 現有的驗證方式。

環境變數：`HELPSHIPPING_DB`、`OWN_ACCOUNT`（預設沿用 `BRAND_BANK_ACCOUNT`，用來偵測末五碼誤填成自家帳號片段）。

## 用法

網頁：`/recon` 上傳 → 報表 → 下載 Excel（六個工作表）。

命令列：

```bash
python -m recon.cli 銷帳檔.txt --start 2026-07-01 --end 2026-07-31 --excel 7月對帳.xlsx
```

exit code：有「短收」或「查無入帳」時回傳 1，方便排程觸發告警。

```
0 8 1 * *  cd /app && python -m recon.cli /data/銷帳檔.txt \
           --start $(date -d 'last month' +%Y-%m-01) \
           --end $(date -d 'this month -1 day' +%Y-%m-%d) || notify.sh
```

## 比對邏輯

1. 帳單依「匯款」欄分流；現金 / 後付 / 管確認 / 收日幣現金不進銀行，直接排除。
2. 以**末五碼分群加總**比對 —— 同客戶當月多筆、或一筆拆兩次匯款都吃得下。
3. 群組金額不符 → 短收 / 溢收。
4. 帳單孤兒 vs 銀行孤兒，用**金額 + 日期容差 7 天**配對 → 判定為末五碼登錄錯誤，
   並細分「他人代匯」「末五碼疑為自家帳號片段」。
5. 「查無入帳」再跟溢收群組比一次，末五碼**編輯距離 ≤ 1** 且金額吻合 → 判定打錯一碼。
6. 剩下的才是真的沒收到。

`is_customer_payment` 會把利息（交易代碼 982）與摘要含 NHI／健保／信用卡／退稅／國稅的
歸到「非運費收入」，不影響對帳差額。

## 銷帳檔格式

```
123456789012000000000000010  000000002+0000000001843+000000005623000(013)0000699***411414 王＊明   20260630
|------- 27 碼主帳號 -------|  |-9碼序號-|+|--13碼金額--|+|--15碼餘額--||行代號||-匯款人帳號-| |匯款人| |-日期-|
```

* 金額單位是**元**，餘額單位是**分**（已驗證：56,230.00 − 54,387.00 = 1,843）
* 主帳號末 3 碼是交易代碼：010 / 040 一般匯入、007 / 060 其他、982 利息
* 匯款人帳號可能遮罩成 `0000699***411414`，取數字末 5 碼即系統上的末五碼

## 相依

`openpyxl`（只有 Excel 匯出用到），其餘標準函式庫。

## 已知限制

* `routes.py` 用 module-level dict 暫存上一次結果，多 worker 部署要改存檔案或 SQLite。
* 「他人代匯」的判定只比對姓名首字，遮罩姓名（`戴＊育`）誤判率不低，僅供參考。
* 同一客戶當月有兩筆金額完全相同、又都登錄錯末五碼時，第二輪可能配錯對象，
  報表上會列出來讓人工複核。

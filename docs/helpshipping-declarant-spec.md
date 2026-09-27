# helpshipping 申報人（多報關人）功能規格

Repo: `fwtjapan/shipping` / main
產出日：2026-09-09

---

## 0. 背景與判斷依據（先讀，不要跳過）

### 為什麼要做
部分客戶單次出貨量大，需要**拆成多份報單、分給多位申報人**。台灣快遞進口以提單分號為單位、**一箱一份簡易申報單**，所以「一個申報人」的粒度是**箱**，不是整張出貨申請、更不是地址。

### 現況（已查證）
- `vendors.py` 兩個範本 **早就有申報人三欄**，但寫死等於收件人：
  - `NIGEL` L126-128：`申報人 / 申報人詳細地址 / 申報人電話號碼` → `ctx["ship_recipient"] / ship_address / ship_phone`
  - `JPD` L161-164：同上，另有 `申報人身份證ID` → 固定 `""`
- 歷史出檔案 1,241 筆，申報人姓名 100% 等於收件人（就是被程式複製的結果）。
- `build_rows`（`vendors.py:206-275`）已經是**逐箱**組 ctx（`for bi, b in enumerate(boxes)`），box 層級欄位（`box_tracking / box_actual_weight / box_length...`）都在 ctx 裡 → 加申報人是同一個模式，低風險。
- `boxes_json` 目前每箱 key：`actual_weight, length, width, height, tracking_num, billed_weight`（`app.py:498` 註解、`admin.html:5341-5352 readBoxInputs`）。

### 法規約束（影響欄位設計，不要自作主張加欄位）
- 依空運快遞貨物通關辦法：**以進口簡易申報單申報收貨人「實名認證行動通訊門號」者，得免申報身分證統一編號**。關務署亦明示完成 EZ WAY 實名認證後不需再提供身分證號給物流業者。
  → **絕對不要收集、不要儲存身分證字號。** `JPD` 的「申報人身份證ID」維持輸出空字串。
- 申報人 = 報單收貨人 = **納稅義務人**（海關的稅單開給他、EZ WAY 推播給他、扣他的半年進口次數）。
- 2026/03/01 起全面「預先確認委任」：每位申報人都要自己在 EZ WAY 按確認，未確認則該箱無法報關放行。

### 名詞定義（實作時的口徑，UI 文案也照這個寫）
| | 收件人 (ship_recipient) | 申報人 (declarant) |
|---|---|---|
| 意義 | 台灣宅配實際簽收的人 | 報單收貨人／納稅義務人 |
| 歸屬 | 地址的屬性 | **人**的屬性，與地址無關 |
| 電話 | 宅配聯絡電話 | **必須是該人 EZ WAY 實名認證綁定的門號**（可能與宅配電話不同支） |
| 粒度 | 一張出貨申請一個 | **一箱一個** |

---

## 1. 範圍

### 做
1. 新表 `declarants`（會員層級申報人清單，1..N）
2. 客戶端：獨立的「申報人管理」區塊（與地址簿並列，**不是**地址簿的欄位）
3. 出貨申請：客戶勾選本次授權使用的申報人（可多選）
4. 後台出貨作業：每箱一個申報人下拉
5. `vendors.py`：兩個範本改讀真實申報人，舊資料 fallback 回收件人

### 不做（明確排除）
- **不動 `addresses` 表**，不加任何欄位
- **不收身分證字號**
- **不動任何計費邏輯**（申報人與運費、材積、理貨費完全無關）
- 半年進口次數計數／提醒 → 留 Phase 2

---

## 2. Schema

### 2.1 新表（放在 `app.py` `init_db()` 內，接在 `addresses` 表定義之後，約 L409）

```sql
CREATE TABLE IF NOT EXISTS declarants (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    g_code      TEXT    NOT NULL,
    name        TEXT    NOT NULL,
    phone       TEXT    NOT NULL,   -- EZ WAY 實名認證綁定門號，正規化為 09xxxxxxxx
    address     TEXT    NOT NULL,
    is_default  INTEGER DEFAULT 0,
    created_at  TEXT    NOT NULL
)
```

建索引：`CREATE INDEX IF NOT EXISTS idx_declarants_gcode ON declarants(g_code)`

### 2.2 `shipment_requests` 新增欄位

用既有的 ALTER 清單機制（`app.py:493` 那組 `(欄位, 型別, 預設)`）追加：

```python
("declarant_name",    "TEXT", "''"),
("declarant_phone",   "TEXT", "''"),
("declarant_address", "TEXT", "''"),
("declarants_json",   "TEXT", "''"),   # 本次授權可用的申報人快照 [{name,phone,address}]
```

前三欄 = **主申報人快照**（後台每箱未指定時的預設值）。

### 2.3 `boxes_json` 每箱新增 key

```
declarant_name, declarant_phone, declarant_address
```

**存快照字串，不存 declarant_id。** 理由與現有 `ship_recipient` 一致：客戶事後改申報人清單，不能回溯汙染已報關的歷史單。舊資料沒有這些 key → 由 fallback 處理（見 §5）。

---

## 3. 後端 API（`app.py`）

### 3.1 申報人 CRUD
完全比照既有 `/api/addresses`（`app.py:3873-3985`）的寫法與歸屬檢查：

| Method | Path | 說明 |
|---|---|---|
| GET | `/api/declarants?g_code=` | 依 `is_default DESC, id DESC` 排序 |
| POST | `/api/declarants` | 新增；第一筆自動 `is_default=1` |
| PUT | `/api/declarants/<id>` | 修改；先查 `g_code` 歸屬 |
| DELETE | `/api/declarants/<id>` | 刪除；若刪掉的是預設，把最舊一筆補成預設 |
| POST | `/api/declarants/<id>/default` | 設為預設（先全部歸零再設 1） |

驗證規則：
- `name / phone / address` 皆必填
- **`phone` 正規化 + 驗證**：去掉空白與 `-`；`+886 9xxxxxxxx` / `8869xxxxxxxx` → `09xxxxxxxx`；最終必須符合 `^09\d{8}$`，否則回 `{"success": false, "error": "申報人電話必須是台灣手機門號（09 開頭 10 碼），且須為 EZ WAY 實名認證綁定的號碼"}`
- 每個 `g_code` **上限 10 筆**，超過回錯誤
- 同一 `g_code` 下 `phone` 重複 → 擋下（一個門號只能對應一個 EZ WAY 帳號）

### 3.2 出貨申請 `/api/shipment-request`（`app.py` 約 L4110-4260）

新增接收 `declarant_ids: [int]`：
1. 以 `g_code` 為條件從 `declarants` 撈出對應資料（**不信任前端傳來的姓名電話**，只信 id）
2. 撈到的清單 → `declarants_json`（JSON 陣列，`ensure_ascii=False`）
3. 主申報人 = 清單中 `is_default=1` 者，無則取第一筆 → 寫入 `declarant_name/phone/address`
4. **未傳 `declarant_ids` 或撈不到** → 三欄與 `declarants_json` 皆留空（不擋下申請），由 §5 fallback 成收件人，維持現況行為
5. `declarant_ids` 若含不屬於此 `g_code` 的 id → 直接忽略該筆（不報錯）

同時在 INSERT 語句補上這四個欄位（`app.py:4247`）。

### 3.3 後台計費／出貨儲存（`app.py` 約 L5061）

`boxes_json = json.dumps(data.get("boxes", []), ...)` 這行不用改，但要確保後台傳來的 box 物件帶著三個 declarant key（前端負責，見 §4.2）。

### 3.4 出貨申請讀取端點

`app.py:4791` 附近（`"boxes": _parse_boxes(...)`）與 L4631/4764 兩處組 dict 的地方，把 `declarant_name/phone/address/declarants_json` 一併帶出來給前端與 `vendors.build_rows`。

---

## 4. 前端

### 4.1 客戶端 `templates/index.html`

**新增「申報人管理」區塊**，放在地址簿 modal 旁邊或同一 modal 的第二個分頁。UI 結構直接複製地址簿（L1403-1495 那組 render / 新增 / 設預設 / 刪除），欄位換成：

- 申報人姓名 *
- EZ WAY 綁定手機 *（placeholder：`09xxxxxxxx，須為此人 EZ WAY 實名認證的門號`）
- 申報人地址 *
- 設為預設申報人

**必要文案（照抄，不要自行潤飾成更軟的說法）：**

> 申報人＝報單上的納稅義務人。海關的稅單、EZ WAY 的確認推播都會發給他，並占用他的半年進口次數。
> 電話必須是該申報人本人 EZ WAY 實名認證綁定的手機門號，填錯將收不到推播、包裹會卡關。
> 自 2026/03/01 起，每位申報人都必須自己在 EZ WAY 按下「申報相符」，該箱才能放行。
> 本欄不需要、也請勿提供身分證字號。

**出貨申請流程**：在選寄送地址之後加一段「本次申報人」多選（checkbox 清單，來源 `/api/declarants`）。
- 預設勾選預設申報人
- 送出時附 `declarant_ids`
- 勾選超過 1 人時，顯示提示：`您選了 N 位申報人，出貨時我們會依箱數分配，每位都需各自在 EZ WAY 確認。`
- 加一個**必勾**的聲明 checkbox：`我確認以上申報人皆為本人或已取得其同意，且各自已完成 EZ WAY 實名認證。`（未勾不可送出）

### 4.2 後台 `templates/admin.html`

改 `addBox` / `readBoxInputs` / `renderBoxes`（L5330-5372）：

- `_billBoxes` 每筆多帶 `declarant_name / declarant_phone / declarant_address`
- `renderBoxes` 每箱那排加一個 `<select class="bx-dec">`：
  - options 來源 = 當前申請的 `declarants_json`（載入計費 modal 時存進一個 `_reqDeclarants` 變數）
  - 第一個 option 為 `（同主申報人）`，value 空字串
  - `declarants_json` 為空時，整個 select 隱藏（舊單／單一申報人的情境不干擾現有作業流程）
- `readBoxInputs` 從 select 的 selectedIndex 反查 `_reqDeclarants`，把**三個字串**寫進 box 物件（不是 id）
- `calcBoxes` 不受影響（申報人不參與任何計算）

**i18n**：新增的所有文字要進 `I18N.ja`（`admin.html` 約 L1468 起的區塊），靜態文字用 `data-i18n`，JS 動態產生的用 `T('中文原字')`。key 不要動既有的。建議譯法：申報人＝`申告者`、同主申報人＝`主申告者と同じ`。

---

## 5. `vendors.py`

### 5.1 `build_rows` ctx 加三個 key（約 L252-270，在 `for bi, b in enumerate(boxes)` 迴圈內）

```python
"declarant_name":    (b.get("declarant_name")    or s.get("declarant_name")    or ctx_ship_recipient),
"declarant_address": (b.get("declarant_address") or s.get("declarant_address") or ctx_ship_address),
"declarant_phone":   (b.get("declarant_phone")   or s.get("declarant_phone")   or ctx_ship_phone),
```

三層 fallback：**箱層級 → 出貨申請主申報人 → 收件人**。第三層保證舊資料出檔案結果與現在完全一致。

### 5.2 改六個 lambda

- `NIGEL` L126-128 → `ctx["declarant_name"] / ctx["declarant_address"] / ctx["declarant_phone"]`
- `JPD` L161/163/164 → 同上
- `JPD` L162「申報人身份證ID」→ **維持 `""`，不要動**

---

## 6. 驗收清單（push 前逐條跑過）

1. **語法**
   - `python3 -c "import ast; ast.parse(open('app.py').read())"`
   - `python3 -c "import ast; ast.parse(open('vendors.py').read())"`
   - `admin.html` 有兩個 `<script>` 區塊 → 用 `re.findall` 全抓合併後 `node --check`
   - `index.html` 同樣要驗
2. **回歸（最重要）**：取一張舊的已出貨單（`boxes_json` 無 declarant key、`declarant_*` 皆空），跑 Nigel 與 JpD 出檔案，**逐欄比對改動前後的輸出**，必須完全相同。
3. **新流程**：建一張三箱、兩位申報人的測試單 → 後台第 1 箱指定 A、第 2 箱指定 B、第 3 箱留空 → 出檔案應為 A / B / 主申報人。
4. **CRUD**：新增 3 位申報人、切換預設（確認同時只有一筆 `is_default=1`）、刪除預設後有新預設遞補。
5. **越權**：用 g_code X 的 session 去 PUT/DELETE 屬於 g_code Y 的 declarant → 必須被擋。
6. **電話驗證**：`0912345678` 通過；`+886912345678` 正規化通過；`0212345678`、`09123456` 被擋。
7. **上限**：第 11 筆被擋。
8. 整合測試照慣例：`DB_PATH=/tmp/x.db`、`os.environ['SHOPIFY_STORE']=''`、測試寫成 `.py` 檔跑（heredoc 在此環境不穩）、結果用 `sys.stderr.write` 輸出。

---

## 7. commit 與部署

- 分兩個 commit：① 後端（schema + API + vendors）② 前端（index.html + admin.html + i18n）
- commit message 要寫清楚：新增 declarants 表、boxes_json 新 key、三層 fallback 保證舊單輸出不變
- push 後 Zeabur 自動 redeploy 約 1-2 分鐘；DB 在 `/data/packages.db` 持久 volume，新表由 `init_db()` 自動建立，不需手動遷移

---

## 8. Phase 2（本次不做，記錄待辦）

- 半年進口次數計數與提醒（依申報人統計近 6 個月出貨箱數，接近 6 次時在後台標紅）
- 地址簿每筆綁定預設申報人（`addresses.default_declarant_id`）
- 後台批次指派（20 箱平均分配給 N 位申報人的一鍵功能）
- 申報人同意聲明的稽核紀錄（寫入 `operation_logs`）

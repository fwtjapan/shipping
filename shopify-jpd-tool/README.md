# 🏯 御用達 × JPD 雲倉 串接工具

Shopify 訂單管理 & JPD 物流串接的網頁工具

## 📋 功能

- ✅ 查看 Shopify 待出貨訂單
- ✅ 一鍵創建 JPD 運單（支援自出貨 / 倉庫代發兩種模式）
- ✅ 查看 JPD 運單狀態
- ✅ 確認發貨 / 取消訂單
- ✅ 自動回寫 Shopify 出貨資訊

## 🚀 部署到 Zeabur

### 1. 上傳到 GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/你的帳號/shopify-jpd-tool.git
git push -u origin main
```

### 2. Zeabur 部署

1. 登入 [Zeabur](https://zeabur.com)
2. 建立新服務 → Deploy from GitHub
3. 選擇此 Repository

### 3. 設定環境變數

在 Zeabur 的 Variables 中設定：

| 變數名稱 | 說明 | 範例 |
|----------|------|------|
| `SHOPIFY_STORE` | Shopify 店舖名稱 | `fd249b-ba` |
| `SHOPIFY_ACCESS_TOKEN` | Admin API Token | `shpca_xxxxxx` |
| `JPD_EMAIL` | JPD 登入郵箱 | `you@example.com` |
| `JPD_PASSWORD` | JPD 登入密碼 | `yourpassword` |
| `JPD_BASE_URL` | JPD API 網址 | `https://biz.cloudwh.jp` |
| `JPD_WAREHOUSE_ID` | 倉庫 ID | `1` |
| `JPD_DELIV_ID` | 物流方式 ID | `40` |

### 4. 設定網域

部署完成後在 Networking 新增網域即可使用。

## 📁 檔案結構

```
shopify-jpd-tool/
├── app.py              # 主程式
├── requirements.txt    # 依賴套件
├── Procfile            # Zeabur 啟動指令
├── runtime.txt         # Python 版本
├── .env.example        # 環境變數範例
├── .gitignore          # Git 忽略清單
├── README.md           # 說明文件
└── templates/
    └── index.html      # 網頁介面
```

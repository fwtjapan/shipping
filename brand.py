"""FWT JAPAN 集運品牌設定

所有對客人顯示的品牌資訊集中在這裡，全部可用環境變數覆寫（部署時在 Zeabur 設定即可，不用改程式）。
模板裡用 {{ brand.xxx }} 取值；app.py 透過 context_processor 注入。
"""
import os


def _env(key, default=""):
    # 空白值視同沒設定 → 用預設（.env 裡留 KEY= 不會蓋掉預設）
    return (os.environ.get(key) or "").strip() or default


BRAND = {
    # 品牌名稱 / 公司
    "name":          _env("BRAND_NAME", "FWT JAPAN"),
    "company":       _env("BRAND_COMPANY", "FWTJAPAN株式会社"),
    "site_url":      _env("BRAND_SITE_URL", "https://fwtjapan.com"),
    # 客人申請會員的網址（Shopify 註冊頁）
    "signup_url":    _env("BRAND_SIGNUP_URL", "https://fwtjapan.com/account/register"),
    # 集運系統對外網址（後台「複製登入資訊」會用到）；空白 = 用目前瀏覽器網址
    "app_url":       _env("BRAND_APP_URL", ""),

    # 客服 LINE（例：@abc1234）；空白 = 不顯示 LINE 連結
    "line_id":       _env("BRAND_LINE_ID", "@fwtjapan"),

    # 集運倉地址（客人寄包裹用）
    "wh_zip":        _env("BRAND_WAREHOUSE_ZIP", "〒130-0022"),
    "wh_address":    _env("BRAND_WAREHOUSE_ADDRESS", "東京都墨田区江東橋4丁目21-6 錦糸町ハイタウン208室"),
    "wh_phone":      _env("BRAND_WAREHOUSE_PHONE", "070-9244-8894"),
    "wh_map_url":    _env("BRAND_WAREHOUSE_MAP_URL", ""),

    # 匯款資訊（客人付運費用；會員頁「匯款資訊」區塊顯示）
    "bank_name":     _env("BRAND_BANK_NAME", "(807) 永豐銀行 營業部"),
    "bank_account":  _env("BRAND_BANK_ACCOUNT", "20101800120496"),
    "bank_holder":   _env("BRAND_BANK_HOLDER", "潘建綱"),

    # 貨物保險商品（空白 = 不顯示保險區塊的連結）
    "insurance_url":   _env("BRAND_INSURANCE_URL", ""),
    "insurance_label": _env("BRAND_INSURANCE_LABEL", "立即加購貨物保險"),
    "insurance_desc":  _env("BRAND_INSURANCE_DESC", "貨物保障，每箱上限 5 萬台幣"),
}

# 收費（會員頁收費說明用；實際出帳費率仍以後台/會員設定為準）
BRAND["rate_twd"] = int(_env("DEFAULT_SHIPPING_RATE", "180"))          # 每公斤台幣
BRAND["rate_jpy"] = round(BRAND["rate_twd"] * float(_env("TWD_TO_JPY_RATE", "5.0")))
BRAND["free_storage_days"] = int(_env("FREE_STORAGE_DAYS", "30"))
BRAND["tier_kg"] = int(_env("TIER_NEGOTIATE_KG", "300"))                # 每月超過多少 kg 分層談價
BRAND["fragile_jpy"] = int(_env("FRAGILE_PACKING_JPY", "300"))          # 易碎品包裝 每個日幣

# 檔名用（出貨明細 Excel 等）：去掉空白
BRAND["slug"] = BRAND["name"].replace(" ", "")
# LINE 連結
BRAND["line_url"] = ("https://line.me/R/ti/p/" + BRAND["line_id"]) if BRAND["line_id"] else ""

# Shopify 客戶 metafield（custom.<key>）存放集運會員編號
MEMBER_METAFIELD_KEY = _env("MEMBER_METAFIELD_KEY", "shipping_member_id")

# 代理商模組（分潤、代理品牌 referral 頁）；預設關閉
ENABLE_AGENTS = _env("ENABLE_AGENTS", "0").lower() in ("1", "true", "yes", "on")

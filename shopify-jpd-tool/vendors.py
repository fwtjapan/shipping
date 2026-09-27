"""
廠商出貨檔案範本（config-driven）— shopify-jpd-tool 專用版本

跟 helpshipping 的 vendors.py 同一套設計，但兩邊資料模型不同：
  - helpshipping（集運）：one_per_package + 亂數品名（倉庫只核對、不知道每箱裝什麼）
  - 本檔（Shopify）：one_per_item + 真實品名/數量/金額（資料直接來自 Shopify 訂單品項）

輸出規則（依用戶 2026-06 需求）：
  1. 品項全部輸出：items_json 每個商品一列，用真實 product_name_local / product_num / product_price
  2. 同一訂單（同一箱）的前段欄位（客戶編號/運單號、清關號碼/JpD包裹ID、收件人、
     地址、電話、申報人…）全部相同；不同箱 = 不同 order_history 列，自然有不同號
  3. 客戶編號（Nigel）/ 客戶運單號（JpD）= 注文番號（customer_order_id，去掉 # 與前後空白）
  4. 清關號碼（Nigel）/ JpD包裹ID（JpD）= 留空，由 Nigel / JpD 端工作人員自行填入

新增廠商只要加一份 dict 到 VENDORS 即可，不用改其他 code。
"""
from __future__ import annotations
import json
import random
from datetime import datetime


# ============ 產地隨機產生器（items 沒有產地欄位，用固定 seed 保持每次匯出一致） ============

_ORIGIN_POOL = ["japan"] * 6 + ["china"] * 4  # 60% japan, 40% china


def random_origin(order_id, item_index: int = 0) -> str:
    """每筆訂單的每個品項用固定 seed → 同一筆每次匯出結果一樣（避免亂跳）"""
    return random.Random(f"jpd-origin-{order_id}-{item_index}").choice(_ORIGIN_POOL)


# ============ 廠商範本 ============
# getter 接收 ctx dict（每個「商品」一份），回字串/數字

NIGEL = {
    "id": "nigel",
    "display_name": "Nigel",
    "filename_template": "{date}_Nigel_出貨單.xlsx",
    "row_strategy": "one_per_item",
    "columns": [
        ("客戶編號",       lambda ctx: ctx["customer_order_id"]),   # 注文番號
        ("清關號碼",       lambda ctx: ""),                         # 由 Nigel 端工作人員自行填入
        ("收件人",         lambda ctx: ctx["recipient"]),
        ("收件人詳細地址", lambda ctx: ctx["address"]),
        ("收件人電話號碼", lambda ctx: ctx["phone"]),
        ("申報人",         lambda ctx: ctx["recipient"]),           # 預設 = 收件人
        ("申報人詳細地址", lambda ctx: ctx["address"]),
        ("申報人電話號碼", lambda ctx: ctx["phone"]),
        ("品名",           lambda ctx: ctx["product_name"]),
        ("數量",           lambda ctx: ctx["product_num"]),
        ("金額",           lambda ctx: ctx["product_price"]),
        ("產地",           lambda ctx: random_origin(ctx["order_id"], ctx["item_index"])),
        ("URL/JanCode",    lambda ctx: ""),
    ],
}


JPD = {
    "id": "jpd",
    "display_name": "小客戶大價值 (JpD)",
    "filename_template": "{date}_JpD_出貨單.xlsx",
    "row_strategy": "one_per_item",
    "columns": [
        ("客戶運單號",      lambda ctx: ctx["customer_order_id"]),  # 注文番號
        ("JpD包裹ID",       lambda ctx: ""),                        # 由 JpD 端工作人員自行填入
        ("運單ID",          lambda ctx: ""),
        ("包裹特殊服務",     lambda ctx: ""),
        ("收件人",          lambda ctx: ctx["recipient"]),
        ("收件人身份證ID",   lambda ctx: ""),  # 不需要
        ("收件人詳細地址",   lambda ctx: ctx["address"]),
        ("收件人电话号码",   lambda ctx: ctx["phone"]),
        ("備註",            lambda ctx: ""),
        ("特殊服务",         lambda ctx: ""),
        ("渠道ID",          lambda ctx: 40),  # JpD 固定 40
        ("申報人",          lambda ctx: ctx["recipient"]),
        ("申報人身份證ID",   lambda ctx: ""),
        ("申報人詳細地址",   lambda ctx: ctx["address"]),
        ("申報人电话号码",   lambda ctx: ctx["phone"]),
        ("品名",            lambda ctx: ctx["product_name"]),
        ("数量",            lambda ctx: ctx["product_num"]),
        ("金额",            lambda ctx: ctx["product_price"]),
        ("材質",            lambda ctx: ""),
        ("產地",            lambda ctx: random_origin(ctx["order_id"], ctx["item_index"])),
        ("URL/JanCode",     lambda ctx: ""),
    ],
}


VENDORS = {
    "nigel": NIGEL,
    "jpd": JPD,
}


def list_vendors() -> list[dict]:
    """給前端 UI 顯示用的廠商清單"""
    return [{"id": v["id"], "display_name": v["display_name"]} for v in VENDORS.values()]


def get_vendor(vendor_id: str) -> dict | None:
    return VENDORS.get(vendor_id)


def _parse_items(raw) -> list[dict]:
    """解析 order_history.items_json → list[dict]。
    items_json 來源是建單時的 declare_list，每項含：
      product_name / product_name_local / product_num / product_price
    壞掉或空 → 回 []。
    """
    if not raw:
        return []
    try:
        data = json.loads(raw) if isinstance(raw, str) else raw
    except (ValueError, TypeError):
        return []
    return data if isinstance(data, list) else []


def _item_fields(item: dict) -> tuple[str, int, int]:
    """從單一品項 dict 抽出 (品名, 數量, 金額)，容錯處理缺欄/型別。"""
    name = (item.get("product_name_local")
            or item.get("product_name")
            or item.get("title")
            or "商品")
    try:
        num = int(item.get("product_num", item.get("quantity", 1)) or 1)
    except (ValueError, TypeError):
        num = 1
    try:
        price = int(float(item.get("product_price", item.get("price", 0)) or 0))
    except (ValueError, TypeError):
        price = 0
    return str(name).strip(), num, price


def build_rows(vendor_id: str, orders: list[dict]) -> tuple[list[str], list[list]]:
    """
    依範本產生 (headers, rows)。
    orders 結構：每個 dict = 一筆 order_history（含 items_json / logis_num / customer_order_id …）。

    row_strategy = one_per_item：items_json 每個商品一列；同一訂單的前段欄位相同。
    若某訂單沒有任何可解析的品項 → 仍出 1 列（品名留空），避免整筆漏掉。

    Returns:
      headers: List[str] — 第一列標頭
      rows:    List[List] — 資料列
    """
    vendor = get_vendor(vendor_id)
    if not vendor:
        raise ValueError(f"unknown vendor: {vendor_id}")

    headers = [col[0] for col in vendor["columns"]]
    rows = []

    for o in orders:
        # 注文番號：customer_order_id 去掉 # 與前後空白；空的話用本地流水號當 fallback
        coid = (o.get("customer_order_id") or "").strip().lstrip("#").strip()
        if not coid:
            coid = f"ORD{o.get('id', 0)}"

        recipient = (o.get("recipient") or "").strip()
        address   = (o.get("address") or "").strip()
        phone     = (o.get("phone") or "").strip()
        order_id  = o.get("id", 0)

        items = _parse_items(o.get("items_json"))
        if not items:
            items = [{}]  # 沒品項也出 1 列（前段欄位照填，品名/數量/金額留空）

        for idx, item in enumerate(items):
            name, num, price = _item_fields(item)
            ctx = {
                "order_id":          order_id,
                "customer_order_id": coid,
                "recipient":         recipient,
                "address":           address,
                "phone":             phone,
                "item_index":        idx,
                "product_name":      name if item else "",
                "product_num":       num if item else "",
                "product_price":     price if item else "",
            }
            row = [getter(ctx) for _, getter in vendor["columns"]]
            rows.append(row)

    return headers, rows


def filename_for(vendor_id: str, when: datetime | None = None) -> str:
    vendor = get_vendor(vendor_id)
    if not vendor:
        return "export.xlsx"
    when = when or datetime.now()
    return vendor["filename_template"].format(date=when.strftime("%Y%m%d_%H%M%S"))

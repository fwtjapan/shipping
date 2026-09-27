"""
廠商出貨檔案範本（config-driven）

每個廠商一份 dict，描述：
  - id: 內部識別字串（'nigel' / 'jpd'）
  - display_name: UI 顯示用
  - filename_template: 下載檔名格式
  - columns: Excel 欄位定義（header + getter）
  - row_strategy: 'one_per_package'（一個包裹一行）

新增廠商只要加一份 dict 到 VENDORS 即可，不用改其他 code。
"""
from __future__ import annotations
import random
import re
from datetime import datetime


# ============ 隨機資料產生器（出檔案一律用白名單隨機，不使用客戶預報資料） ============

# 唯一允許出現在廠商 Excel 的品名（客戶預報的真實品名不外流）
_PRODUCT_NAME_POOL = [
    "玩具", "糖果", "上衣", "文具用品", "髮夾", "廚房用品", "貴鞋子",
    "化妝品", "娃娃", "包包", "毛巾", "保健食品", "便宜鞋子", "飾品－A",
    "吊飾", "水壺", "襪子", "蝦餅", "口罩", "小朋友涼鞋", "飾品－Ｂ",
    "沐浴球", "卡片",
]
_ORIGIN_POOL = ["japan"] * 6 + ["china"] * 4  # 60% japan, 40% china

MAX_QTY = 10          # 每個品項數量上限
PRICE_CAP = 20000     # 單價上限（JPY）；僅「貴鞋子」不受此限
PREMIUM_NAME = "貴鞋子"

# 重量(kg) → 數量範圍（含上界）；越重件數越多，避免 0.3kg 出現「糖果 9 包」
_QTY_BY_WEIGHT = [
    (0.5, 1, 2),
    (1.0, 1, 3),
    (2.0, 2, 4),
    (3.0, 3, 5),
    (5.0, 4, 7),
    (8.0, 5, 8),
]
_QTY_HEAVY = (6, 10)   # > 8kg


def _seed_for(shipment_id: int, package_id: int = 0) -> random.Random:
    """每筆出貨用固定 seed → 同一筆每次匯出結果一樣（避免亂跳）"""
    return random.Random(f"{shipment_id}-{package_id}")


def _qty_for_weight(rng, kg: float) -> int:
    """依包裹重量決定件數（區間內隨機）。"""
    lo, hi = _QTY_HEAVY
    for limit, a, b in _QTY_BY_WEIGHT:
        if kg <= limit:
            lo, hi = a, b
            break
    return min(rng.randint(lo, hi), MAX_QTY)


def _price_for(rng, name: str, kg: float, qty: int) -> int:
    """單價（JPY）：與重量掛鉤，讓「單價 × 件數」的申報總額與包裹重量大致成正比。

    基準：每 kg 約 ¥1,500~4,500 的申報總額 → 單價 = 總額 / 件數，整百。
    上限：一律 ≤ PRICE_CAP(¥20,000)，只有「貴鞋子」可超過（最高 ¥45,000）。
    下限：¥200。
    """
    kg = max(float(kg or 0), 0.1)
    total = kg * rng.randint(1500, 4500)          # 該包裹的申報總額
    price = int(round(total / max(qty, 1) / 100)) * 100
    price = max(price, 200)
    if name == PREMIUM_NAME:
        return min(price, 45000)                   # 貴鞋子可高於 2 萬
    return min(price, PRICE_CAP)


def build_item_plan(shipment_id: int, weights: list) -> list[dict]:
    """為一筆出貨單配置品項（每個包裹一列）。

    規則：
      • 品名只從白名單取
      • 同一筆出貨單內【品名不重複】→ 不會有「同品名不同價格」；同名必同價
      • 數量依【該包裹重量】決定（0.3kg 只會 1~2 件；上限 10）
      • 單價與重量掛鉤（申報總額合理），除「貴鞋子」外一律 ≤ ¥20,000
      • 固定 seed → 同一筆每次匯出結果一致
    weights: 各包裹重量(kg) 的 list；長度即列數。
    """
    rng = _seed_for(shipment_id, 0)
    pool = list(_PRODUCT_NAME_POOL)
    rng.shuffle(pool)

    plan = []
    price_by_name = {}
    for i, w in enumerate(weights):
        try:
            kg = float(w or 0)
        except (TypeError, ValueError):
            kg = 0.0
        name = pool[i % len(pool)]               # 不重複；超過池長才循環
        qty = _qty_for_weight(rng, kg)
        if name not in price_by_name:            # 同名必同價
            price_by_name[name] = _price_for(rng, name, kg, qty)
        plan.append({
            "name": name,
            "price": price_by_name[name],
            "qty": qty,
            "origin": rng.choice(_ORIGIN_POOL),
        })
    return plan


# ============ 廠商範本 ============

NIGEL = {
    "id": "nigel",
    "display_name": "Nigel",
    "filename_template": "{date}_Nigel_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        # (header, getter — 接收 ctx dict 回字串/數字)
        ("客戶編號",       lambda ctx: f"{ctx['g_code']}-{ctx['packaging_mmdd']}"),
        ("清關號碼",       lambda ctx: ctx["tracking_num"]),  # 出貨追蹤號碼（多箱換行）
        ("收件人",         lambda ctx: ctx["ship_recipient"]),
        ("收件人詳細地址", lambda ctx: ctx["ship_address"]),
        ("收件人電話號碼", lambda ctx: ctx["ship_phone"]),
        ("申報人",         lambda ctx: ctx["declarant_name"]),         # 未指定時 fallback = 收件人
        ("申報人詳細地址", lambda ctx: ctx["declarant_address"]),
        ("申報人電話號碼", lambda ctx: ctx["declarant_phone"]),
        ("品名",           lambda ctx: ctx["item"]["name"]),
        ("數量",           lambda ctx: ctx["item"]["qty"]),
        ("金額",           lambda ctx: ctx["item"]["price"]),
        ("產地",           lambda ctx: ctx["item"]["origin"]),
        ("URL/JanCode",    lambda ctx: ""),
        # ── N~R：每箱資訊（甲案，接在 URL/JanCode 之後）──
        ("清關號碼2",      lambda ctx: ctx.get("box_tracking", "")),      # N = 該箱追蹤號
        ("重量",           lambda ctx: ctx.get("box_actual_weight", "")), # O = 實際重量（不進位）
        ("長",             lambda ctx: ctx.get("box_length", "")),        # P
        ("寬",             lambda ctx: ctx.get("box_width", "")),         # Q
        ("高",             lambda ctx: ctx.get("box_height", "")),        # R
    ],
}


JPD = {
    "id": "jpd",
    "display_name": "JpD 集運（小客戶大價值）",
    "filename_template": "{date}_JpD_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        ("客戶運單號",      lambda ctx: f"{ctx['g_code']}-{ctx['packaging_mmdd']}"),
        ("JpD包裹ID",       lambda ctx: ctx["tracking_num"]),  # 出貨追蹤號碼（多箱換行）
        ("運單ID",          lambda ctx: ""),
        ("包裹特殊服務",     lambda ctx: ""),
        ("收件人",          lambda ctx: ctx["ship_recipient"]),
        ("收件人身份證ID",   lambda ctx: ""),  # 不需要
        ("收件人詳細地址",   lambda ctx: ctx["ship_address"]),
        ("收件人电话号码",   lambda ctx: ctx["ship_phone"]),
        ("備註",            lambda ctx: ""),
        ("特殊服务",         lambda ctx: ""),
        ("渠道ID",          lambda ctx: 40),  # JpD 固定 40
        ("申報人",          lambda ctx: ctx["declarant_name"]),
        ("申報人身份證ID",   lambda ctx: ""),   # EZ WAY 實名認證後免附身分證號 → 永遠留空
        ("申報人詳細地址",   lambda ctx: ctx["declarant_address"]),
        ("申報人电话号码",   lambda ctx: ctx["declarant_phone"]),
        ("品名",            lambda ctx: ctx["item"]["name"]),
        ("数量",            lambda ctx: ctx["item"]["qty"]),
        ("金额",            lambda ctx: ctx["item"]["price"]),
        ("材質",            lambda ctx: ""),
        ("產地",            lambda ctx: ctx["item"]["origin"]),
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


def _packaging_mmdd(s: dict) -> str:
    """打包日 MMDD：updated_at（標記已出貨那刻）> created_at（客戶申請）> 今天。"""
    for src in (s.get("updated_at"), s.get("created_at")):
        if src and len(str(src)) >= 10:
            try:
                return datetime.strptime(str(src)[:10], "%Y-%m-%d").strftime("%m%d")
            except ValueError:
                pass
    return datetime.now().strftime("%m%d")


def export_code_for(s: dict) -> str:
    """出檔案給廠商的客戶編號 = {g_code}-{MMDD}（與 Excel 內容一致，供台灣配送貨況比對）。"""
    return f"{s.get('g_code', '')}-{_packaging_mmdd(s)}"


def build_rows(vendor_id: str, shipments: list[dict]) -> tuple[list[str], list[list]]:
    """
    依範本產生 (headers, rows)。
    shipments 結構：每個 dict 包含 shipment 本身 + 該 shipment 內的 packages list。

    Returns:
      headers: List[str] — 第一列標頭
      rows:    List[List] — 資料列（每個 package 一列）
    """
    vendor = get_vendor(vendor_id)
    if not vendor:
        raise ValueError(f"unknown vendor: {vendor_id}")

    headers = [col[0] for col in vendor["columns"]]
    rows = []

    for s in shipments:
        packaging_mmdd = _packaging_mmdd(s)

        # 整串追蹤號（舊制單箱 fallback 用）
        tracking_all = "\n".join(
            t.strip()
            for t in re.split(r"[\n,，、]+", str(s.get("tracking_num") or ""))
            if t.strip()
        )

        # 決定「箱子」清單：新制讀 boxes；舊制（無 boxes）合成一箱、用出貨層級資料
        boxes = s.get("boxes") or []
        if not boxes:
            boxes = [{
                "tracking_num":  tracking_all,
                "actual_weight": "",     # 舊資料未存實際重量 → 留空（符合「舊的照舊」）
                "length": "", "width": "", "height": "",
                "billed_weight": s.get("billed_weight") or 0,
            }]

        # 各箱重量（給品名/數量配置用）：優先計費重量、其次實際重量、再 fallback 1kg
        box_weights = []
        for b in boxes:
            try:
                w = float(b.get("billed_weight") or 0) or float(b.get("actual_weight") or 0)
            except (TypeError, ValueError):
                w = 0.0
            box_weights.append(w if w > 0 else 1.0)
        item_plan = build_item_plan(s["id"], box_weights)

        for bi, b in enumerate(boxes):
            box_tracking = str(b.get("tracking_num") or "").strip() or tracking_all
            ctx = {
                "shipment_id":          s["id"],
                "g_code":               s["g_code"],
                "packaging_mmdd":       packaging_mmdd,
                "tracking_num":         box_tracking,      # B 欄清關號＝該箱追蹤號
                "item":                 item_plan[bi],
                "pkg_index":            bi,
                "ship_recipient":       str(s["ship_recipient"]) if s["ship_recipient"] else "",
                "ship_address":         str(s["ship_address"]) if s["ship_address"] else "",
                "ship_phone":           str(s["ship_phone"]) if s["ship_phone"] else "",
                "billed_weight":        s.get("billed_weight") or 0,
                "total_fee":            s.get("total_fee") or 0,
                # N~R：該箱資訊
                "box_tracking":         box_tracking,
                "box_actual_weight":    b.get("actual_weight", "") if b.get("actual_weight") not in (None,) else "",
                "box_length":           b.get("length", ""),
                "box_width":            b.get("width", ""),
                "box_height":           b.get("height", ""),
            }
            # 申報人（報單收貨人／納稅義務人）三層 fallback：
            #   ① 該箱指定 → ② 該出貨申請的主申報人 → ③ 收件人（舊資料照舊，輸出與過去完全一致）
            # ⚠️ 設計理由見 docs/helpshipping-declarant-spec.md §5.1：第三層是 1,241 筆歷史單
            #    出檔案結果不變的保證，三層缺一不可，不要自行簡化或改順序。
            ctx["declarant_name"] = (
                str(b.get("declarant_name") or "").strip()
                or str(s.get("declarant_name") or "").strip()
                or ctx["ship_recipient"]
            )
            ctx["declarant_address"] = (
                str(b.get("declarant_address") or "").strip()
                or str(s.get("declarant_address") or "").strip()
                or ctx["ship_address"]
            )
            ctx["declarant_phone"] = (
                str(b.get("declarant_phone") or "").strip()
                or str(s.get("declarant_phone") or "").strip()
                or ctx["ship_phone"]
            )
            row = [getter(ctx) for _, getter in vendor["columns"]]
            rows.append(row)

    return headers, rows


def filename_for(vendor_id: str, when: datetime | None = None) -> str:
    vendor = get_vendor(vendor_id)
    if not vendor:
        return "export.xlsx"
    when = when or datetime.now()
    return vendor["filename_template"].format(date=when.strftime("%Y%m%d_%H%M%S"))

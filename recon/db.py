# -*- coding: utf-8 -*-
"""
帳單來源：從 helpshipping 的 SQLite 撈出指定付款日區間的帳單。

已對照 app.py 實際 schema：
  資料表 shipment_requests
    g_code         客戶編號
    customer_name  客戶姓名
    total_fee      合計
    payment_last5  匯款欄（末五碼，或「現金 / 後付 / 管確認 / 收日幣現金」）
    payment_at     銷帳時間 ← 用來做付款日區間篩選
    updated_at     標記已出貨那天（＝出貨日），舊單 fallback created_at
    agent_id       0 = 主管理員
"""
from __future__ import annotations

import os
import re
import sqlite3
from datetime import date, datetime

from .matcher import BillRecord


def _norm_pay_mark(v) -> str:
    """正規化匯款欄：SQLite 可能把末五碼存成浮點字串（11414.0）→ 還原成 11414。
    非數字（現金/管確認/後付…）原樣保留。"""
    s = str(v or "").strip()
    m = re.fullmatch(r"(\d+)\.0+", s)   # 11414.0 / 11414.00 → 11414
    if m:
        s = m.group(1)
    if s.isdigit():                     # 純數字補回5位，還原被吃掉的前導0（0→00000、123→00123）
        s = s.zfill(5)
    return s

# 與 app.py 一致
DB_PATH = os.environ.get("DB_PATH", "packages.db")

# 自家收款帳號（偵測「末五碼誤填成自家帳號片段」）
# 自家收款帳號：沒設 OWN_ACCOUNT 就沿用會員頁顯示的 BRAND_BANK_ACCOUNT
OWN_ACCOUNT = os.environ.get("OWN_ACCOUNT") or os.environ.get("BRAND_BANK_ACCOUNT", "")

BILL_SQL = """
SELECT
    sr.id                                                 AS bill_id,
    sr.g_code                                             AS customer_code,
    sr.customer_name                                      AS customer_name,
    COALESCE(NULLIF(sr.updated_at, ''), sr.created_at)    AS ship_date,
    sr.total_fee                                          AS amount,
    sr.payment_last5                                      AS pay_mark,
    sr.payment_at                                         AS paid_at
FROM shipment_requests sr
WHERE sr.status = '已出貨'
  AND COALESCE(sr.payment_last5, '') <> ''
  AND (
        date(replace(substr(COALESCE(NULLIF(sr.payment_at,''), sr.updated_at, sr.created_at), 1, 10), '/', '-'))
        BETWEEN date(:start) AND date(:end)
     OR date(replace(substr(COALESCE(NULLIF(sr.updated_at,''), sr.created_at), 1, 10), '/', '-'))
        BETWEEN date(:start) AND date(:end)
      )
  {agent_filter}
ORDER BY sr.payment_at DESC, sr.id DESC
"""


def diagnose(conn, start: date, end: date, agent_id: int = 0) -> dict:
    """回報帳單在各條件下的筆數，用來判斷『0 筆』是卡在哪一關。"""
    af = "AND agent_id = :agent_id" if agent_id > 0 else ""
    p = {"start": start.isoformat(), "end": end.isoformat()}
    if agent_id > 0:
        p["agent_id"] = agent_id
    def q(where):
        return conn.execute(f"SELECT COUNT(*) c FROM shipment_requests sr WHERE {where} {af}", p).fetchone()["c"]
    in_range = (
        "sr.status='已出貨' AND COALESCE(sr.payment_last5,'')<>'' AND ("
        "date(replace(substr(COALESCE(NULLIF(sr.payment_at,''),sr.updated_at,sr.created_at),1,10),'/','-')) BETWEEN date(:start) AND date(:end)"
        " OR date(replace(substr(COALESCE(NULLIF(sr.updated_at,''),sr.created_at),1,10),'/','-')) BETWEEN date(:start) AND date(:end))")
    # 取樣：區間內帳單的 payment_last5 實際長怎樣（判斷是數字末五碼還是文字如『管確認/現金』）
    samples = [str(r["payment_last5"]) for r in conn.execute(
        f"SELECT DISTINCT payment_last5 FROM shipment_requests sr WHERE {in_range} {af} LIMIT 12", p).fetchall()]
    numeric = conn.execute(
        f"SELECT COUNT(*) c FROM shipment_requests sr WHERE {in_range} {af} "
        "AND replace(sr.payment_last5,'.0','') GLOB '[0-9][0-9][0-9][0-9][0-9]' "
        "AND length(replace(sr.payment_last5,'.0',''))=5", p).fetchone()["c"]
    return {
        "已出貨總數": q("sr.status='已出貨'"),
        "有填末五碼": q("sr.status='已出貨' AND COALESCE(sr.payment_last5,'')<>''"),
        "區間內(出貨或付款日)": q(in_range),
        "區間內_末五碼是純5碼數字": numeric,
        "末五碼取樣": samples,
    }


def _to_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, date):
        return v
    s = str(v)[:19].replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def connect(db_path: str | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path or DB_PATH, timeout=5)
    conn.row_factory = sqlite3.Row
    return conn


def fetch_bills(conn, start: date, end: date, agent_id: int = 0,
                sql: str | None = None) -> list[BillRecord]:
    params = {"start": start.isoformat(), "end": end.isoformat()}
    if sql is None:
        sql = BILL_SQL.format(
            agent_filter="AND sr.agent_id = :agent_id" if agent_id > 0 else "")
        if agent_id > 0:
            params["agent_id"] = agent_id

    out = []
    for r in conn.execute(sql, params).fetchall():
        out.append(BillRecord(
            bill_id=str(r["bill_id"]),
            customer_code=str(r["customer_code"] or ""),
            customer_name=str(r["customer_name"] or ""),
            ship_date=_to_date(r["ship_date"]),
            amount=int(round(float(r["amount"] or 0))),
            pay_mark=_norm_pay_mark(r["pay_mark"]),
            paid_at=_to_date(r["paid_at"]),
        ))
    return out


def inspect_schema(conn) -> dict:
    out = {}
    for (name,) in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
        out[name] = [c[1] for c in conn.execute(f"PRAGMA table_info({name})")]
    return out

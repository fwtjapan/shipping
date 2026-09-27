# -*- coding: utf-8 -*-
"""
對帳比對引擎

比對邏輯（依 2026/06 人工對帳過程歸納）：
  1. 帳單依「匯款註記」分流：純數字 = 末五碼（應有銀行入帳）；
     現金／後付／管確認／收日幣現金等 = 不進銀行，直接排除。
  2. 以「末五碼」分群加總比對（同一客戶一個月多筆、或一筆拆兩次匯款都能吃）。
  3. 群組金額不符 → 標記短收 / 溢收。
  4. 帳單有、銀行無的孤兒 → 第二輪用「金額 + 日期容差」去比對銀行孤兒，
     命中即判定為「末五碼登錄錯誤」（實務上最常見：打錯一碼、他人代匯）。
  5. 末五碼若是自家主帳號的片段 → 直接標記為複製貼上錯誤。
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta

# 帳單「匯款」欄非銀行匯款的註記
NON_BANK_MARKS = ("現金", "後付", "管確認", "日幣", "刷卡", "抵扣", "line pay", "linepay")

# 孤兒配對時允許的日期差（天）
DATE_TOLERANCE_DAYS = 7


@dataclass
class BillRecord:
    """一筆帳單（由 helpshipping 資料庫取出）"""
    bill_id: str
    customer_code: str
    customer_name: str
    ship_date: date | None
    amount: int
    pay_mark: str           # 帳單「匯款」欄位原始值：末五碼或 現金/後付/管確認…
    paid_at: date | None = None

    @property
    def last5(self) -> str:
        s = self.pay_mark.strip()
        if s.endswith(".0"):            # 防禦：11414.0 → 11414（SQLite 浮點字串）
            s = s[:-2]
        return s.zfill(5) if s.isdigit() else ""

    @property
    def is_bank_transfer(self) -> bool:
        low = self.pay_mark.strip().lower()
        if any(k in low for k in NON_BANK_MARKS):
            return False
        s = self.pay_mark.strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s.isdigit()


@dataclass
class Result:
    period: tuple[date, date] | None = None
    own_account: str = ""

    bills_total: int = 0
    bills_count: int = 0
    non_bank: list = field(default_factory=list)      # 現金/後付/管確認
    expected_total: int = 0                            # 應匯款金額

    matched: list = field(default_factory=list)        # 完全相符的末五碼群組
    short_paid: list = field(default_factory=list)     # 短收
    over_paid: list = field(default_factory=list)      # 溢收
    wrong_last5: list = field(default_factory=list)    # 末五碼登錄錯誤（款項已收）
    missing: list = field(default_factory=list)        # 查無入帳
    unmatched_bank: list = field(default_factory=list) # 銀行有、帳單無
    parse_errors: list = field(default_factory=list)

    @property
    def received_total(self) -> int:
        """實際銷掉的帳單金額。溢收部分不算沖銷其他帳單，故取帳單金額。"""
        got = sum(g["bank_amount"] for g in self.matched)
        got += sum(g["bank_amount"] for g in self.short_paid)
        got += sum(g["bill_amount"] for g in self.over_paid)
        got += sum(w["amount"] for w in self.wrong_last5 if not w.get("in_group"))
        return got

    @property
    def overpaid_total(self) -> int:
        return sum(g["diff"] for g in self.over_paid)

    @property
    def shortfall(self) -> int:
        return self.expected_total - self.received_total

    def summary(self) -> dict:
        return {
            "帳單筆數": self.bills_count,
            "帳單總額": self.bills_total,
            "非匯款筆數": len(self.non_bank),
            "非匯款金額": sum(b.amount for b in self.non_bank),
            "應匯款金額": self.expected_total,
            "已收金額": self.received_total,
            "差額": self.shortfall,
            "相符群組": len(self.matched),
            "短收": len(self.short_paid),
            "溢收": len(self.over_paid),
            "末五碼錯誤": len(self.wrong_last5),
            "溢收金額": self.overpaid_total,
            "查無入帳": len(self.missing),
            "銀行有帳單無": len(self.unmatched_bank),
        }


def _edit_distance(a: str, b: str) -> int:
    if a == b:
        return 0
    if abs(len(a) - len(b)) > 1:
        return 2
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _own_account_fragments(own_account: str) -> set[str]:
    """自家帳號的所有 5 碼連續片段，用來偵測複製貼上錯誤。"""
    d = "".join(ch for ch in own_account if ch.isdigit())
    return {d[i:i + 5] for i in range(len(d) - 4)} if len(d) >= 5 else set()


def reconcile(bills, bank_records, own_account: str = "",
              period=None, parse_errors=None) -> Result:
    res = Result(period=period, own_account=own_account,
                 parse_errors=list(parse_errors or []))

    res.bills_count = len(bills)
    res.bills_total = sum(b.amount for b in bills)
    res.non_bank = [b for b in bills if not b.is_bank_transfer]
    transfers = [b for b in bills if b.is_bank_transfer]
    res.expected_total = sum(b.amount for b in transfers)

    customer_bank = [r for r in bank_records if r.is_customer_payment]
    other_bank = [r for r in bank_records if not r.is_customer_payment]

    bill_g, bank_g = defaultdict(list), defaultdict(list)
    for b in transfers:
        bill_g[b.last5].append(b)
    for r in customer_bank:
        if r.last5:
            bank_g[r.last5].append(r)

    orphan_bills, orphan_bank = [], []
    for key in set(bill_g) | set(bank_g):
        bs, ks = bill_g.get(key, []), bank_g.get(key, [])
        sb, sk = sum(x.amount for x in bs), sum(x.amount for x in ks)
        if bs and ks:
            group = {
                "last5": key, "bill_amount": sb, "bank_amount": sk,
                "diff": sk - sb,
                "customers": sorted({f"{x.customer_code} {x.customer_name}" for x in bs}),
                "bills": bs, "banks": ks,
            }
            if sk == sb:
                res.matched.append(group)
            elif sk < sb:
                res.short_paid.append(group)
            else:
                res.over_paid.append(group)
        elif bs:
            orphan_bills.extend(bs)
        else:
            orphan_bank.extend(ks)

    # 第二輪：孤兒依「金額 + 日期容差」配對
    used = set()
    own_frags = _own_account_fragments(own_account)
    for b in sorted(orphan_bills, key=lambda x: x.paid_at or x.ship_date or date.min):
        ref = b.paid_at or b.ship_date
        cands = [
            (abs(((r.date - ref).days if ref else 0)), i, r)
            for i, r in enumerate(orphan_bank)
            if i not in used and r.amount == b.amount
            and (ref is None or abs((r.date - ref).days) <= DATE_TOLERANCE_DAYS)
        ]
        if cands:
            _, idx, r = min(cands)
            used.add(idx)
            res.wrong_last5.append({
                "bill_id": b.bill_id, "customer": f"{b.customer_code} {b.customer_name}",
                "amount": b.amount, "bill_last5": b.last5, "bank_last5": r.last5,
                "bank_date": r.date, "payer_name": r.payer_name,
                "reason": "末五碼疑為自家帳號片段" if b.last5 in own_frags
                          else ("他人代匯" if r.payer_name and r.payer_name[0] not in b.customer_name
                                else "末五碼登錄錯誤"),
            })
        else:
            res.missing.append({
                "bill_id": b.bill_id, "customer": f"{b.customer_code} {b.customer_name}",
                "amount": b.amount, "last5": b.last5,
                "ship_date": b.ship_date, "paid_at": b.paid_at,
            })

    # 第三輪：查無入帳的帳單 vs 溢收群組（末五碼只差一碼 → 打錯一碼）
    still_missing = []
    for m in res.missing:
        hit = None
        for g in res.over_paid:
            if g["diff"] >= m["amount"] and _edit_distance(m["last5"], g["last5"]) <= 1:
                cand = [r for r in g["banks"] if r.amount == m["amount"]]
                if cand:
                    hit = (g, cand[0])
                    break
        if not hit:
            still_missing.append(m)
            continue
        g, r = hit
        g["diff"] -= m["amount"]
        g["bill_amount"] += m["amount"]
        res.wrong_last5.append({
            "bill_id": m["bill_id"], "customer": m["customer"], "amount": m["amount"],
            "bill_last5": m["last5"], "bank_last5": g["last5"], "bank_date": r.date,
            "payer_name": r.payer_name, "reason": "末五碼打錯一碼",
            "in_group": True,   # 金額已計入該末五碼群組，避免重複計算
        })
        if g["diff"] == 0:
            res.over_paid.remove(g)
            res.matched.append(g)
    res.missing = still_missing
    res.wrong_last5.sort(key=lambda w: w["bank_date"])

    res.unmatched_bank = [
        {"date": r.date, "amount": r.amount, "last5": r.last5,
         "payer_name": r.payer_name, "memo": r.memo, "tx_code": r.tx_code,
         "category": "客戶款（無對應帳單）"}
        for i, r in enumerate(orphan_bank) if i not in used
    ] + [
        {"date": r.date, "amount": r.amount, "last5": r.last5,
         "payer_name": r.payer_name, "memo": r.memo, "tx_code": r.tx_code,
         "category": "非運費收入"}
        for r in other_bank
    ]
    res.unmatched_bank.sort(key=lambda x: (x["date"], -x["amount"]))
    return res

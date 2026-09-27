# -*- coding: utf-8 -*-
"""
銷帳檔（銀行媒體交換固定長度檔）解析器

實際欄位配置（依 2026/06 銷帳檔 B0103 驗證）：

    123456789012000000000000010  000000002+0000000001843+000000005623000(013)0000699***411414 王＊明        20260630
    |------------ 27 --------|  |--- 9 --|+|--- 13 ---|+|---- 15 ----||3||--- 帳號 ---| |-匯款人-|      |-日期-|
     主帳號+交易代碼(末3碼)       序號        金額(元)      餘額(分)     行代號   匯款人帳號                  YYYYMMDD

註：
* 金額為「元」，餘額為「分」（除以 100）。實測 06-30 餘額 56,230.00 − 54,387.00 = 1,843 ✓
* 匯款人帳號可能被遮罩成 0000699***411414，取數字部分末 5 碼即為系統上的「末五碼」
* 交易代碼（27 碼末 3 位）：010/040 = 一般匯入、007/060 = 其他、982 = 利息等非客戶款
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from datetime import date

# 27碼主帳號 + 空白 + 9碼序號 + 符號 + 13碼金額 + 符號 + 15碼餘額 + 其餘 + 8碼日期
LINE_RE = re.compile(
    r"^(?P<acct>\d{27})\s+"
    r"(?P<seq>\d{9})(?P<s1>[+-])"
    r"(?P<amount>\d{13})(?P<s2>[+-])"
    r"(?P<balance>\d{15})"
    r"(?P<rest>.*?)"
    r"(?P<date>\d{8})\s*$"
)

# (行代號)帳號  — 帳號可能含 * 遮罩
PAYER_RE = re.compile(r"^\((?P<bank>\d{3})\)(?P<acct>[\d*]+)")

# 非客戶收款的交易代碼（利息、行內調撥等），對帳時排除
NON_CUSTOMER_TX_CODES = {"982"}

# 摘要含這些關鍵字者視為非集運運費收入
NON_FREIGHT_KEYWORDS = ("NHI", "健保", "信用卡", "利息", "退稅", "國稅")


@dataclass
class BankRecord:
    """一筆銀行入帳"""
    date: date
    amount: int                 # 元
    balance: float              # 元（原檔為分）
    tx_code: str                # 交易代碼（主帳號末 3 碼）
    payer_bank: str = ""        # 匯款行代號
    payer_acct: str = ""        # 匯款人帳號（可能含遮罩）
    last5: str = ""             # 帳號末五碼
    payer_name: str = ""        # 匯款人姓名（可能遮罩，如 王＊明）
    memo: str = ""              # 附言／摘要
    raw: str = field(default="", repr=False)

    @property
    def is_customer_payment(self) -> bool:
        if self.tx_code in NON_CUSTOMER_TX_CODES:
            return False
        blob = self.memo + self.payer_name
        return not any(k in blob for k in NON_FREIGHT_KEYWORDS)

    def to_dict(self):
        d = asdict(self)
        d["date"] = self.date.isoformat()
        d.pop("raw", None)
        return d


class ParseError(ValueError):
    def __init__(self, lineno: int, line: str):
        super().__init__(f"第 {lineno} 行格式不符：{line[:60]!r}")
        self.lineno, self.line = lineno, line


def _extract_last5(acct: str) -> str:
    digits = re.sub(r"\D", "", acct)
    return digits[-5:] if len(digits) >= 5 else digits


def parse_line(line: str) -> BankRecord | None:
    if not line.strip():
        return None
    m = LINE_RE.match(line)
    if not m:
        raise ValueError("格式不符")
    g = m.groupdict()
    rest = g["rest"].rstrip()

    payer_bank = payer_acct = ""
    remainder = rest
    pm = PAYER_RE.match(rest)
    if pm:
        payer_bank = pm.group("bank")
        payer_acct = pm.group("acct")
        remainder = rest[pm.end():]

    # 匯款人姓名 vs 附言：姓名多為 2-4 個中文字（可能含全形＊遮罩）且前有空白
    payer_name, memo = "", remainder.strip()
    nm = re.match(r"^\s+([\u4e00-\u9fffA-Za-z＊*]{2,6})(?:\s|$)", remainder)
    if nm:
        payer_name = nm.group(1)
        memo = remainder[nm.end():].strip()

    ds = g["date"]
    return BankRecord(
        date=date(int(ds[:4]), int(ds[4:6]), int(ds[6:8])),
        amount=int(g["amount"]) * (-1 if g["s1"] == "-" else 1),
        balance=int(g["balance"]) / 100,
        tx_code=g["acct"][-3:],
        payer_bank=payer_bank,
        payer_acct=payer_acct,
        last5=_extract_last5(payer_acct),
        payer_name=payer_name,
        memo=re.sub(r"[\u3000\s]+", " ", memo).strip(),
        raw=line.rstrip("\n"),
    )


def parse_file(content: str | bytes, strict: bool = False):
    """回傳 (records, errors)。strict=True 時遇到壞行直接拋 ParseError。"""
    if isinstance(content, bytes):
        for enc in ("utf-8", "utf-8-sig", "big5", "cp950"):
            try:
                content = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            content = content.decode("utf-8", errors="replace")

    records, errors = [], []
    for i, line in enumerate(content.splitlines(), 1):
        if not line.strip():
            continue
        try:
            rec = parse_line(line)
        except ValueError:
            if strict:
                raise ParseError(i, line)
            errors.append({"lineno": i, "line": line[:120]})
            continue
        if rec:
            records.append(rec)
    return records, errors

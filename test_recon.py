# -*- coding: utf-8 -*-
"""合成資料測試（不含真實客戶資料）"""
from datetime import date
from recon.parser import parse_file
from recon.matcher import BillRecord, reconcile

SAMPLE = (
    "123456789012000000000000010  000000002+0000000001843+000000005623000(013)0000699***411414 王＊明                      20260630\n"
    "123456789012000000000000040  000000002+0000000000287+000000005438700(826)0081201***087853日本集運                     20260629\n"
    "123456789012000000000000040  000000002+0000000000200+000000005410000(822)0000613***206780                             20260626\n"
    "123456789012000000000000040  000000002+0000000000739+000000005390000(824)0000111***331876                             20260625\n"
    "123456789012000000000000982  000000002+0000000000015+000000005389850                                                  20260615\n"
)

def bill(bid, code, name, amt, mark, d):
    return BillRecord(bid, code, name, d, amt, mark, d)

bills = [
    bill("1", "G0001", "王一", 1843, "11414", date(2026, 6, 30)),   # 相符
    bill("2", "G0002", "李二", 287, "87853", date(2026, 6, 29)),    # 相符
    bill("3", "G0003", "張三", 211, "06780", date(2026, 6, 26)),    # 短收 11
    bill("4", "G0004", "陳四", 739, "10472", date(2026, 6, 25)),    # 末五碼錯 → 31876
    bill("5", "G0005", "林五", 500, "99999", date(2026, 6, 20)),    # 查無入帳
    bill("6", "G0006", "吳六", 300, "現金", date(2026, 6, 18)),      # 非匯款
]

recs, errs = parse_file(SAMPLE)
assert not errs, errs
assert len(recs) == 5
assert recs[0].amount == 1843 and recs[0].last5 == "11414"
assert recs[0].balance == 56230.0
assert recs[4].tx_code == "982" and not recs[4].is_customer_payment  # 利息排除

res = reconcile(bills, recs, own_account="123456789012",
                period=(date(2026, 6, 1), date(2026, 6, 30)), parse_errors=errs)

assert len(res.non_bank) == 1
assert res.expected_total == 1843 + 287 + 211 + 739 + 500
assert len(res.matched) == 2
assert len(res.short_paid) == 1 and res.short_paid[0]["diff"] == -11
assert len(res.wrong_last5) == 1 and res.wrong_last5[0]["bank_last5"] == "31876"
assert len(res.missing) == 1 and res.missing[0]["amount"] == 500
assert res.shortfall == 511
print("✅ 全部通過")

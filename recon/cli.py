# -*- coding: utf-8 -*-
"""
命令列對帳：

    python -m recon.cli 銷帳檔B0103.txt --start 2026-06-01 --end 2026-06-30
    python -m recon.cli 銷帳檔.txt --start 2026-07-01 --end 2026-07-31 --excel 7月對帳.xlsx
    python -m recon.cli --schema          # 印出資料表結構，用來確認 BILL_SQL
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime

from . import db
from .matcher import reconcile
from .parser import parse_file
from .report import to_excel, to_text


def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def main(argv=None):
    p = argparse.ArgumentParser(description="helpshipping 銷帳檔對帳")
    p.add_argument("file", nargs="?", help="銷帳檔路徑")
    p.add_argument("--start", type=_d, help="付款日起 YYYY-MM-DD")
    p.add_argument("--end", type=_d, help="付款日迄 YYYY-MM-DD")
    p.add_argument("--db", default=None, help="SQLite 路徑")
    p.add_argument("--excel", help="輸出 xlsx 路徑")
    p.add_argument("--json", help="輸出 json 路徑")
    p.add_argument("--schema", action="store_true", help="印出資料表結構後結束")
    args = p.parse_args(argv)

    conn = db.connect(args.db)
    if args.schema:
        print(json.dumps(db.inspect_schema(conn), ensure_ascii=False, indent=2))
        return 0

    if not (args.file and args.start and args.end):
        p.error("需要 file、--start、--end")

    with open(args.file, "rb") as f:
        records, errors = parse_file(f.read())

    bills = db.fetch_bills(conn, args.start, args.end)
    res = reconcile(bills, records, own_account=db.OWN_ACCOUNT,
                    period=(args.start, args.end), parse_errors=errors)

    print(to_text(res))

    if args.excel:
        with open(args.excel, "wb") as f:
            f.write(to_excel(res))
        print(f"\n已輸出 {args.excel}")
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(res.summary(), f, ensure_ascii=False, indent=2)

    # 有短收或查無入帳 → exit code 1，方便排程時觸發告警
    return 1 if (res.missing or res.short_paid) else 0


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""
Flask blueprint：上傳銷帳檔 → 對帳 → 網頁報表 / Excel 下載

app.py 加兩行：
    from recon.routes import bp as recon_bp
    app.register_blueprint(recon_bp)
（DB_PATH 沿用 app.py 的環境變數，不必另外設定）
"""
from __future__ import annotations

import io
from datetime import date, datetime

from flask import (Blueprint, current_app, render_template, request,
                   send_file, session)

from . import db
from .matcher import reconcile
from .parser import parse_file
from .report import to_excel, to_text

bp = Blueprint("recon", __name__, url_prefix="/recon",
               template_folder="../templates")

# 最近一次對帳結果暫存（單機夠用；多 worker 請改存 SQLite 或檔案）
_LAST = {}


def _month_range(today: date | None = None):
    today = today or date.today()
    first = today.replace(day=1)
    prev_end = first - __import__("datetime").timedelta(days=1)
    return prev_end.replace(day=1), prev_end


def _login_required():
    """對帳是財務功能，限老闆（super）：與 app.py 的 is_boss() 一致。"""
    return session.get("user_type") == "admin" and session.get("role") == "super"


def _current_agent_id() -> int:
    """代理只看自己的帳單；主管理員（0）看全部。"""
    try:
        return int(session.get("agent_id", 0) or 0)
    except (TypeError, ValueError):
        return 0


@bp.route("/", methods=["GET", "POST"])
def index():
    if not _login_required():
        return render_template("recon.html", need_login=True), 401

    s_def, e_def = _month_range()
    ctx = {"start": s_def.isoformat(), "end": e_def.isoformat(),
           "res": None, "text": "", "error": None}

    if request.method == "POST":
        f = request.files.get("file")
        try:
            start = datetime.strptime(request.form["start"], "%Y-%m-%d").date()
            end = datetime.strptime(request.form["end"], "%Y-%m-%d").date()
            ctx.update(start=request.form["start"], end=request.form["end"])
            if not f or not f.filename:
                raise ValueError("請選擇銷帳檔")

            records, errors = parse_file(f.read())
            if not records:
                raise ValueError("銷帳檔沒有解析到任何資料，請確認檔案格式")

            conn = db.connect(current_app.config.get("DB_PATH"))
            bills = db.fetch_bills(conn, start, end,
                                   agent_id=_current_agent_id())
            # 一律附診斷（撈到帳單數 / 銀行筆數 / 帳單末五碼樣態），方便排查對不上的原因
            diag = db.diagnose(conn, start, end, agent_id=_current_agent_id())
            diag["撈到帳單"] = len(bills)
            diag["銀行入帳筆數"] = len(records)
            ctx["diag"] = diag
            res = reconcile(bills, records, own_account=db.OWN_ACCOUNT,
                            period=(start, end), parse_errors=errors)
            _LAST[session.get("user_id") or session.get("username") or "?"] = res
            ctx["res"], ctx["text"] = res, to_text(res)
        except Exception as exc:                     # noqa: BLE001
            current_app.logger.exception("對帳失敗")
            ctx["error"] = str(exc)

    return render_template("recon.html", **ctx)


@bp.route("/export.xlsx")
def export():
    if not _login_required():
        return "unauthorized", 401
    res = _LAST.get(session.get("user_id") or session.get("username") or "?")
    if not res:
        return "尚未執行對帳", 400
    p = res.period
    name = f"對帳_{p[0]:%Y%m%d}-{p[1]:%Y%m%d}.xlsx" if p else "對帳.xlsx"
    return send_file(io.BytesIO(to_excel(res)), as_attachment=True,
                     download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument."
                              "spreadsheetml.sheet")

"""FWT JAPAN 雲端集運 PWA 整合模組

提供 3 條路由（與現有路由零衝突）：
  GET /sw.js                       → service worker
  GET /manifest.webmanifest        → 客戶端 PWA（啟動頁 /）
  GET /admin-manifest.webmanifest  → 後台 PWA（啟動頁 /admin）

用法（在 app.py 加一行）：
    from pwa import register_pwa
    register_pwa(app)

iOS Safari 16.4+ 與 Android Chrome 全支援。
桌面版（Chrome / Edge）也會出現「安裝」按鈕，可當桌機獨立 App 使用。
"""
from flask import send_from_directory, jsonify, make_response
from brand import BRAND

# 品牌色（跟 templates/index.html UI 一致）
_THEME = "#16213e"   # 海軍藍
_BG    = "#16213e"


def register_pwa(app):

    @app.route("/sw.js")
    def _pwa_sw():
        """Service worker。從根目錄提供 → scope 涵蓋整站。"""
        resp = send_from_directory(app.root_path, "sw.js")
        resp.headers["Content-Type"] = "application/javascript"
        resp.headers["Service-Worker-Allowed"] = "/"
        resp.headers["Cache-Control"] = "no-cache"  # SW 本身不長期快取，方便推新版
        return resp

    @app.route("/manifest.webmanifest")
    def _pwa_manifest_customer():
        manifest = {
            "name":             f"{BRAND['name']} 雲端集運",
            "short_name":       "集運",
            "description":      f"{BRAND['name']} 雲端集運 — 包裹預報、查詢、出貨申請",
            "start_url":        "/?source=pwa",
            "scope":            "/",
            "display":          "standalone",
            "orientation":      "portrait",
            "background_color": _BG,
            "theme_color":      _THEME,
            "icons": [
                {"src": "/static/icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
                {"src": "/static/icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
                {"src": "/static/icons/icon-512-maskable.png", "sizes": "512x512",
                 "type": "image/png", "purpose": "maskable"},
            ],
        }
        resp = jsonify(manifest)
        resp.headers["Content-Type"] = "application/manifest+json"
        return resp

    @app.route("/admin-manifest.webmanifest")
    def _pwa_manifest_admin():
        manifest = {
            "name":             "集運管理後台",
            "short_name":       "集運後台",
            "description":      f"{BRAND['name']} 後台 — 到貨、出貨、帳單、出檔案",
            "start_url":        "/admin?source=pwa",
            "scope":            "/",
            "display":          "standalone",
            "orientation":      "any",
            "background_color": _BG,
            "theme_color":      _THEME,
            "icons": [
                {"src": "/static/icons/icon-admin-192.png", "sizes": "192x192", "type": "image/png"},
                {"src": "/static/icons/icon-admin-512.png", "sizes": "512x512", "type": "image/png"},
                {"src": "/static/icons/icon-admin-512-maskable.png", "sizes": "512x512",
                 "type": "image/png", "purpose": "maskable"},
            ],
        }
        resp = jsonify(manifest)
        resp.headers["Content-Type"] = "application/manifest+json"
        return resp

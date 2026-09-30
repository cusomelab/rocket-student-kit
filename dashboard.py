# -*- coding: utf-8 -*-
"""쿠팡 로켓배송 운영 대시보드 (수강생용).

탭: 신규 상품 · 발주/쉽먼트 · 발주서 만들기 · 설정
실행: 대시보드_실행.bat  (또는 python dashboard.py)  →  http://127.0.0.1:5080
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import json
import os
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from collections import deque
from pathlib import Path
from urllib.parse import quote

from flask import Flask, Response, jsonify, render_template, request, send_file

import app_settings

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(__file__).parent
WEB = BASE / "web"
# 바코드·라벨 웹앱 주소 (강사가 배포한 Streamlit 주소). 설정에서 바꿀 수 있다.
BARCODE_URL_FILE = BASE / "barcode_app_url.txt"
PY = sys.executable
PORT = int(os.environ.get("DASHBOARD_PORT", "5080"))
CHROME_PORT = 9334
VERSION = (BASE / "VERSION").read_text(encoding="utf-8").strip() if (BASE / "VERSION").exists() else ""
CSV_LATEST = BASE / "쿠팡_공급SKU_최신.csv"
SUPPLIER_URL = "https://supplier.coupang.com/plan/ticket/supplySkuList"
PO_URL = "https://supplier.coupang.com/po-web/purchase/order/list"
PENDING_PATH = BASE / "po_pending.json"

MODES = {
    "collect": {
        "label": "신규 상품 수집",
        "hint": "서플라이어 허브 공급 SKU 전체를 CSV로 모은다",
        "argv": [str(BASE / "쿠팡수집.py")],
        "timeout": 5400,
    },
    "po_list": {
        "label": "미확정 발주 조회",
        "hint": "공급사 허브에서 발주 목록을 읽는다",
        "argv": [str(BASE / "po_shipment.py"), "list"],
        "timeout": 1800,
    },
    "po_confirm": {
        "label": "발주확정",
        "hint": "미확정 발주를 공급사 확인 처리한다",
        "argv": [str(BASE / "po_shipment.py"), "confirm"],
        "timeout": 5400,
    },
    "ship_upload": {
        "label": "쉽먼트 업로드",
        "hint": "Shipment_Upload 엑셀을 허브에 올린다",
        "argv": [str(BASE / "po_shipment.py"), "upload"],
        "timeout": 1800,
    },
    "po_ship": {
        "label": "발주확정 → 쉽먼트 업로드",
        "hint": "조회 후 확정하고 엑셀을 업로드한다",
        "argv": [str(BASE / "po_shipment.py"), "pipeline"],
        "timeout": 5400,
    },
    "ship_plan": {
        "label": "쉽먼트 시트 ①~⑤",
        "hint": "리스트업 → 동기화 → 계획 → 정렬 → 예상쉽먼트",
        "argv": [str(BASE / "sheet_shipment.py")],
        "timeout": 3600,
    },
    "drive_list": {
        "label": "드라이브 목록",
        "hint": "쉽먼트 폴더의 xlsx 목록을 읽는다",
        "argv": [str(BASE / "drive_ship.py"), "list"],
        "timeout": 180,
    },
    "order_fill": {
        "label": "발주서 만들기",
        "hint": "허브 다운로드 → INBOX → 집계 → 차감 → 채우기 → 커밋",
        "argv": [str(BASE / "order_build.py")],
        "timeout": 3600,
    },
    "order_commit": {
        "label": "발주 커밋",
        "hint": "발주서 확인 후 PO/SO 확정",
        "argv": [str(BASE / "sheet_order.py"), "commit"],
        "timeout": 900,
    },
}

app = Flask(
    __name__,
    template_folder=str(WEB / "templates"),
    static_folder=str(WEB / "static"),
)


# ── CSV 캐시 ──────────────────────────────────────────────
class Store:
    def __init__(self):
        self.lock = threading.Lock()
        self.mtime = 0.0
        self.rows: list[dict] = []
        self.prev_path: Path | None = None
        self.prev_skus: set[str] = set()
        self.prev_mtime = 0.0

    def _read_csv(self, path: Path) -> list[dict]:
        rows = []
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            for raw in csv.DictReader(fh):
                sku = (raw.get("SKU ID") or "").strip()
                if not sku:
                    continue
                rows.append({
                    "sku": sku,
                    "name": (raw.get("상품명") or "").strip(),
                    "barcode": (raw.get("바코드") or "").strip(),
                    "status": (raw.get("발주가능상태") or "").strip() or "미기재",
                    "bm": (raw.get("담당 BM") or "").strip(),
                    "scm": (raw.get("발주담당자") or "").strip(),
                    "image": (raw.get("이미지url") or "").strip(),
                    "moq": (raw.get("최소구매수량") or "").strip(),
                })
        return rows

    def _previous_path(self) -> Path | None:
        # 수집하면 최신.csv 와 그날 날짜 파일이 같이 생긴다.
        # 최신.csv 가 만들어진 날짜보다 앞선 날짜 파일과 비교한다.
        files = sorted(BASE.glob("쿠팡_공급SKU_20??-??-??.csv"))
        if CSV_LATEST.exists():
            day = dt.date.fromtimestamp(CSV_LATEST.stat().st_mtime).isoformat()
        else:
            day = dt.date.today().isoformat()
        cands = [f for f in files if f.stem.rsplit("_", 1)[-1] < day]
        return cands[-1] if cands else None

    def refresh(self) -> None:
        mtime = CSV_LATEST.stat().st_mtime if CSV_LATEST.exists() else 0.0
        prev = self._previous_path()
        prev_mtime = prev.stat().st_mtime if prev and prev.exists() else 0.0
        with self.lock:
            if mtime != self.mtime and CSV_LATEST.exists():
                self.rows = self._read_csv(CSV_LATEST)
                self.mtime = mtime
            if prev != self.prev_path or prev_mtime != self.prev_mtime:
                self.prev_path = prev
                self.prev_mtime = prev_mtime
                self.prev_skus = {r["sku"] for r in self._read_csv(prev)} if prev else set()

    def new_rows(self) -> tuple[list[dict], Path | None, float, int]:
        self.refresh()
        with self.lock:
            if not self.prev_path:
                return [], None, self.mtime, len(self.rows)
            items = [r for r in self.rows if r["sku"] not in self.prev_skus]
            items.reverse()  # 최근 SKU가 위로
            return items, self.prev_path, self.mtime, len(self.rows)


store = Store()


# ── 작업 실행 ─────────────────────────────────────────────
class Job:
    def __init__(self):
        self.lock = threading.Lock()
        self.running = False
        self.mode = ""
        self.label = ""
        self.started_at = ""
        self.ended_at = ""
        self.returncode: int | None = None
        self.lines: deque[str] = deque(maxlen=4000)
        self.proc: subprocess.Popen | None = None
        self.need_login = False
        self.denied = False
        self._cv = threading.Condition(self.lock)

    def status(self) -> dict:
        with self.lock:
            return {
                "running": self.running,
                "mode": self.mode,
                "label": self.label,
                "started_at": self.started_at,
                "ended_at": self.ended_at,
                "returncode": self.returncode,
                "need_login": self.need_login,
                "denied": self.denied,
            }

    def tail(self, n: int = 200) -> list[str]:
        with self.lock:
            return list(self.lines)[-n:]

    def start(self, mode: str) -> tuple[bool, str]:
        spec = MODES.get(mode)
        if spec is None:
            return False, "알 수 없는 작업입니다."
        with self.lock:
            if self.running:
                return False, "이미 작업이 실행 중입니다."
            self.running = True
            self.mode = mode
            self.label = spec["label"]
            self.started_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.ended_at = ""
            self.returncode = None
            self.need_login = False
            self.denied = False
            self.lines.clear()
            self.lines.append(f"▶ {spec['label']} 시작 ({spec['hint']})")
        threading.Thread(target=self._run, args=(spec,), daemon=True).start()
        return True, spec["label"]

    def stop(self) -> bool:
        with self.lock:
            proc = self.proc
            if not self.running or proc is None:
                return False
            self.lines.append("■ 사용자가 작업을 중단했습니다.")
        try:
            proc.terminate()
        except Exception:
            pass
        return True

    def _note(self, line: str) -> None:
        low = line.lower()
        with self.lock:
            self.lines.append(line.rstrip())
            if "직접 로그인" in line or "캡차" in line or "슬라이더" in line:
                self.need_login = True
            if "access denied" in low:
                self.denied = True
            self._cv.notify_all()

    def _run(self, spec: dict) -> None:
        cmd = [PY, "-u", *spec["argv"]]
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                cwd=str(BASE),
            )
        except Exception as e:
            with self.lock:
                self.lines.append(f"실행 실패: {e}")
                self.running = False
                self.ended_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.returncode = 1
                self._cv.notify_all()
            return
        with self.lock:
            self.proc = proc
        assert proc.stdout is not None
        t0 = time.time()
        timeout = spec["timeout"]
        try:
            while True:
                if time.time() - t0 > timeout:
                    proc.kill()
                    self._note(f"시간 초과({timeout // 60}분)로 중단했습니다.")
                    break
                line = proc.stdout.readline()
                if line == "" and proc.poll() is not None:
                    break
                if line:
                    self._note(line.rstrip("\n"))
            code = proc.wait(timeout=5)
        except Exception as e:
            self._note(f"작업 오류: {e}")
            try:
                proc.kill()
            except Exception:
                pass
            code = proc.poll() if proc.poll() is not None else 1
        with self.lock:
            self.proc = None
            self.returncode = int(code or 0)
            self.ended_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ok = self.returncode == 0
            self.lines.append(
                f"{'✓' if ok else '✗'} {spec['label']} 종료 (코드 {self.returncode})"
            )
            self.running = False
            self._cv.notify_all()
        store.refresh()


job = Job()


def chrome_up() -> bool:
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{CHROME_PORT}/json/version", timeout=0.6)
        return True
    except Exception:
        return False


def _po_payload() -> dict:
    empty = {"fetched_at": "", "total": 0, "pending": 0, "rows": [], "pending_rows": []}
    if not PENDING_PATH.exists():
        return empty
    try:
        return json.loads(PENDING_PATH.read_text(encoding="utf-8"))
    except Exception:
        return empty


def _ship_files() -> dict:
    if app_settings.is_drive():
        url = app_settings.drive_url()
        cache = {}
        p = app_settings.DRIVE_CACHE
        if p.exists():
            try:
                cache = json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                cache = {}
        files = cache.get("files") or []
        if cache.get("url") and cache.get("url") != url:
            files = []
        return {
            "kind": "drive",
            "dir": url,
            "exists": True,
            "fetched_at": cache.get("fetched_at") or "",
            "files": files,
        }
    folder = app_settings.ship_folder()
    out = []
    for p in app_settings.list_ship_excels():
        out.append({
            "name": p.name,
            "mtime": dt.datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
            "size_kb": round(p.stat().st_size / 1024, 1),
        })
    return {
        "kind": "local",
        "dir": str(folder),
        "exists": folder.exists(),
        "fetched_at": "",
        "files": out,
    }


def _settings_view() -> dict:
    data = app_settings.load()
    raw = app_settings.ship_folder_raw()
    data["ship_folder"] = app_settings.drive_url() if app_settings.is_drive(raw) else raw
    data["ship_kind"] = "drive" if app_settings.is_drive(raw) else "local"
    data["ship_folder_resolved"] = str(app_settings.ship_folder())
    data["inbox_folder"] = app_settings.inbox_drive_url()
    data["version"] = VERSION
    data["barcode_app_url"] = _barcode_url()
    return data


def _barcode_url() -> str:
    url = (app_settings.load().get("barcode_app_url") or "").strip()
    if not url and BARCODE_URL_FILE.exists():
        url = BARCODE_URL_FILE.read_text(encoding="utf-8").strip()
    return url



# ── 페이지 / 정적 ─────────────────────────────────────────
@app.get("/")
def index():
    return render_template("index.html", version=VERSION)


@app.get("/favicon.ico")
def favicon():
    ico = BASE / "rocket.ico"
    if ico.exists():
        return send_file(ico, mimetype="image/x-icon")
    return ("", 404)


# ── API ───────────────────────────────────────────────────
@app.get("/api/health")
def api_health():
    return jsonify({"ok": True, "port": PORT, "version": VERSION})


@app.get("/api/status")
def api_status():
    return jsonify({
        "chrome": chrome_up(),
        "job": job.status(),
        "version": VERSION,
        "settings": _settings_view(),
    })


@app.get("/api/new")
def api_new():
    items, prev_path, mtime, total = store.new_rows()
    return jsonify({
        "total": len(items),
        "all_total": total,
        "csv_mtime": dt.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M") if mtime else "",
        "previous": prev_path.name if prev_path else "",
        "rows": items[:500],
    })


@app.get("/api/new/export")
def api_new_export():
    items, *_ = store.new_rows()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["SKU ID", "상품명", "바코드", "발주가능상태", "담당 BM", "발주담당자", "최소구매수량", "이미지url"])
    for r in items:
        w.writerow([r["sku"], r["name"], r["barcode"], r["status"], r["bm"], r["scm"], r["moq"], r["image"]])
    raw = ("﻿" + buf.getvalue()).encode("utf-8")
    fname = f"신규상품_{dt.date.today().isoformat()}.csv"
    return Response(
        raw,
        mimetype="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": f"attachment; filename=new_products.csv; filename*=UTF-8''{quote(fname)}"
        },
    )


@app.get("/api/po")
def api_po():
    return jsonify(_po_payload())


@app.get("/api/shipment/files")
def api_ship_files():
    return jsonify(_ship_files())


@app.get("/api/settings")
def api_settings_get():
    return jsonify(_settings_view())


@app.post("/api/settings")
def api_settings_save():
    body = request.get_json(silent=True) or {}
    patch = {}
    if "order_sheet_url" in body:
        url = (body.get("order_sheet_url") or "").strip()
        if url and not app_settings.SHEET_ID_RE.search(url):
            return jsonify({"ok": False, "error": "구글시트 주소가 올바르지 않습니다. (docs.google.com/spreadsheets/d/... 형식)"}), 400
        patch["order_sheet_url"] = url
    if "ship_folder" in body:
        folder = (body.get("ship_folder") or "").strip()
        if folder and app_settings.parse_drive_id(folder):
            patch["ship_folder"] = app_settings.drive_url(folder)
        elif folder:
            p = Path(folder)
            if not p.is_dir():
                return jsonify({"ok": False, "error": f"폴더가 없습니다: {folder}"}), 400
            patch["ship_folder"] = str(p)
        else:
            patch["ship_folder"] = ""
    if "inbox_folder" in body:
        inbox = (body.get("inbox_folder") or "").strip()
        if inbox and not app_settings.parse_drive_id(inbox):
            return jsonify({"ok": False, "error": "INBOX 폴더는 구글 드라이브 폴더 주소를 넣어 주세요."}), 400
        patch["inbox_folder"] = app_settings.drive_url(inbox) if inbox else ""
    if "inbound_date" in body:
        day = (body.get("inbound_date") or "").strip()
        if day:
            try:
                dt.datetime.strptime(day, "%Y-%m-%d")
            except ValueError:
                return jsonify({"ok": False, "error": "입고예정일은 YYYY-MM-DD 형식입니다."}), 400
        patch["inbound_date"] = day
    if "barcode_app_url" in body:
        url = (body.get("barcode_app_url") or "").strip()
        if url and not url.startswith("http"):
            return jsonify({"ok": False, "error": "바코드 앱 주소는 https:// 로 시작해야 합니다."}), 400
        patch["barcode_app_url"] = url
    if not patch:
        return jsonify({"ok": False, "error": "저장할 값이 없습니다."}), 400
    app_settings.save(patch)
    return jsonify({"ok": True, **_settings_view(), "ship": _ship_files()})


@app.post("/api/barcode/open")
def api_barcode_open():
    url = _barcode_url()
    if not url:
        return jsonify({"ok": False, "error": "설정 탭에 바코드·라벨 앱 주소를 넣어 주세요. (강사에게 받은 주소)"}), 400
    webbrowser.open(url)
    return jsonify({"ok": True, "url": url})


@app.get("/api/job")
def api_job():
    return jsonify(job.status())


@app.get("/api/job/log")
def api_job_log():
    return jsonify({"lines": job.tail(400), "status": job.status()})


@app.post("/api/job")
def api_job_start():
    data = request.get_json(silent=True) or {}
    mode = (data.get("mode") or "").strip()
    ok, msg = job.start(mode)
    return jsonify({"ok": ok, "message": msg, **job.status()}), (200 if ok else 409)


@app.post("/api/job/stop")
def api_job_stop():
    stopped = job.stop()
    return jsonify({"ok": stopped, **job.status()})


@app.get("/api/job/stream")
def api_job_stream():
    def gen():
        idx = 0
        yield "event: hello\ndata: {}\n\n"
        while True:
            with job._cv:
                lines = list(job.lines)
                running = job.running
                status = {
                    "running": running,
                    "mode": job.mode,
                    "label": job.label,
                    "need_login": job.need_login,
                    "denied": job.denied,
                    "returncode": job.returncode,
                    "ended_at": job.ended_at,
                }
                if idx >= len(lines) and running:
                    job._cv.wait(timeout=0.8)
                    continue
            fresh = lines[idx:]
            idx = len(lines)
            for line in fresh:
                payload = json.dumps({"line": line}, ensure_ascii=False)
                yield f"event: log\ndata: {payload}\n\n"
            yield f"event: status\ndata: {json.dumps(status, ensure_ascii=False)}\n\n"
            if not running and idx >= len(lines):
                yield "event: done\ndata: {}\n\n"
                break
            time.sleep(0.2)

    return Response(
        gen(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no", "Connection": "keep-alive"},
    )


def _open_hub_chrome() -> None:
    """작업용 Chrome을 띄우고 발주 목록을 앞으로 가져온다."""
    import 쿠팡수집 as hub
    from playwright.sync_api import sync_playwright

    hub.ensure_chrome()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CHROME_PORT}")
        ctx = browser.contexts[0]
        page = next(
            (pg for pg in ctx.pages
             if not pg.is_closed() and "supplier.coupang.com" in (pg.url or "")),
            None,
        )
        if page is None:
            page = ctx.new_page()
            try:
                page.goto(PO_URL, wait_until="domcontentloaded")
            except Exception:
                pass
        try:
            page.bring_to_front()
        except Exception:
            pass


@app.post("/api/open")
def api_open():
    data = request.get_json(silent=True) or {}
    target = data.get("target")
    if target == "hub_chrome":
        try:
            _open_hub_chrome()
            return jsonify({"ok": True})
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500
    mapping = {
        "sheet": app_settings.order_sheet_url(),
        "supplier": SUPPLIER_URL,
        "po": PO_URL,
        "folder": str(BASE),
        "ship_folder": app_settings.drive_url() if app_settings.is_drive() else str(app_settings.ship_folder()),
        "inbox_folder": app_settings.inbox_drive_url(),
    }
    if target not in mapping:
        return jsonify({"ok": False, "error": "unknown target"}), 400
    path = mapping[target]
    if not path:
        return jsonify({"ok": False, "error": "설정 탭에서 주소를 먼저 저장해 주세요."}), 400
    try:
        if path.startswith("http"):
            webbrowser.open(path)
        else:
            Path(path).mkdir(parents=True, exist_ok=True)
            os.startfile(path)  # type: ignore[attr-defined]
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


def main():
    print(f"쿠팡 로켓 대시보드 (수강생용 {VERSION})  http://127.0.0.1:{PORT}", flush=True)
    print("이 창을 닫으면 대시보드가 꺼집니다.", flush=True)
    if os.environ.get("DASHBOARD_NO_BROWSER") != "1":
        threading.Timer(0.8, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    app.run(host="127.0.0.1", port=PORT, threaded=True, use_reloader=False)


if __name__ == "__main__":
    main()

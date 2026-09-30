# -*- coding: utf-8 -*-
"""쿠팡 로켓배송(Supplier Hub) 공급 SKU 전체 수집 → CSV
- 작업용 Chrome(포트 9334)이 없으면 자동 실행
- 로그인 세션은 chrome_profile 폴더에 유지됨 (만료 시 로그인 안내 후 자동 대기)
- 결과: 쿠팡_공급SKU_최신.csv (구글시트 A~H열 순서와 동일)
"""
import csv
import datetime
import json
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
BASE = Path(__file__).parent
LIST_URL = "https://supplier.coupang.com/plan/ticket/supplySkuList"
PORT = 9334  # 작업용 Chrome 디버그 포트
PROFILE = BASE / "chrome_profile"

DETAIL_JS = """async (ids) => {
  const out = {};
  await Promise.all(ids.map(async (id) => {
    try {
      const r = await fetch('/plan/v1/ticket/sku/get/detail?skuId=' + id + '&locale=ko', {
        headers: { 'x-requested-with': 'XMLHttpRequest', 'locale': 'ko' },
        credentials: 'include',
      });
      const ct = r.headers.get('content-type') || '';
      if (!ct.includes('json')) { out[id] = { __err: r.status }; return; }
      const j = await r.json();
      out[id] = (j && j.successful && j.body) ? j.body : { __err: 'bad' };
    } catch (e) { out[id] = { __err: String(e) }; }
  }));
  return out;
}"""

FETCH_JS = """async (body) => {
  const r = await fetch('/plan/v1/ticket/sku/listTicketSku?locale=ko', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-requested-with': 'XMLHttpRequest',
      'locale': 'ko',
    },
    body: JSON.stringify(body),
    credentials: 'include',
  });
  const ct = r.headers.get('content-type') || '';
  if (!ct.includes('json')) {
    let denied = (r.status === 403 || r.status === 429);
    try {
      const txt = await r.text();
      if (/access denied/i.test(txt)) denied = true;
    } catch (e) {}
    return { __err: denied ? 'denied' : 'notjson', status: r.status };
  }
  return await r.json();
}"""


# 쿠키 문자열이 너무 길면 400이 난다. _abck(Akamai 신뢰 쿠키)는 지우면
# Access Denied가 나므로 bm_* 만 줄인다.
TRIM_COOKIES_JS = """() => {
  if (document.cookie.length < 3500) return document.cookie.length;
  const kill = ['bm_sz', 'bm_sv', 'bm_mi', 'bm_lso', 'bm_so', 'ak_bmsc', 'bm_s'];
  for (const name of kill) {
    for (const dom of ['.coupang.com', '.supplier.coupang.com', location.hostname, '']) {
      document.cookie = name + '=; Max-Age=0; path=/' + (dom ? '; domain=' + dom : '');
    }
  }
  return document.cookie.length;
}"""

DENIED_JS = """() => {
  const t = ((document.title || '') + ' ' + (document.body && document.body.innerText || '')).slice(0, 4000);
  return /access denied|접근이 거부|errors\\.edgesuite|akamai/i.test(t);
}"""


def trim_fat_cookies(page_obj):
    try:
        page_obj.evaluate(TRIM_COOKIES_JS)
    except Exception:
        pass


def is_access_denied(page_obj):
    try:
        if page_obj is None or page_obj.is_closed():
            return False
        return bool(page_obj.evaluate(DENIED_JS))
    except Exception:
        return False


def revive_page(ctx, page_obj):
    """전용 탭이 닫혔으면(사용자가 닫는 등) 새 탭을 열어 돌려준다."""
    try:
        if page_obj is not None and not page_obj.is_closed():
            return page_obj
    except Exception:
        pass
    pg = ctx.new_page()
    try:
        pg.goto(LIST_URL, wait_until="domcontentloaded")
        time.sleep(2)
    except Exception:
        pass
    return pg


def port_open():
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{PORT}/json/version", timeout=1)
        return True
    except Exception:
        return False


def ensure_chrome():
    if port_open():
        return
    chrome = next(
        (c for c in [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            str(Path.home() / "AppData/Local/Google/Chrome/Application/chrome.exe"),
        ] if Path(c).exists()),
        None,
    )
    if not chrome:
        print("Chrome 실행 파일을 찾지 못했습니다.")
        sys.exit(1)
    subprocess.Popen([
        chrome,
        f"--remote-debugging-port={PORT}",
        f"--user-data-dir={PROFILE}",
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
        LIST_URL,
    ])
    for _ in range(60):
        if port_open():
            return
        time.sleep(0.5)
    print("Chrome 디버그 포트 연결 실패. 열려 있는 이 프로필의 Chrome을 모두 닫고 다시 실행해 주세요.")
    sys.exit(1)


def make_body(page_no, size):
    return {
        "skuId": "", "skuName": "", "barcode": "", "orderingStatus": "",
        "unit1": "", "unit2": "", "issueStatus": "", "issueType": "",
        "size": size, "page": page_no,
    }


def api(page_obj, page_no, size):
    return page_obj.evaluate(FETCH_JS, make_body(page_no, size))


def get_content(res):
    if not isinstance(res, dict) or res.get("__err") or not res.get("successful"):
        return None, None
    body = res.get("body") or {}
    return body.get("content") or [], body.get("total")


def img_https(u):
    if isinstance(u, str) and u.startswith("http://"):
        return "https://" + u[7:]
    return u or ""


def main():
    ensure_chrome()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{PORT}")
        ctx = browser.contexts[0]
        # 전용 탭에서 plan 페이지를 연다 (사용자 탭은 건드리지 않음).
        # plan 세션이 없으면 xauth 로그인 페이지로 리다이렉트됨.
        page = ctx.new_page()
        try:
            page.goto(LIST_URL, wait_until="domcontentloaded")
        except Exception:
            pass
        time.sleep(3)

        # 로그인 확인 (plan API가 JSON을 주면 로그인된 것).
        # 로그인 버튼은 절대 자동으로 누르지 않는다 — Akamai가 Access Denied로 막는다.
        first = None
        warned_manual = False
        warned_denied = False
        for i in range(1440):  # 최대 2시간 대기
            if page.is_closed():
                page = ctx.new_page()
                try:
                    page.goto(LIST_URL, wait_until="domcontentloaded")
                except Exception:
                    pass
                time.sleep(3)

            denied = is_access_denied(page)
            url = ""
            try:
                url = page.url or ""
            except Exception:
                url = ""

            if denied:
                if not warned_denied:
                    warned_denied = True
                    print(">>> Access Denied가 떴습니다. 자동 클릭/새로고침은 하지 않습니다.", flush=True)
                    print("    Chrome 창에서 주소창을 누르고 F5로 새로고침한 뒤, 직접 로그인하세요.", flush=True)
                time.sleep(5)
                continue

            if "supplier.coupang.com" in url:
                try:
                    time.sleep(1.5)  # Akamai 센서 스크립트가 돌 시간
                    res = api(page, 1, 10)
                    content, total = get_content(res)
                    if content is not None:
                        first = (res, total)
                        break
                    err = res.get("__err") if isinstance(res, dict) else None
                    if err == "denied":
                        if not warned_denied:
                            warned_denied = True
                            print(">>> API가 Access Denied를 반환했습니다. 잠시 기다립니다.", flush=True)
                            print("    Chrome에서 페이지를 직접 새로고침한 뒤 로그인 상태를 확인해 주세요.", flush=True)
                        time.sleep(8)
                        continue
                except Exception:
                    pass
            elif "xauth" in url:
                if not warned_manual:
                    warned_manual = True
                    print(">>> 쿠팡 로그인 페이지입니다. Chrome 창에서 직접 로그인해 주세요.", flush=True)
                    print("    (자동으로 버튼을 누르지 않습니다. 막히면 더 심해집니다.)", flush=True)
            elif i == 0 or (i % 12 == 0 and not denied):
                try:
                    page.goto(LIST_URL, wait_until="domcontentloaded")
                except Exception:
                    pass

            if i == 0 and not warned_manual and not denied:
                print(">>> 로그인 세션을 확인합니다. 필요하면 Chrome에서 직접 로그인해 주세요.", flush=True)
            time.sleep(5)
        if first is None:
            print("로그인을 확인하지 못했습니다. 다시 실행해 주세요.")
            sys.exit(1)

        _, total = first
        print(f"로그인 확인 완료. 전체 상품 수: {total}개")

        # 페이지당 개수 확대 시도
        size = 10
        for trial in (500, 100, 50):
            try:
                res = api(page, 1, trial)
                content, _ = get_content(res)
                if content and len(content) > size:
                    size = len(content)
                    if size >= trial:
                        break
            except Exception:
                pass
        print(f"페이지당 {size}개씩 수집합니다.")

        total_pages = (total + size - 1) // size if total else 4000
        rows = []
        seen = set()
        t0 = time.time()
        pg_no = 1
        err_streak = 0
        while pg_no <= total_pages:
            if err_streak > 30:
                print("연속 오류 30회 초과로 중단합니다. (기존 CSV는 그대로 두고 종료)", flush=True)
                sys.exit(1)
            try:
                res = api(page, pg_no, size)
                content, _ = get_content(res)
                if content is None:
                    err_streak += 1
                    denied_api = isinstance(res, dict) and res.get("__err") == "denied"
                    if denied_api or is_access_denied(page):
                        wait = min(60, 8 * err_streak)
                        print(f"{pg_no}페이지 Access Denied. {wait}초 쉰 뒤 재시도 "
                              f"(페이지는 건드리지 않음)...", flush=True)
                        time.sleep(wait)
                        continue
                    print(f"{pg_no}페이지에서 세션 만료/응답 이상. 쿠키만 가볍게 줄이고 8초 뒤 재시도...")
                    trim_fat_cookies(page)
                    time.sleep(8)
                    continue
                if not content:
                    break
                for it in content:
                    key = it.get("skuId")
                    if key in seen:
                        continue
                    seen.add(key)
                    rows.append(it)
                err_streak = 0
            except Exception as e:
                err_streak += 1
                print(f"{pg_no}페이지 오류({e}). 3초 후 재시도...")
                page = revive_page(ctx, page)  # 탭이 닫혔으면 새 탭으로 복구
                time.sleep(3)
                continue
            if pg_no % 5 == 0 or pg_no == total_pages:
                el = time.time() - t0
                print(f"진행 {pg_no}/{total_pages}페이지, 누적 {len(rows)}개 ({el:.0f}초)", flush=True)
            pg_no += 1
            time.sleep(0.4)

        print(f"수집 완료: {len(rows)}개 (전체 {total}개)")

        # ── 상세 정보(최소구매수량 등) 수집 — 캐시 사용, 새 상품만 조회 ──
        cache_path = BASE / "sku_detail_cache.json"
        cache = {}
        if cache_path.exists():
            try:
                cache = json.loads(cache_path.read_text(encoding="utf-8"))
            except Exception:
                cache = {}
        refresh_all = "전체갱신" in sys.argv
        ids = [str(it.get("skuId")) for it in rows if it.get("skuId") is not None]
        todo = [i for i in ids if refresh_all or i not in cache]
        print(f"상세 조회 대상: {len(todo)}개 (캐시 보유 {len(ids) - len(todo)}개)")
        BATCH = 6
        DELAY = 0.4
        t1 = time.time()
        gave_up = False
        consecutive_bad = 0
        bi = 0
        last_save = 0
        while bi < len(todo):
            batch = todo[bi:bi + BATCH]
            try:
                res = page.evaluate(DETAIL_JS, batch) or {}
            except Exception as e:
                print(f"상세 배치 오류({e}). 10초 후 재시도...", flush=True)
                page = revive_page(ctx, page)  # 탭이 닫혔으면 새 탭으로 복구
                time.sleep(10)
                continue
            errs = 0
            for sid, body in res.items():
                if isinstance(body, dict) and "__err" not in body:
                    cache[str(sid)] = body
                else:
                    errs += 1
            if errs >= max(1, len(batch) // 2):
                consecutive_bad += 1
                if consecutive_bad > 12:
                    print("차단이 오래 지속됩니다. 지금까지 결과로 CSV를 만들고 종료합니다.", flush=True)
                    print("(나중에 다시 실행하면 남은 것만 이어서 수집합니다)", flush=True)
                    gave_up = True
                    break
                wait = min(300, 30 * consecutive_bad)
                print(f"서버 차단 감지({errs}/{len(batch)} 실패) → {wait}초 쉬었다가 재개합니다"
                      f" [진행 {bi}/{len(todo)}]", flush=True)
                cache_path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
                trim_fat_cookies(page)
                time.sleep(wait)
                continue  # 같은 배치 재시도
            consecutive_bad = 0
            bi += len(batch)
            if bi - last_save >= 500 or bi >= len(todo):
                el = time.time() - t1
                print(f"상세 진행 {bi}/{len(todo)}개 ({el:.0f}초, 캐시 {len(cache)}개)", flush=True)
                cache_path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
                last_save = bi
            time.sleep(DELAY)
        cache_path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
        missing = len([i for i in ids if i not in cache])
        print(f"상세 조회 {'중단' if gave_up else '완료'} (캐시 {len(cache)}개, 미수집 {missing}개)", flush=True)

        out = BASE / "쿠팡_공급SKU_최신.csv"
        dated = BASE / f"쿠팡_공급SKU_{datetime.date.today().isoformat()}.csv"
        header = ["SKU ID", "요청", "상품명", "바코드", "발주가능상태", "담당 BM", "발주담당자", "이미지url", "최소구매수량"]
        for path in (out, dated):
            with path.open("w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(header)
                for it in rows:
                    det = cache.get(str(it.get("skuId")), {})
                    moq = det.get("moq")
                    w.writerow([
                        it.get("skuId", ""),
                        it.get("issueCount", ""),
                        it.get("skuName", ""),
                        it.get("barCode", ""),
                        it.get("orderStatus", ""),
                        it.get("mdName", ""),
                        it.get("scmName", ""),
                        img_https(it.get("imagePath", "")),
                        "" if moq is None else moq,
                    ])
        print(f"저장 완료:\n  {out}\n  {dated}")


def run_collection():
    main()


if __name__ == "__main__":
    try:
        run_collection()
    except RuntimeError as exc:
        print(f"RESULT: 수집 실행 보류 · {exc}", flush=True)
        sys.exit(2)

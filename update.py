# -*- coding: utf-8 -*-
"""수강생용 대시보드 업데이트.

update_source.txt 에 적힌 GitHub 저장소(main 브랜치)의 zip을 받아 프로그램 파일만 덮어쓴다.
내 설정(settings.json), 로그인(chrome_profile), 수집한 CSV 등은 건드리지 않는다.
"""
from __future__ import annotations

import io
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(__file__).parent
SOURCE_FILE = BASE / "update_source.txt"
# 업데이트로 덮어쓰지 않는 것 (사람마다 다른 데이터)
KEEP = {
    "settings.json", "update_source.txt", "chrome_profile", "쉽먼트",
    "_ship_dl", "_order_inbox_dl", "po_pending.json", "po_apis.json",
    "drive_ship_cache.json", "sku_detail_cache.json", "service_account.json",
    "secrets.toml", ".tts_cache",
}


def local_version() -> str:
    p = BASE / "VERSION"
    return p.read_text(encoding="utf-8").strip() if p.exists() else "?"


def main() -> int:
    repo = SOURCE_FILE.read_text(encoding="utf-8").strip() if SOURCE_FILE.exists() else ""
    if not repo:
        print("update_source.txt 에 저장소 주소가 없습니다. 강사에게 문의해 주세요.")
        return 1
    repo = repo.removesuffix(".git").rstrip("/")
    url = f"{repo}/archive/refs/heads/main.zip"
    print(f"현재 버전: {local_version()}")
    print(f"받는 중: {url}")
    try:
        with urllib.request.urlopen(url, timeout=60) as resp:
            data = resp.read()
    except Exception as e:
        print(f"다운로드 실패: {e}")
        return 1

    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        root = zf.namelist()[0].split("/")[0]
        new_version = ""
        count = 0
        for info in zf.infolist():
            rel = info.filename[len(root) + 1:]
            if not rel or info.is_dir():
                continue
            if any(part in KEEP for part in rel.split("/")) or rel.startswith("쿠팡_공급SKU_"):
                continue
            dest = BASE / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, dest.open("wb") as out:
                shutil.copyfileobj(src, out)
            count += 1
            if rel == "VERSION":
                new_version = dest.read_text(encoding="utf-8").strip()

    print(f"파일 {count}개 갱신 → 버전 {new_version or '?'}")
    print("필요한 패키지를 확인합니다...")
    subprocess.call([sys.executable, "-m", "pip", "install", "-q", "-r", str(BASE / "requirements.txt")])
    print("업데이트 완료. 대시보드를 다시 실행해 주세요.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

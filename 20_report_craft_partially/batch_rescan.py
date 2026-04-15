"""
12개 프로젝트 일괄 재스캔 + 재prepare

- scan_result.json 백업 (scan_result_backup.json)
- 새 코드로 재스캔 (캐시 재사용, OCR 없음)
- 재prepare (prompt_for_claude.md 재생성)
- 이전 scan_result.json vs 새 scan_result.json 비교 → 신규 발견 항목 리포트

사용: python batch_rescan.py
"""

import io
import json
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
PYTHON = sys.executable


def run(cmd: list[str]) -> tuple[int, str]:
    result = subprocess.run(
        cmd, cwd=str(BASE_DIR),
        capture_output=True, encoding="utf-8", errors="replace"
    )
    out = (result.stdout or "") + (result.stderr or "")
    return result.returncode, out


def backup_scan_result(proj_dir: Path):
    src = proj_dir / "scan_result.json"
    dst = proj_dir / "scan_result_backup.json"
    if src.exists() and not dst.exists():
        shutil.copy2(src, dst)
        return True
    return False


def compare_scan_results(proj_dir: Path) -> list[dict]:
    """백업과 신규 scan_result.json 비교 → 신규 발견 항목"""
    old_path = proj_dir / "scan_result_backup.json"
    new_path = proj_dir / "scan_result.json"
    if not old_path.exists() or not new_path.exists():
        return []
    old = json.loads(old_path.read_text(encoding="utf-8"))
    new = json.loads(new_path.read_text(encoding="utf-8"))
    old_files = {Path(i.get("file_path", "")).name for i in old.get("items", [])}
    new_items = []
    for item in new.get("items", []):
        fname = Path(item.get("file_path", "")).name
        if fname not in old_files:
            new_items.append(item)
    return new_items


def main():
    print("=" * 65)
    print("  일괄 재스캔 + 재prepare")
    print("=" * 65)
    print()

    results = []

    for proj_dir in sorted(OUTPUT_DIR.iterdir()):
        if not proj_dir.is_dir():
            continue
        scan_result_path = proj_dir / "scan_result.json"
        if not scan_result_path.exists():
            continue

        proj_name = proj_dir.name
        print(f"[{proj_name}]")

        # 백업
        backed_up = backup_scan_result(proj_dir)
        if backed_up:
            print(f"  백업 완료 -> scan_result_backup.json")

        # vendor_dirs 읽기
        d = json.loads(scan_result_path.read_text(encoding="utf-8"))
        vendor_dirs = d.get("vendor_dirs") or [d.get("vendor_dir", "")]
        if not vendor_dirs or not vendor_dirs[0]:
            print(f"  [SKIP] vendor_dirs 없음")
            continue

        # 재스캔
        scan_cmd = [PYTHON, "main.py", "scan"] + vendor_dirs + ["--project", proj_name]
        print(f"  scan 실행 중...")
        rc, out = run(scan_cmd)
        if rc != 0:
            print(f"  [ERROR] scan 실패: {out[-300:]}")
            results.append({"name": proj_name, "scan_ok": False})
            continue
        print(f"  scan 완료")

        # 재prepare
        prepare_cmd = [PYTHON, "main.py", "prepare", proj_name]
        print(f"  prepare 실행 중...")
        rc, out = run(prepare_cmd)
        if rc != 0:
            print(f"  [ERROR] prepare 실패: {out[-300:]}")
            results.append({"name": proj_name, "scan_ok": True, "prepare_ok": False})
            continue
        print(f"  prepare 완료")

        # 신규 항목 비교
        new_items = compare_scan_results(proj_dir)
        if new_items:
            print(f"  [NEW] 신규 발견 항목 {len(new_items)}건:")
            for item in new_items:
                fname = Path(item.get("file_path", "")).name
                subj = item.get("subject", "")[:40]
                date = item.get("date", "")
                print(f"    - {date} | {subj} | {fname}")
        else:
            print(f"  신규 항목 없음")

        results.append({
            "name": proj_name,
            "scan_ok": True,
            "prepare_ok": True,
            "new_items": len(new_items),
        })
        print()

    print("=" * 65)
    print("완료 요약:")
    for r in results:
        status = "OK" if r.get("prepare_ok") else ("SCAN_FAIL" if not r.get("scan_ok") else "PREPARE_FAIL")
        new = r.get("new_items", 0)
        flag = f"  [신규 {new}건]" if new else ""
        print(f"  {status} | {r['name']}{flag}")
    print("=" * 65)


if __name__ == "__main__":
    main()

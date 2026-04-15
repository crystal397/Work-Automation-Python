"""
main.py — 공수 취합 파이프라인 실행
-------------------------------------
실행:
    python main.py

흐름:
    1. input/  하위 모든 .xlsx 파일 스캔 → 노무비 시트 파싱
    2. 연월별 인원 취합 (중복 이름 합산)
    3. template/ 의 산출내역서 원본을 복사, 노임 시트에 인원·출역 기입
    4. AD~BK 열 COUNTIF 수식 입력
    5. output/ 에 완성 파일 저장
"""

import sys
import io
from pathlib import Path

# Windows 콘솔 UTF-8
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE_DIR     = Path(__file__).parent
INPUT_DIR    = BASE_DIR / "input"
TEMPLATE_DIR = BASE_DIR / "template"
OUTPUT_DIR   = BASE_DIR / "output"

from aggregator    import collect_all
from filler        import fill_template
from formula_writer import process_formulas


def main():
    print("=" * 60)
    print("  공수 취합 파이프라인 시작")
    print("=" * 60)

    # ── Step 1·2: 수신자료 스캔 및 취합 ─────────────────────────────
    print("\n[Step 1] input/ 폴더 스캔 중...")
    data = collect_all(INPUT_DIR)

    if not data:
        print("[ERR] 처리할 데이터가 없습니다.")
        print("      input/ 폴더에 업체별 xlsx 파일을 넣어주세요.")
        return

    # ── Step 3: 템플릿에 인원·출역 기입 ─────────────────────────────
    print("\n[Step 2] 산출내역서 노임 시트 기입 중...")
    out_path = fill_template(TEMPLATE_DIR, OUTPUT_DIR, data)

    # ── Step 4: COUNTIF 수식 입력 ────────────────────────────────────
    print("\n[Step 3] COUNTIF 수식 입력 중...")
    process_formulas(out_path)

    print("\n" + "=" * 60)
    print(f"  완료: {out_path.name}")
    print("=" * 60)


if __name__ == "__main__":
    main()

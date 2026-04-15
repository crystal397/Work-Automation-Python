"""
aggregator.py — input/ 폴더 전체 스캔 및 데이터 취합
-----------------------------------------------------
input/ 하위 모든 .xlsx 파일을 재귀 탐색 → 노무비 시트 파싱
→ {(year, month): [person_dict, ...]} 형태로 반환
"""

from pathlib import Path
from readers.common import read_file
from readers.pdf_reader import read_pdf


def collect_all(input_dir: str | Path) -> dict:
    """
    input/ 폴더를 재귀 탐색하여 전체 출역 데이터 수집 (xlsx + pdf)

    Returns:
        {
            (year, month): [
                {'name': str, 'year': int, 'month': int, 'attendance': {day: float}},
                ...
            ]
        }
    """
    input_path = Path(input_dir)
    all_files = sorted(
        list(input_path.rglob("*.xlsx")) +
        list(input_path.rglob("*.pdf"))
    )

    if not all_files:
        print(f"[WARN] {input_path} 에 xlsx/pdf 파일이 없습니다.")
        return {}

    aggregated: dict[tuple, list] = {}

    for f in all_files:
        print(f"  파일: {f.relative_to(input_path)}")
        if f.suffix.lower() == '.pdf':
            records = read_pdf(f)
        else:
            records = read_file(f)

        for rec in records:
            key = (rec['year'], rec['month'])
            aggregated.setdefault(key, []).append(rec)

    # 같은 연월 내 중복 이름 처리: 출역일 합산 (동명이인 주의)
    merged: dict[tuple, list] = {}
    for (year, month), people in aggregated.items():
        # name → attendance 합산
        by_name: dict[str, dict] = {}
        for p in people:
            name = p['name']
            if name not in by_name:
                by_name[name] = dict(p['attendance'])
            else:
                # 같은 날짜가 이미 있으면 최대값 사용 (중복 집계 방지)
                for day, val in p['attendance'].items():
                    by_name[name][day] = max(by_name[name].get(day, 0), val)

        merged[(year, month)] = [
            {'name': name, 'year': year, 'month': month, 'attendance': att}
            for name, att in by_name.items()
        ]

    total = sum(len(v) for v in merged.values())
    periods = sorted(merged.keys())
    print(f"\n[집계] {len(periods)}개 기간, 총 {total}명·월 레코드")
    for y, m in periods:
        print(f"  {y}년 {m}월: {len(merged[(y, m)])}명")

    return merged

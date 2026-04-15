"""
readers/pdf_reader.py — PDF 노무비 파일 파서
--------------------------------------------
지원 포맷:
  A) 일용노무비지급명세서 (스마트에스지 형식)
     - 날짜 1~15 / 16~31 그리드 + 1인당 2행 구조
     - pdfplumber 테이블 추출

  B) 급여명세서 (월별 임금명세서)
     - 일별 공수 없음 → 성명·연월만 추출 후 경고 출력

  C) 스캔본 → 텍스트 없음 → 스킵

반환 형식 (일용노무비만):
  [{'name': str, 'year': int, 'month': int, 'attendance': {day: float}}, ...]
"""

import re
import calendar
import pdfplumber
from pathlib import Path


# ── 유틸 ──────────────────────────────────────────────────────────────────────
def _to_float(v) -> float | None:
    if v is None:
        return None
    s = str(v).strip().replace(',', '').replace(' ', '')
    if s in ('', '0', '0.0', '-'):
        return None
    try:
        f = float(s)
        return f if f > 0 else None
    except ValueError:
        return None


def _extract_yearmonth(text: str):
    m = re.search(r'(\d{4})년\s*(\d{1,2})월', text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


# ── 포맷 감지 ─────────────────────────────────────────────────────────────────
def _is_scanned(doc) -> bool:
    """처음 3페이지 텍스트가 모두 없으면 스캔본으로 판단"""
    check = min(3, len(doc.pages))
    for page in doc.pages[:check]:
        if page.extract_text():
            return False
    return True


def _detect_format(page) -> str:
    """
    페이지 첫 테이블을 보고 포맷 판별
    'illyong'  : 일용노무비지급명세서 (날짜 그리드 있음)
    'salary'   : 급여명세서 (지급내역/공제내역 있음)
    'unknown'  : 알 수 없음
    """
    text = page.extract_text() or ''
    if '일용노무비' in text or '출역상황' in text or '출역일수' in text:
        return 'illyong'
    if '급여명세서' in text or '지급내역' in text or '공제내역' in text:
        return 'salary'
    # 테이블에서 날짜 숫자 배열 확인
    tables = page.extract_tables()
    for table in tables:
        for row in table:
            nums = [str(c).strip() for c in row if c is not None]
            # 1,2,3,4,5,...,15 연속 존재
            if all(str(i) in nums for i in range(1, 16)):
                return 'illyong'
    return 'unknown'


# ── 일용노무비 파싱 ───────────────────────────────────────────────────────────
def _find_date_col_in_table(table) -> tuple[int | None, int | None]:
    """
    테이블 행 중 날짜 헤더(1~15 숫자)를 찾아 (header_row_idx, date_start_col) 반환
    """
    for row_idx, row in enumerate(table):
        col_map = {}
        for col_idx, v in enumerate(row):
            if v is not None:
                s = str(v).strip()
                if s.isdigit():
                    col_map[int(s)] = col_idx
        if all(i in col_map for i in range(1, 16)):
            return row_idx, col_map[1]
    return None, None


def _parse_illyong_table(table) -> list:
    """
    일용노무비지급명세서 테이블 한 장 파싱
    Returns: list of person dicts
    """
    # 연/월 추출 (상위 3행 텍스트에서)
    year, month = None, None
    for row in table[:4]:
        for cell in row:
            if cell:
                y, m = _extract_yearmonth(str(cell))
                if y:
                    year, month = y, m
                    break
        if year:
            break

    if not (year and month):
        return []

    # 날짜 헤더 행 위치 찾기
    hrow1_idx, date_start_col = _find_date_col_in_table(table)
    if hrow1_idx is None:
        return []

    max_day = calendar.monthrange(year, month)[1]
    results = []
    i = hrow1_idx + 2  # 첫 데이터 행 (헤더row1, row2 건너뜀)

    while i < len(table):
        row = table[i]

        # 이름 열 탐지: col 1 이 한글 이름
        name = row[1] if len(row) > 1 else None
        if not name or not isinstance(name, str):
            i += 1
            continue
        name = name.strip()
        if not re.search(r'[가-힣]{2,}', name):
            i += 1
            continue
        # 합계·메모 행 제외
        if any(kw in (row[0] or '') for kw in ('메모', '계', '*')):
            break
        if name in ('계', '합계', '소계'):
            i += 1
            continue

        attendance = {}

        # 1~15일 (현재 행)
        for day in range(1, 16):
            col = date_start_col + (day - 1)
            v = _to_float(row[col] if col < len(row) else None)
            if v:
                attendance[day] = v

        # 16~31일 (다음 행)
        if i + 1 < len(table):
            next_row = table[i + 1]
            for day in range(16, max_day + 1):
                col = date_start_col + (day - 16)
                v = _to_float(next_row[col] if col < len(next_row) else None)
                if v:
                    attendance[day] = v

        if attendance:
            results.append({
                'name': name,
                'year': year,
                'month': month,
                'attendance': attendance,
            })

        i += 2  # 1인당 2행

    return results


# ── 급여명세서 파싱 (경고용) ──────────────────────────────────────────────────
def _parse_salary_page(page) -> list:
    """
    급여명세서에서 성명·연월만 추출하고 경고 출력.
    일별 공수가 없어 출역 데이터로 변환 불가.
    """
    text = page.extract_text() or ''
    year, month = _extract_yearmonth(text)

    # 사원명 추출
    m = re.search(r'사원명[:\s]*([가-힣]{2,5})', text)
    name = m.group(1) if m else None

    if name and year:
        print(f"    [급여명세서] {year}년 {month}월 / {name} — "
              f"일별 공수 없음, 노임 시트 반영 불가 (수동 확인 필요)")
    return []


# ── PDF 파일 전체 파싱 ────────────────────────────────────────────────────────
def read_pdf(path: Path) -> list:
    """
    PDF 파일에서 노무비 출역 데이터 추출

    Returns: list of person dicts (일용노무비 형식만)
    """
    try:
        doc = pdfplumber.open(str(path))
    except Exception as e:
        print(f"  [ERR] {path.name} 열기 실패: {e}")
        return []

    # 스캔본 감지
    all_text = ''.join(p.extract_text() or '' for p in doc.pages[:5])
    if not all_text.strip():
        print(f"  [SKIP] {path.name} — 스캔본(텍스트 없음), 수동 입력 필요")
        doc.close()
        return []

    all_records = []
    prev_ym = (None, None)    # 페이지 간 연/월 유지 (연속 페이지)

    for page_num, page in enumerate(doc.pages, 1):
        fmt = _detect_format(page)

        if fmt == 'illyong':
            tables = page.extract_tables()
            for table in tables:
                records = _parse_illyong_table(table)
                if records:
                    ym = (records[0]['year'], records[0]['month'])
                    if ym != prev_ym:
                        print(f"    [p{page_num}] {ym[0]}년 {ym[1]}월 — {len(records)}명 추출")
                        prev_ym = ym
                    all_records.extend(records)

        elif fmt == 'salary':
            _parse_salary_page(page)

        # 'unknown' 은 조용히 스킵

    doc.close()
    return all_records

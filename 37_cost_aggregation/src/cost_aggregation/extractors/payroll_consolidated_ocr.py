"""일반직 급여명세서 합본 PDF (635쪽 이미지) 한국어 OCR 자동 추출.

PDF 구조 (C공구(철도) ㈜A건설 일반직 양식)
- 1쪽 = 1명 × 1개월
- 텍스트:
    '<YYYY년 N월 급여명세서>'
    '성명' '직급' '소속' '근무일수' '지급일자'
    <사번 7자리> <성명 OCR — 한자/한글 정확도 변동>
    'C공구(철도)현장' <근무일수> <지급일자 YYYYMMDD>
    '소득총액' <금액> '공제계' <금액> '수령액' <금액>
    매월지급/부정기지급 명세 + 공제 명세

OCR 캐시: 같은 PDF·DPI 조합 결과를 json 으로 저장 (재실행 시 즉시).
"""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field
from pathlib import Path

import fitz
from rapidocr_onnxruntime import RapidOCR

from contract_meta.audit import make_source
from contract_meta.models import Sourced


_DEFAULT_KOREAN_MODEL = Path(__file__).resolve().parents[4] / "25_claim_analysis" / "ocr_models" / "korean_PP-OCRv4_rec_infer.onnx"
_DEFAULT_KOREAN_DICT = _DEFAULT_KOREAN_MODEL.with_name("korean_dict.txt")


@dataclass
class OcrPage:
    page_no: int                                  # 1-base
    boxes: list[tuple[str, float]] = field(default_factory=list)


@dataclass
class ConsolidatedSlip:
    page_no: int
    year: int | None
    month: int | None
    employee_id: str | None
    work_days: int | None
    pay_date: str | None                          # YYYYMMDD
    gross_income_krw: int | None
    deduction_krw: int | None
    net_pay_krw: int | None


def ocr_consolidated(
    pdf_path: str | Path,
    *,
    dpi: int = 200,
    cache_path: str | Path | None = None,
    page_range: tuple[int, int] | None = None,
    verbose: bool = True,
) -> list[OcrPage]:
    """전체 또는 부분 페이지를 OCR 후 OcrPage[] 반환. 캐시 파일 있으면 즉시 로드."""
    pdf_path = Path(pdf_path)
    if cache_path:
        cache_path = Path(cache_path)
        if cache_path.exists():
            data = json.loads(cache_path.read_text(encoding="utf-8"))
            return [OcrPage(page_no=p["page_no"], boxes=[(b[0], float(b[1])) for b in p["boxes"]]) for p in data]

    if not _DEFAULT_KOREAN_MODEL.exists():
        raise FileNotFoundError(f"한국어 OCR 모델 없음: {_DEFAULT_KOREAN_MODEL}")

    ocr = RapidOCR(
        rec_model_path=str(_DEFAULT_KOREAN_MODEL),
        rec_keys_path=str(_DEFAULT_KOREAN_DICT),
    )
    doc = fitz.open(str(pdf_path))
    start = (page_range[0] if page_range else 1)
    end = (page_range[1] if page_range else len(doc))

    pages: list[OcrPage] = []
    t0 = time.time()
    for pno in range(start, end + 1):
        pix = doc[pno - 1].get_pixmap(dpi=dpi)
        result, _ = ocr(pix.tobytes("png"))
        boxes = [(text, float(score)) for _box, text, score in (result or [])]
        pages.append(OcrPage(page_no=pno, boxes=boxes))
        if verbose and (pno - start + 1) % 10 == 0:
            elapsed = time.time() - t0
            done = pno - start + 1
            rate = done / elapsed
            remain = (end - pno) / rate if rate else 0
            print(f"  ... {done}/{end - start + 1}p, {elapsed:.0f}s elapsed, ETA {remain:.0f}s")
    doc.close()

    if cache_path:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(
            json.dumps([{"page_no": p.page_no, "boxes": p.boxes} for p in pages], ensure_ascii=False),
            encoding="utf-8",
        )
    return pages


# ─────────────────────────────────────────────────────────────────
# 페이지 → 슬립 추출
# ─────────────────────────────────────────────────────────────────

_YEAR_MONTH_PAT = re.compile(r"(\d{4})년\s*(\d{1,2})월")
_EMPLOYEE_ID_PAT = re.compile(r"^\d{7,8}$")
_PAY_DATE_PAT = re.compile(r"^(20\d{6})$")


def parse_slip(p: OcrPage) -> ConsolidatedSlip:
    year = month = None
    employee_id = None
    work_days = None
    pay_date = None

    texts = [t for t, s in p.boxes]
    big_ints: list[int] = []
    for t, s in p.boxes:
        if s < 0.7:
            continue
        # 연·월
        m = _YEAR_MONTH_PAT.search(t)
        if m:
            year = int(m.group(1))
            month = int(m.group(2))
        # 사번 (7~8자리)
        if employee_id is None and _EMPLOYEE_ID_PAT.fullmatch(t.strip()):
            employee_id = t.strip()
            continue
        # 지급일자 (YYYYMMDD)
        if pay_date is None and _PAY_DATE_PAT.fullmatch(t.strip()):
            pay_date = t.strip()
            continue
        # 큰 정수 (5자리 이상, 콤마 없음 — OCR이 콤마 누락 가능)
        s2 = t.strip().replace(",", "")
        if s2.isdigit() and len(s2) >= 5:
            big_ints.append(int(s2))

    # 근무일수 — 단독 1~2자리 숫자, 1~31
    for t, s in p.boxes:
        if s < 0.7 or work_days is not None:
            continue
        s2 = t.strip()
        if s2.isdigit() and 1 <= int(s2) <= 31 and len(s2) <= 2:
            work_days = int(s2)
            break

    # 큰 정수들 중 첫 3개 (정렬 기준) = 소득총액 / 공제계 / 수령액 휴리스틱
    big_ints.sort(reverse=True)
    gross = big_ints[0] if big_ints else None
    deduction = big_ints[1] if len(big_ints) >= 2 else None
    net = big_ints[2] if len(big_ints) >= 3 else None

    return ConsolidatedSlip(
        page_no=p.page_no, year=year, month=month,
        employee_id=employee_id, work_days=work_days,
        pay_date=pay_date,
        gross_income_krw=gross, deduction_krw=deduction, net_pay_krw=net,
    )

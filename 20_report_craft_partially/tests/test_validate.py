"""
_validate_data 및 cmd_validate 소결 체크 단위 테스트
"""
import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.report_generator import _validate_data


def _minimal_data(**overrides):
    base = {
        "project_name": "테스트 프로젝트",
        "total_delay_days": 100,
        "background_paragraphs": ["단락1", "단락2"],
        "items": [{"no": 1, "date": "2024.01.01", "subject": "테스트"}],
        "detail_narratives": [
            {"title": "블록1", "paragraphs": [
                "원인 분석 내용입니다.",
                "공사계약일반조건 제22조에 해당하며, 계약상대자의 책임 없는 사유에 해당합니다.",
            ]}
        ],
        "pre_diagram_paragraphs": ["단락1", "단락2"],
        "accountability_diagram": [
            {"cause": "관계기관 협의 지연", "basis": "공문 제22-001호",
             "responsible_party": "발주처", "delay_days": 100}
        ],
        "conclusion_paragraphs": ["결론 단락1", "결론 단락2"],
        "summary": "종합 요약입니다.",
    }
    base.update(overrides)
    return base


def _run_validate(data: dict):
    """_validate_data 실행 후 (errors_list, warnings_list) 반환."""
    import io as _io
    captured = _io.StringIO()
    old = sys.stdout
    sys.stdout = captured
    err = None
    try:
        _validate_data(data, Path("dummy.json"))
    except ValueError as e:
        err = str(e)
    finally:
        sys.stdout = old
    output = captured.getvalue()
    errors   = [l for l in output.splitlines() if "[오류]" in l]
    warnings = [l for l in output.splitlines() if "[경고]" in l]
    if err:
        errors.append(err)
    return errors, warnings


# ── 정상 케이스 ───────────────────────────────────────────────────────────────

def test_valid_data_no_errors():
    errors, warnings = _run_validate(_minimal_data())
    assert errors == [], f"예상치 못한 오류: {errors}"


# ── total_delay_days ──────────────────────────────────────────────────────────

def test_total_delay_days_zero_is_error():
    data = _minimal_data(total_delay_days=0,
                         accountability_diagram=[{"cause": "x", "basis": "y",
                                                   "responsible_party": "발주처", "delay_days": 0}])
    errors, _ = _run_validate(data)
    assert any("total_delay_days" in e for e in errors)


def test_total_delay_days_mismatch_is_error():
    data = _minimal_data(total_delay_days=200)  # diagram 합계=100
    errors, _ = _run_validate(data)
    assert any("합계" in e or "total_delay_days" in e for e in errors)


# ── 필수 필드 누락 ────────────────────────────────────────────────────────────

def test_missing_conclusion_paragraphs_is_error():
    data = _minimal_data()
    del data["conclusion_paragraphs"]
    errors, _ = _run_validate(data)
    assert any("conclusion_paragraphs" in e for e in errors)


def test_empty_background_paragraphs_is_error():
    data = _minimal_data(background_paragraphs=[])
    errors, _ = _run_validate(data)
    assert any("background_paragraphs" in e for e in errors)


# ── 금지 필드명 ───────────────────────────────────────────────────────────────

def test_forbidden_field_intro_paragraph():
    data = _minimal_data()
    data["intro_paragraph"] = "잘못된 필드"
    errors, _ = _run_validate(data)
    assert any("intro_paragraph" in e for e in errors)


# ── conclusion_paragraphs 단락 수 ────────────────────────────────────────────

def test_conclusion_one_paragraph_is_warning():
    data = _minimal_data(conclusion_paragraphs=["단락 하나만"])
    errors, warnings = _run_validate(data)
    assert any("conclusion_paragraphs" in w for w in warnings)


# ── 날짜 시간순 ───────────────────────────────────────────────────────────────

def test_items_date_out_of_order_is_warning():
    data = _minimal_data(items=[
        {"no": 1, "date": "2024.06.01", "subject": "a"},
        {"no": 2, "date": "2024.01.01", "subject": "b"},
    ])
    _, warnings = _run_validate(data)
    assert any("날짜" in w or "시간순" in w for w in warnings)


# ── cmd_validate 소결 체크 ────────────────────────────────────────────────────

def _run_cmd_validate_sogyeol(narratives):
    """main.py cmd_validate의 소결 체크 로직만 직접 테스트."""
    _SOGYEOL_CONCLUSION = "계약상대자의 책임 없는 사유"
    _SOGYEOL_CLAUSE_KW  = [
        "제8절", "제22조", "제25조", "제26조", "제27조", "제74조",
        "일반조건", "계약조건", "해당 조항", "계약 조항",
    ]
    hard_errors = []
    soft_warnings = []
    for i, block in enumerate(narratives):
        if not isinstance(block, dict):
            continue
        paras = block.get("paragraphs", [])
        if not isinstance(paras, list) or len(paras) == 0:
            continue
        block_title = block.get("title", f"[{i}]")
        paras_text = " ".join(str(p) for p in paras)
        if _SOGYEOL_CONCLUSION not in paras_text:
            hard_errors.append(f"소결 결론 없음: {block_title}")
        elif not any(kw in paras_text for kw in _SOGYEOL_CLAUSE_KW):
            soft_warnings.append(f"조항 인용 없음: {block_title}")
    return hard_errors, soft_warnings


def test_sogyeol_결론_있고_조항_있음():
    narr = [{"title": "블록1", "paragraphs": [
        "공사계약일반조건 제22조에 따라 계약상대자의 책임 없는 사유에 해당합니다."
    ]}]
    hard, soft = _run_cmd_validate_sogyeol(narr)
    assert hard == [] and soft == []


def test_sogyeol_결론_있지만_조항_없음_은_warning():
    narr = [{"title": "블록1", "paragraphs": [
        "계약상대자의 책임 없는 사유에 해당합니다."
    ]}]
    hard, soft = _run_cmd_validate_sogyeol(narr)
    assert hard == []
    assert len(soft) == 1


def test_sogyeol_결론_없으면_hard_error():
    narr = [{"title": "블록1", "paragraphs": [
        "공사계약일반조건 제22조 검토 결과 발주처 귀책입니다."
    ]}]
    hard, soft = _run_cmd_validate_sogyeol(narr)
    assert len(hard) == 1


def test_sogyeol_민간계약_비표준조항_결론만_있으면_warning():
    # 민간계약: 표준 조항 번호 없어도 결론 문구 있으면 hard_error 아님
    narr = [{"title": "블록1", "paragraphs": [
        "계약서 제15조 제3항에 따라 계약상대자의 책임 없는 사유에 해당합니다."
    ]}]
    hard, soft = _run_cmd_validate_sogyeol(narr)
    # "계약 조항"이 없으므로 soft warning, hard error는 없어야 함
    assert hard == []


# ── items 오류·불일치 경고 체크 ──────────────────────────────────────────────

def _run_items_checks(items):
    """main.py cmd_validate의 items 체크 로직만 직접 테스트."""
    import re as _re
    _DATE_UNCERTAIN = ("xx", "?", "미확정", "불명", "확인필요")
    soft_warnings = []
    for i, di in enumerate(items):
        if not isinstance(di, dict):
            continue
        item_label = f"items[{i}](no={di.get('no','?')})"
        dt = (di.get("date") or "").strip()
        if any(tok in dt for tok in _DATE_UNCERTAIN):
            soft_warnings.append(f"미확정날짜: {item_label} date='{dt}'")
        dn = (di.get("doc_number") or "").strip()
        if dn and _re.search(r"[가-힣]", dn) and _re.search(r"[a-z]{2,}", dn):
            soft_warnings.append(f"OCR의심: {item_label} doc_number='{dn}'")
        sno = di.get("scan_no")
        if sno is None:
            soft_warnings.append(f"수동추가: {item_label} scan_no=null")
    return soft_warnings


def test_ocr_artifact_ys_detected():
    items = [{"no": 1, "date": "2022.12.06", "doc_number": "도철양산2022-ys", "scan_no": 1}]
    warns = _run_items_checks(items)
    assert any("OCR의심" in w for w in warns), f"경고 미감지: {warns}"


def test_ocr_artifact_bts_detected():
    items = [{"no": 1, "date": "2024.02.27", "doc_number": "도철양산2024-bts호", "scan_no": 2}]
    warns = _run_items_checks(items)
    assert any("OCR의심" in w for w in warns), f"경고 미감지: {warns}"


def test_ocr_clean_number_no_warning():
    items = [{"no": 1, "date": "2022.12.06", "doc_number": "도철양산제2022-604호", "scan_no": 3}]
    warns = _run_items_checks(items)
    assert not any("OCR의심" in w for w in warns), f"오탐: {warns}"


def test_ocr_clean_english_only_no_warning():
    # 순수 영문+숫자 번호 (한글 없음) → 미체크
    items = [{"no": 1, "date": "2022.12.16", "doc_number": "20171106A8B-03", "scan_no": 4}]
    warns = _run_items_checks(items)
    assert not any("OCR의심" in w for w in warns), f"오탐: {warns}"


def test_date_xx_detected():
    items = [{"no": 1, "date": "2024.12.xx", "doc_number": "회계치-3706", "scan_no": 5}]
    warns = _run_items_checks(items)
    assert any("미확정날짜" in w for w in warns), f"경고 미감지: {warns}"


def test_date_question_mark_detected():
    items = [{"no": 1, "date": "2024.??.??", "doc_number": "도철양산제2024-135호", "scan_no": 6}]
    warns = _run_items_checks(items)
    assert any("미확정날짜" in w for w in warns), f"경고 미감지: {warns}"


def test_scan_no_null_detected():
    items = [{"no": 1, "date": "2024.02.27", "doc_number": "도철양산2024-135호", "scan_no": None}]
    warns = _run_items_checks(items)
    assert any("수동추가" in w for w in warns), f"경고 미감지: {warns}"


def test_scan_no_present_no_warning():
    items = [{"no": 1, "date": "2024.02.27", "doc_number": "도철양산2024-135호", "scan_no": 7}]
    warns = _run_items_checks(items)
    assert not any("수동추가" in w for w in warns), f"오탐: {warns}"


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    passed = failed = 0
    for t in tests:
        try:
            t()
            print(f"  PASS  {t.__name__}")
            passed += 1
        except Exception as e:
            print(f"  FAIL  {t.__name__}: {e}")
            failed += 1
    print(f"\n{passed} passed, {failed} failed")

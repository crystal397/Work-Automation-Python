"""
_extract_contract_changes 파싱 함수 단위 테스트
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.prompt_builder import _extract_contract_changes


def _item(subj="", full=""):
    return {"subject": subj, "full_text": full, "no": 1, "date": "2024.01.01", "file_path": "test.pdf"}


# ── 키워드 매칭 ──────────────────────────────────────────────────────────────

def test_no_match_returns_empty():
    items = [_item("일반 공문", "인허가 요청 건입니다.")]
    assert _extract_contract_changes(items) == []


def test_matches_변경계약서_keyword():
    items = [_item("공사변경계약서", "준공일자 변경 관련")]
    result = _extract_contract_changes(items)
    assert len(result) == 1


def test_matches_공기연장_keyword():
    items = [_item("공기연장 요청", "공기 연장 50일 계약기간 연장")]
    result = _extract_contract_changes(items)
    assert len(result) == 1


# ── 연장일수 추출 ─────────────────────────────────────────────────────────────

def test_days_기본_패턴():
    items = [_item("변경계약서", "준공기한 변경, 120일 연장")]
    result = _extract_contract_changes(items)
    assert result[0]["연장일수"] == "120일"


def test_days_공기연장_N일_패턴():
    items = [_item("변경계약서", "공기연장 85일 승인")]
    result = _extract_contract_changes(items)
    assert result[0]["연장일수"] == "85일"


def test_days_중복_제거():
    # 같은 일수가 여러 번 등장하면 중복 없이 한 번만
    items = [_item("변경계약서", "공기연장 60일, 60일 연장 처리")]
    result = _extract_contract_changes(items)
    assert result[0]["연장일수"] == "60일"


def test_days_복수_합산():
    items = [_item("변경계약서", "1차 30일 연장, 2차 45일 연장")]
    result = _extract_contract_changes(items)
    assert "30" in result[0]["연장일수"]
    assert "45" in result[0]["연장일수"]


# ── 차수 추출 ─────────────────────────────────────────────────────────────────

def test_ord_제N차_변경계약():
    items = [_item("변경계약서", "제3차 변경계약 체결")]
    result = _extract_contract_changes(items)
    assert result[0]["차수"] == "제3차"


def test_ord_제N회_변경():
    items = [_item("변경계약서", "제2회 변경계약서")]
    result = _extract_contract_changes(items)
    assert result[0]["차수"] == "제2차"


def test_ord_없으면_빈값():
    items = [_item("변경계약서", "준공기한 변경 안내")]
    result = _extract_contract_changes(items)
    assert result[0]["차수"] == ""


# ── 준공변경 추출 ─────────────────────────────────────────────────────────────

def test_change_date_명시_패턴():
    full = "당초 2024.06.30 → 변경 2024.12.31 준공기한 변경계약서"
    items = [_item("변경계약서", full)]
    result = _extract_contract_changes(items)
    assert "2024.06.30" in result[0]["준공변경"]
    assert "2024.12.31" in result[0]["준공변경"]


def test_change_date_fallback_날짜쌍():
    # 명시 패턴 없을 때 첫/마지막 날짜 쌍 fallback
    full = "변경계약서 체결 2023.12.01 관련 공문, 준공일 2025.03.30"
    items = [_item("변경계약서", full)]
    result = _extract_contract_changes(items)
    assert result[0]["준공변경"] != ""


def test_change_date_날짜_하나뿐이면_빈값():
    full = "변경계약서 체결 2024.01.15 준공기한 변경"
    items = [_item("변경계약서", full)]
    result = _extract_contract_changes(items)
    assert result[0]["준공변경"] == ""


# ── 변경사유 추출 ─────────────────────────────────────────────────────────────

def test_reason_추출():
    full = "변경계약서 체결\n변경 사유: 관계기관 협의 지연으로 인한 공기연장\n기타"
    items = [_item("변경계약서", full)]
    result = _extract_contract_changes(items)
    assert "협의 지연" in result[0]["변경사유"]


def test_reason_없으면_빈값():
    items = [_item("변경계약서", "준공기한 변경 2024.06.30")]
    result = _extract_contract_changes(items)
    assert result[0]["변경사유"] == ""


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

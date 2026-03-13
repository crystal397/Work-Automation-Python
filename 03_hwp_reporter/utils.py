"""
utils.py
공통 헬퍼 함수 모음
"""
from datetime import datetime


def move_to_field_safe(hwp, field_name: str) -> bool:
    """누름틀로 안전하게 이동 (여러 오버로드 시도)"""
    attempts = [
        lambda: hwp.MoveToField(field_name, True, True, False),
        lambda: hwp.MoveToField(field_name, True, True),
        lambda: hwp.MoveToField(field_name, True),
        lambda: hwp.MoveToField(field_name),
    ]
    for attempt in attempts:
        try:
            if attempt():
                return True
        except Exception:
            continue
    return False


def insert_text_to_hwp(hwp, text: str) -> None:
    """현재 한글 커서 위치에 텍스트 입력"""
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


def format_date(val) -> str:
    """날짜 값을 'YYYY. MM. DD.' 형식으로 변환"""
    if isinstance(val, datetime):
        return val.strftime("%Y. %m. %d.")
    return str(val) if val else ""


def format_percent(val) -> str:
    """소수 비율을 백분율 문자열(소수점 2자리)로 변환 (예: 0.1234 → '12.34')"""
    if val is None:
        return "0.00"
    try:
        return f"{float(val) * 100:.2f}"
    except (ValueError, TypeError):
        return "0.00"


def format_val(val) -> str:
    """엑셀 값을 한글 표용 문자열로 변환 (0이면 '-' 반환)"""
    if val is None or val == "":
        return "-"
    if isinstance(val, (int, float)):
        return "-" if val == 0 else format(int(round(val)), ",")
    return str(val).replace('\r', ' ').replace('\n', ' ').strip()


def format_val_nothing(val) -> str:
    """엑셀 값을 한글 표용 문자열로 변환 (0이면 빈 문자열 반환)"""
    if val is None or val == "":
        return ""
    if isinstance(val, (int, float)):
        return "" if val == 0 else format(int(round(val)), ",")
    return str(val).replace('\r', ' ').replace('\n', ' ').strip()

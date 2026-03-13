"""
hwp_writer.py
한글(HWP) 문서에 데이터를 입력하는 모듈
"""
from utils import move_to_field_safe, insert_text_to_hwp, format_val, format_val_nothing


# ─────────────────────────────────────────────
# 내부 헬퍼
# ─────────────────────────────────────────────

def _fill_vertical_table(hwp, field_name: str, values: list, formatter=format_val_nothing) -> bool:
    """
    세로 방향 표를 채우는 공통 함수.
    첫 번째 값은 누름틀에 직접 쓰고, 나머지는 아래 셀로 이동하며 입력.
    """
    if not move_to_field_safe(hwp, field_name):
        print(f"[오류] '{field_name}' 누름틀을 찾을 수 없습니다.")
        return False

    # 첫 번째 칸
    hwp.PutFieldText(field_name, formatter(values[0]))
    move_to_field_safe(hwp, field_name)

    # 나머지 칸
    for val in values[1:]:
        hwp.Run("Cancel")
        hwp.Run("TableLowerCell")
        insert_text_to_hwp(hwp, formatter(val))

    return True


def _fill_grid_table(hwp, field_name: str, rows: list, formatter=format_val_nothing) -> bool:
    """
    격자(행×열) 표를 채우는 공통 함수.
    첫 번째 행의 첫 번째 열은 누름틀에, 나머지는 커서 이동으로 입력.
    """
    if not rows:
        print(f"[경고] '{field_name}' 표에 입력할 데이터가 없습니다.")
        return False

    hwp.Run("MoveDocBegin")
    if not move_to_field_safe(hwp, field_name):
        print(f"[오류] '{field_name}' 누름틀을 찾을 수 없습니다.")
        return False

    col_count = len(rows[0])

    for r, row_values in enumerate(rows):
        if r > 0:
            hwp.Run("TableLowerCell")
            for _ in range(col_count - 1):
                hwp.Run("TableLeftCell")

        for c, cell_value in enumerate(row_values):
            text = formatter(cell_value)
            if r == 0 and c == 0:
                hwp.PutFieldText(field_name, text)
                move_to_field_safe(hwp, field_name)
                hwp.Run("Cancel")
            else:
                insert_text_to_hwp(hwp, text)

            if c < col_count - 1:
                hwp.Run("TableRightCell")

    return True


# ─────────────────────────────────────────────
# 공개 API
# ─────────────────────────────────────────────

def write_header_fields(hwp, header: dict) -> None:
    """본문 누름틀(기간, 기간수, 비율)에 값 입력"""
    hwp.PutFieldText("start_day", header["start_day"])
    hwp.PutFieldText("end_day",   header["end_day"])
    hwp.PutFieldText("duration",  header["duration"])

    # 비율은 표 외부에 있으므로 커서 이동 후 별도 입력
    for field in ("rate_san", "rate_go", "rate_il"):
        hwp.Run("Cancel")
        hwp.Run("MoveDocBegin")
        hwp.MoveToField(field, True, True, False)
        hwp.PutFieldText(field, header[field])


def write_main_table(hwp, values: list) -> None:
    """첫 번째 표 (전체 요약) 입력"""
    print("[1/6] 첫 번째 표 입력 중...")
    _fill_vertical_table(hwp, "table_start", values, formatter=format_val)


def write_indirect_labor_table(hwp, rows: list) -> None:
    """두 번째 표 (간접노무비 집계) 입력"""
    print(f"[2/6] 두 번째 표 입력 중... ({len(rows)}행)")
    _fill_grid_table(hwp, "second_table_start", rows)


def write_severance_table(hwp, rows: list, totals: dict) -> None:
    """세 번째 표 (퇴직금) 입력 + 합계 누름틀"""
    print(f"[3/6] 세 번째 표 입력 중... ({len(rows)}행)")
    hwp.Run("Cancel")
    _fill_grid_table(hwp, "third_table_start", rows)

    hwp.PutFieldText("third_total_salary",    format_val_nothing(totals["salary"]))
    hwp.PutFieldText("third_total_severance", format_val_nothing(totals["severance"]))
    hwp.PutFieldText("third_total_sum",       format_val_nothing(totals["sum"]))


def write_fourth_table(hwp, values: list) -> None:
    """네 번째 표 입력"""
    print("[4/6] 네 번째 표 입력 중...")
    _fill_vertical_table(hwp, "fourth_table_start", values)


def write_fifth_table(hwp, values: list) -> None:
    """다섯 번째 표 입력"""
    print("[5/6] 다섯 번째 표 입력 중...")
    _fill_vertical_table(hwp, "fifth_table_start", values)


def write_sixth_table(hwp, values: list) -> None:
    """여섯 번째 표 입력"""
    print("[6/6] 여섯 번째 표 입력 중...")
    _fill_vertical_table(hwp, "sixth_table_start", values)

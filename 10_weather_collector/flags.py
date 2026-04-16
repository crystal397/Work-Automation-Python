"""
작업불가일 판정 플래그 중앙 정의 모듈
────────────────────────────────────────
gui.py / main.py / analyzer.py / kma_client.py 가 모두 여기서 import.
플래그 추가·변경·기본값 수정은 이 파일만 손대면 됩니다.
"""

# ── 플래그 완전 정의 ──────────────────────────────────────────────
# (flag_id, 표시명, DB 컬럼/원시값 컬럼, 비교연산자, 기본기준값, 단위)
# 컬럼 = None : iscs 파싱 기반 범주형 플래그 → 수치 기준값 없음
FLAG_DEFS: list[tuple] = [
    ("is_rain_day",      "우천",            "precipitation",  ">=",  10.0, "mm"),
    ("is_wind_day",      "강풍",            "wind_max",       ">=",  14.0, "m/s"),
    ("is_wind_crane",    "크레인 제한",     "max_ins_wind",   ">=",  10.0, "m/s"),
    ("is_snow_day",      "적설",            "snow_depth",     ">=",   1.0, "cm"),
    ("is_heat_day",      "폭염",            "temp_max",       ">=",  35.0, "℃"),
    ("is_cold_day",      "한파",            "temp_min",       "<=", -10.0, "℃"),
    ("is_no_sunshine",   "일조 부족",       "sunshine_hours", "<",    2.0, "hr"),
    ("is_freeze_day",    "지면 동결",       "ground_temp",    "<=",   0.0, "℃"),
    ("is_high_evap_day", "증발 과다",       "evaporation",    ">=",  10.0, "mm"),
    ("rain_yn",          "강수 유무 (소량)", None,            None,  None, None),
    ("snow_yn",          "강설 유무",        None,            None,  None, None),
    ("fog_yn",           "안개",             None,            None,  None, None),
]

# 수치 플래그만 추출 → analyzer / kma_client 공용
# {flag_id: (db_column, operator, default_threshold)}
FLAG_COMPUTATIONS: dict[str, tuple[str, str, float]] = {
    flag_id: (col, op, default)
    for flag_id, _, col, op, default, _ in FLAG_DEFS
    if col is not None
}

# flag_id → 전체 정의 튜플 역색인 (main.py 등에서 flag_id로 메타 조회)
FLAG_BY_ID: dict[str, tuple] = {
    flag_id: (flag_id, label, col, op, default, unit)
    for flag_id, label, col, op, default, unit in FLAG_DEFS
}

# ── 공종 프리셋 ───────────────────────────────────────────────────
# 순서 보존 리스트 (main.py CLI 번호 메뉴 + gui.py OptionMenu 공용)
WORK_PRESET_ITEMS: list[tuple[str, list[str]]] = [
    ("토공사",           ["is_rain_day", "is_snow_day", "is_freeze_day", "is_cold_day"]),
    ("철근콘크리트공사", ["is_rain_day", "is_heat_day", "is_cold_day",
                         "is_freeze_day", "is_wind_day"]),
    ("타워크레인작업",   ["is_wind_crane", "is_wind_day", "fog_yn"]),
    ("도장·방수공사",   ["is_rain_day", "rain_yn", "is_no_sunshine",
                         "is_cold_day", "is_freeze_day"]),
    ("강구조물공사",    ["is_rain_day", "is_wind_day", "is_cold_day",
                         "is_freeze_day", "is_heat_day"]),
    ("포장공사",        ["is_rain_day", "is_snow_day", "is_freeze_day",
                         "is_cold_day", "is_heat_day"]),
]

# gui.py용: {preset명: flag 목록} + 직접 입력
WORK_PRESETS: dict[str, list[str]] = {
    name: flags for name, flags in WORK_PRESET_ITEMS
}
WORK_PRESETS["직접 입력"] = []

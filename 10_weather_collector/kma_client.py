import os
import requests
import time
from datetime import datetime, timedelta, date
import config as _config
from flags import FLAG_COMPUTATIONS

BASE_URL = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"


def _get_api_key() -> str:
    """항상 최신 API 키를 반환한다."""
    return _config.KMA_API_KEY or os.getenv("KMA_API_KEY", "")


def validate_station(station_code: str) -> bool:
    """
    관측소 코드가 ASOS API에서 유효한지 테스트 요청으로 확인.
    최근 7일치 데이터를 요청해 'body' 키가 존재하면 유효한 관측소.
    """
    test_end = date.today() - timedelta(days=30)
    test_start = test_end - timedelta(days=6)

    params = {
        "serviceKey": _get_api_key(),
        "numOfRows":  7,
        "pageNo":     1,
        "dataType":   "JSON",
        "dataCd":     "ASOS",
        "dateCd":     "DAY",
        "startDt":    test_start.strftime("%Y%m%d"),
        "endDt":      test_end.strftime("%Y%m%d"),
        "stnIds":     station_code,
    }

    try:
        resp = requests.get(BASE_URL, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        result = data.get("response", {})
        if "body" not in result:
            print(f"[validate] {station_code} — body 없음: {result.get('header', {})}")
            return False
        total = result["body"].get("totalCount", 0)
        return int(total) > 0
    except Exception as e:
        print(f"[validate] {station_code} — 오류: {e}")
        return False


def fetch_daily_weather(station_code: str, start_date: str, end_date: str) -> list[dict]:
    """
    ASOS 일자료 조회 (최대 365일씩 분할 요청)
    start_date, end_date: "YYYYMMDD" 형식
    """
    results = []
    start = datetime.strptime(start_date, "%Y%m%d")
    end   = datetime.strptime(end_date,   "%Y%m%d")

    chunk_days = 365
    current = start

    while current <= end:
        chunk_end = min(current + timedelta(days=chunk_days - 1), end)

        params = {
            "serviceKey": _get_api_key(),
            "numOfRows":  chunk_days,
            "pageNo":     1,
            "dataType":   "JSON",
            "dataCd":     "ASOS",
            "dateCd":     "DAY",
            "startDt":    current.strftime("%Y%m%d"),
            "endDt":      chunk_end.strftime("%Y%m%d"),
            "stnIds":     station_code,
        }

        try:
            resp = requests.get(BASE_URL, params=params, timeout=15)
            resp.raise_for_status()
            data = resp.json()
            items = data["response"]["body"]["items"]["item"]
            if isinstance(items, list):
                results.extend(items)
            else:
                results.append(items)
        except KeyError as e:
            print(f"[ERROR] 응답 파싱 실패 {station_code} "
                  f"{current.date()} ~ {chunk_end.date()}: 키 없음 {e}")
            print(f"  응답 내용: {data}")
        except Exception as e:
            print(f"[ERROR] {station_code} {current.date()} ~ {chunk_end.date()}: {e}")

        time.sleep(0.5)
        current = chunk_end + timedelta(days=1)

    return results


def parse_weather(raw: dict, site_id: str) -> dict:
    """API 응답 → 건설 관리 필요 항목 추출"""

    def f(key):
        """빈 문자열·None → None 반환 (결측값과 0을 구분)"""
        v = raw.get(key, "")
        if v is None or v == "":
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    def cmp(val, op, threshold):
        """결측값(None)은 작업불가일 미해당(False)으로 처리"""
        if val is None:
            return False
        if op == ">=": return val >= threshold
        if op == "<=": return val <= threshold
        if op == "<":  return val < threshold
        if op == ">":  return val > threshold
        return False

    # ── 원시 관측값 추출 ────────────────────────────────────────
    # DB 컬럼명 → 파싱된 float 값 (결측 시 None)
    col_values: dict[str, float | None] = {
        "precipitation":  f("sumRn"),        # 일강수량 (mm)
        "wind_avg":       f("avgWs"),        # 평균풍속 (m/s)
        "wind_max":       f("maxWs"),        # 최대풍속 (m/s)
        "max_ins_wind":   f("maxInsWs"),     # 순간최대풍속 (m/s)
        "snow_depth":     f("sumDpthFhsc"),  # 최대적설 (cm)
        "temp_max":       f("maxTa"),        # 최고기온 (℃)
        "temp_min":       f("minTa"),        # 최저기온 (℃)
        "sunshine_hours": f("sumSsHr"),      # 일조시간 (hr)
        "ground_temp":    f("avgTs"),        # 지면온도 (℃)
        "evaporation":    f("sumSmlEv"),     # 증발량 (mm)
        "humidity_avg":   f("avgRhm"),       # 평균습도 (%)
        "pressure":       f("avgPa"),        # 평균기압 (hPa)
    }

    # ── 강수·강설·안개 유무: iscs 필드 문자열 파싱 ─────────────────
    iscs    = raw.get("iscs") or ""
    fog_dur = f("sumFogDur")             # 안개 지속시간 (hr)

    rain_yn = "{비}" in iscs or "{소나기}" in iscs
    snow_yn = "{눈}" in iscs or "{진눈깨비}" in iscs
    fog_yn  = (fog_dur or 0) > 0 or "{안개}" in iscs or "{박무}" in iscs

    # ── 수치 플래그: FLAG_COMPUTATIONS 기본값으로 일괄 계산 ─────────
    # 결측값(None)은 cmp()에서 False 처리 → 오판정 방지
    numeric_flags = {
        flag_id: cmp(col_values.get(col), op, threshold)
        for flag_id, (col, op, threshold) in FLAG_COMPUTATIONS.items()
    }

    return {
        # ── 현장·날짜 식별 ──────────────────────────────
        "site_id":       site_id,
        "date":          raw.get("tm"),
        "station_code":  raw.get("stnId"),

        # ── 기본 기상 관측값 ────────────────────────────
        "temp_max":       col_values["temp_max"],
        "temp_min":       col_values["temp_min"],
        "precipitation":  col_values["precipitation"],
        "wind_avg":       col_values["wind_avg"],
        "wind_max":       col_values["wind_max"],
        "max_ins_wind":   col_values["max_ins_wind"],
        "snow_depth":     col_values["snow_depth"],
        "humidity_avg":   col_values["humidity_avg"],
        "sunshine_hours": col_values["sunshine_hours"],
        "ground_temp":    col_values["ground_temp"],
        "evaporation":    col_values["evaporation"],
        "pressure":       col_values["pressure"],

        # ── 작업불가일 판정 플래그 (flags.py 기본값, 결측 시 False) ─
        **numeric_flags,
        "rain_yn": rain_yn,
        "snow_yn": snow_yn,
        "fog_yn":  fog_yn,
    }
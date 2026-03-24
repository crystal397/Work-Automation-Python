import requests
import time
from datetime import datetime, timedelta
from config import KMA_API_KEY

BASE_URL = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"

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
            "serviceKey": KMA_API_KEY,
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
        """빈 문자열 또는 None이면 0.0으로 변환"""
        v = raw.get(key, "")
        try:
            return float(v) if v != "" else 0.0
        except (ValueError, TypeError):
            return 0.0

    precipitation  = f("sumRn")        # 일강수량 (mm)
    max_wind       = f("maxWs")        # 최대풍속 (m/s)
    max_ins_wind   = f("maxInsWs")     # 순간최대풍속 (m/s)
    snow_depth     = f("sumDpthFhsc")  # 최대적설 (cm)
    temp_max       = f("maxTa")        # 최고기온 (℃)
    temp_min       = f("minTa")        # 최저기온 (℃)
    sunshine_hours = f("sumSsHr")      # 일조시간 (hr)
    ground_temp    = f("avgTs")        # 지면온도 (℃)
    evaporation    = f("sumSmlEv")     # 증발량 (mm) - 소형증발계
    pressure       = f("avgPa")        # 평균기압 (hPa)
    fog_dur        = f("sumFogDur")    # 안개 지속시간 (hr)

    # 강수·강설 유무: iscs 필드 문자열 파싱
    iscs    = raw.get("iscs", "")
    rain_yn = "{비}" in iscs or "{소나기}" in iscs
    snow_yn = "{눈}" in iscs or "{진눈깨비}" in iscs

    # 안개 유무: sumFogDur > 0 이거나 iscs에 안개·박무 포함
    fog_yn  = fog_dur > 0 or "{안개}" in iscs or "{박무}" in iscs

    return {
        # ── 현장·날짜 식별 ──────────────────────────────
        "site_id":       site_id,
        "date":          raw.get("tm"),
        "station_code":  raw.get("stnId"),

        # ── 기본 기상 관측값 ────────────────────────────
        "temp_max":        temp_max,
        "temp_min":        temp_min,
        "precipitation":   precipitation,
        "wind_avg":        f("avgWs"),
        "wind_max":        max_wind,
        "max_ins_wind":    max_ins_wind,
        "snow_depth":      snow_depth,
        "humidity_avg":    f("avgRhm"),
        "sunshine_hours":  sunshine_hours,
        "ground_temp":     ground_temp,
        "evaporation":     evaporation,
        "pressure":        pressure,

        # ── 작업불가일 판정 플래그 ──────────────────────
        "is_rain_day":      precipitation  >= 10,   # 강수 10mm 이상
        "is_wind_day":      max_wind       >= 14,   # 최대풍속 14m/s 이상
        "is_wind_crane":    max_ins_wind   >= 10,   # 크레인 작업 제한 (순간 10m/s)
        "is_snow_day":      snow_depth     >= 1,    # 적설 1cm 이상
        "is_heat_day":      temp_max       >= 35,   # 폭염 35℃ 이상
        "is_cold_day":      temp_min       <= -10,  # 한파 -10℃ 이하
        "is_no_sunshine":   sunshine_hours < 2,     # 일조 2시간 미만
        "is_freeze_day":    ground_temp    <= 0,    # 지면 동결
        "is_high_evap_day": evaporation    >= 10,   # 증발 과다
        "rain_yn":          rain_yn,                # 강수 유무 (iscs 기반)
        "snow_yn":          snow_yn,                # 강설 유무 (iscs 기반)
        "fog_yn":           fog_yn,                 # 안개 유무
    }
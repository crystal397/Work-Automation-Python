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
                results.append(items)  # 1건일 때 dict로 반환되는 경우
        except Exception as e:
            print(f"[ERROR] {station_code} {current.date()} ~ {chunk_end.date()}: {e}")

        time.sleep(0.5)  # API 호출 제한 준수 (초당 1회)
        current = chunk_end + timedelta(days=1)

    return results


def parse_weather(raw: dict, site_id: str) -> dict:
    """API 응답 → 건설 관리 필요 항목 추출"""

    precipitation  = float(raw.get("sumRn")      or 0)  # 일강수량 (mm)
    max_wind       = float(raw.get("maxWs")       or 0)  # 최대풍속 (m/s)
    max_ins_wind   = float(raw.get("maxInsWs")    or 0)  # 순간최대풍속 (m/s)
    snow_depth     = float(raw.get("maxDpth")     or 0)  # 최대적설 (cm)
    temp_max       = float(raw.get("maxTa")       or 0)  # 최고기온 (℃)
    temp_min       = float(raw.get("minTa")       or 0)  # 최저기온 (℃)
    sunshine_hours = float(raw.get("sumSsHr")     or 0)  # 일조시간 (hr)
    ground_temp    = float(raw.get("avgTs")       or 0)  # 지면온도 (℃)
    evaporation    = float(raw.get("sumEv")       or 0)  # 증발량 (mm)
    pressure       = float(raw.get("avgPa")       or 0)  # 평균기압 (hPa)

    return {
        # ── 현장·날짜 식별 ──────────────────────────────
        "site_id":        site_id,
        "date":           raw.get("tm"),
        "station_code":   raw.get("stnId"),

        # ── 기본 기상 관측값 ────────────────────────────
        "temp_max":        temp_max,
        "temp_min":        temp_min,
        "precipitation":   precipitation,
        "wind_avg":        float(raw.get("avgWs")  or 0),  # 평균풍속 (m/s)
        "wind_max":        max_wind,
        "max_ins_wind":    max_ins_wind,                    # 순간최대풍속 (m/s)
        "snow_depth":      snow_depth,
        "humidity_avg":    float(raw.get("avgRhm") or 0),  # 평균습도 (%)
        "sunshine_hours":  sunshine_hours,                  # 일조시간 (hr)
        "ground_temp":     ground_temp,                     # 지면온도 (℃) ★추가
        "evaporation":     evaporation,                     # 증발량 (mm) ★추가
        "pressure":        pressure,                        # 평균기압 (hPa) ★추가

        # ── 작업불가일 판정 플래그 ──────────────────────
        "is_rain_day":     precipitation  >= 10,   # 강수 10mm 이상
        "is_wind_day":     max_wind       >= 14,   # 최대풍속 14m/s 이상
        "is_wind_crane":   max_ins_wind   >= 10,   # 크레인 작업 제한 (순간 10m/s)
        "is_snow_day":     snow_depth     >= 1,    # 적설 1cm 이상
        "is_heat_day":     temp_max       >= 35,   # 폭염 35℃ 이상
        "is_cold_day":     temp_min       <= -10,  # 한파 -10℃ 이하
        "is_no_sunshine":  sunshine_hours < 2,     # 일조 2시간 미만
        "is_freeze_day":   ground_temp    <= 0,    # 지면 동결 (지면온도 0℃ 이하) ★추가
        "is_high_evap_day":evaporation    >= 10,   # 증발 과다 (10mm 이상) ★추가
        "rain_yn":         raw.get("rnDay")  == "1",  # 강수 유무 (소량 포함)
        "fog_yn":          raw.get("fogDay") == "1",  # 안개 유무
    }
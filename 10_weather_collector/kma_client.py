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
    
    # API는 한 번에 최대 365일 → 청크로 분할
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
            results.extend(items)
        except Exception as e:
            print(f"[ERROR] {station_code} {current.date()} ~ {chunk_end.date()}: {e}")
        
        time.sleep(0.5)  # API 호출 제한 준수 (초당 1회)
        current = chunk_end + timedelta(days=1)
    
    return results

def parse_weather(raw: dict, site_id: str) -> dict:
    """API 응답 → 건설 관리 필요 항목 추출"""
    return {
        "site_id":      site_id,
        "date":         raw.get("tm"),
        "station_code": raw.get("stnId"),
        "temp_max":     float(raw.get("maxTa") or 0),    # 최고기온(℃)
        "temp_min":     float(raw.get("minTa") or 0),    # 최저기온(℃)
        "precipitation":float(raw.get("sumRn") or 0),    # 일강수량(mm)
        "wind_avg":     float(raw.get("avgWs") or 0),    # 평균풍속(m/s)
        "wind_max":     float(raw.get("maxWs") or 0),    # 최대풍속(m/s)
        "snow_depth":   float(raw.get("maxDpth") or 0),  # 최대적설(cm)
        "humidity_avg": float(raw.get("avgRhm") or 0),   # 평균습도(%)
        # 작업불가일 판정용 플래그 자동 계산
        "is_rain_day":  float(raw.get("sumRn") or 0) >= 10,   # 강수 10mm 이상
        "is_wind_day":  float(raw.get("maxWs") or 0) >= 14,   # 최대풍속 14m/s 이상
        "is_snow_day":  float(raw.get("maxDpth") or 0) >= 1,  # 적설 1cm 이상
        "is_heat_day":  float(raw.get("maxTa") or 0) >= 35,   # 폭염 35℃ 이상
        "is_cold_day":  float(raw.get("minTa") or 0) <= -10,  # 한파 -10℃ 이하
    }
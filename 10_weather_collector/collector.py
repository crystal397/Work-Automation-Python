from config import SITES
from station_mapper import ASOS_STATIONS, find_nearest_station, haversine
from kma_client import fetch_daily_weather, parse_weather, validate_station
from storage import init_db, upsert_weather


def _find_valid_station(lat: float, lon: float) -> dict | None:
    """
    좌표 기준 가장 가까운 관측소부터 순서대로 ASOS 유효성 검증.
    유효한 관측소를 찾으면 반환, 20개 내에서 찾지 못하면 None.
    """
    candidates = sorted(
        ASOS_STATIONS,
        key=lambda s: haversine(lat, lon, s["lat"], s["lon"])
    )[:20]

    for station in candidates:
        if validate_station(station["code"]):
            return station
    return None


def run_bulk_collection(start_date: str = None, end_date: str = None):
    """
    sites.json의 모든 현장 데이터 일괄 수집.
    start_date / end_date 를 지정하면 해당 기간만 수집 (YYYYMMDD 형식).
    생략 시 각 현장의 설정 기간 전체를 수집.
    """
    init_db()
    for site in SITES:
        station = _find_valid_station(site["lat"], site["lon"])
        if station is None:
            print(f"[{site['name']}] ✗ ASOS 유효 관측소를 찾을 수 없습니다. 건너뜁니다.")
            continue

        print(f"[{site['name']}] → 관측소: {station['name']}({station['code']})")

        s = start_date or site["start"].replace("-", "")
        e = end_date   or site["end"].replace("-", "")

        raw_records = fetch_daily_weather(station["code"], s, e)
        parsed = [parse_weather(r, site["id"]) for r in raw_records]
        upsert_weather(parsed)
        print(f"  → {len(parsed)}일치 수집 완료")


if __name__ == "__main__":
    run_bulk_collection()

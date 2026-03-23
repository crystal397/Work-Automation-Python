from config import SITES
from station_mapper import find_nearest_station
from kma_client import fetch_daily_weather, parse_weather
from storage import init_db, upsert_weather

def run_bulk_collection():
    init_db()
    for site in SITES:
        station = find_nearest_station(site["lat"], site["lon"])
        print(f"[{site['name']}] → 관측소: {station['name']}({station['code']})")
        
        raw_records = fetch_daily_weather(station["code"], site["start"].replace("-",""), site["end"].replace("-",""))
        parsed = [parse_weather(r, site["id"]) for r in raw_records]
        upsert_weather(parsed)
        
        print(f"  → {len(parsed)}일치 수집 완료")

if __name__ == "__main__":
    run_bulk_collection()
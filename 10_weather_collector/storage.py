from sqlalchemy import create_engine, text
from config import DB_PATH

engine = create_engine(f"sqlite:///{DB_PATH}")

def init_db():
    with engine.connect() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS weather_daily (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                site_id       TEXT NOT NULL,
                date          TEXT NOT NULL,
                station_code  TEXT,
                temp_max      REAL, temp_min      REAL,
                precipitation REAL, wind_avg      REAL,
                wind_max      REAL, snow_depth    REAL,
                humidity_avg  REAL,
                is_rain_day   BOOLEAN, is_wind_day BOOLEAN,
                is_snow_day   BOOLEAN, is_heat_day BOOLEAN,
                is_cold_day   BOOLEAN,
                UNIQUE(site_id, date)
            )
        """))
        conn.commit()

def upsert_weather(records: list[dict]):
    if not records:
        return
    with engine.connect() as conn:
        for r in records:
            conn.execute(text("""
                INSERT INTO weather_daily (
                    site_id, date, station_code,
                    temp_max, temp_min, precipitation,
                    wind_avg, wind_max, snow_depth, humidity_avg,
                    is_rain_day, is_wind_day, is_snow_day, is_heat_day, is_cold_day
                ) VALUES (
                    :site_id, :date, :station_code,
                    :temp_max, :temp_min, :precipitation,
                    :wind_avg, :wind_max, :snow_depth, :humidity_avg,
                    :is_rain_day, :is_wind_day, :is_snow_day, :is_heat_day, :is_cold_day
                )
                ON CONFLICT(site_id, date) DO UPDATE SET
                    precipitation = excluded.precipitation,
                    wind_max      = excluded.wind_max
            """), r)
        conn.commit()
    print(f"[DB] {len(records)}건 저장 완료")
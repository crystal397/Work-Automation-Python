from sqlalchemy import create_engine, text
from config import DB_PATH

engine = create_engine(f"sqlite:///{DB_PATH}")

def init_db():
    with engine.connect() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS weather_daily (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                site_id         TEXT NOT NULL,
                date            TEXT NOT NULL,
                station_code    TEXT,
                temp_max        REAL,
                temp_min        REAL,
                precipitation   REAL,
                wind_avg        REAL,
                wind_max        REAL,
                max_ins_wind    REAL,
                snow_depth      REAL,
                humidity_avg    REAL,
                sunshine_hours  REAL,
                ground_temp     REAL,               -- 지면온도 ★추가
                evaporation     REAL,               -- 증발량 ★추가
                pressure        REAL,               -- 평균기압 ★추가
                is_rain_day     BOOLEAN,
                is_wind_day     BOOLEAN,
                is_wind_crane   BOOLEAN,
                is_snow_day     BOOLEAN,
                is_heat_day     BOOLEAN,
                is_cold_day     BOOLEAN,
                is_no_sunshine  BOOLEAN,
                is_freeze_day   BOOLEAN,            -- 지면 동결 ★추가
                is_high_evap_day BOOLEAN,           -- 증발 과다 ★추가
                rain_yn         BOOLEAN,
                fog_yn          BOOLEAN,
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
                    wind_avg, wind_max, max_ins_wind,
                    snow_depth, humidity_avg, sunshine_hours,
                    ground_temp, evaporation, pressure,
                    is_rain_day, is_wind_day, is_wind_crane,
                    is_snow_day, is_heat_day, is_cold_day,
                    is_no_sunshine, is_freeze_day, is_high_evap_day,
                    rain_yn, fog_yn
                ) VALUES (
                    :site_id, :date, :station_code,
                    :temp_max, :temp_min, :precipitation,
                    :wind_avg, :wind_max, :max_ins_wind,
                    :snow_depth, :humidity_avg, :sunshine_hours,
                    :ground_temp, :evaporation, :pressure,
                    :is_rain_day, :is_wind_day, :is_wind_crane,
                    :is_snow_day, :is_heat_day, :is_cold_day,
                    :is_no_sunshine, :is_freeze_day, :is_high_evap_day,
                    :rain_yn, :fog_yn
                )
                ON CONFLICT(site_id, date) DO UPDATE SET
                    temp_max         = excluded.temp_max,
                    temp_min         = excluded.temp_min,
                    precipitation    = excluded.precipitation,
                    wind_avg         = excluded.wind_avg,
                    wind_max         = excluded.wind_max,
                    max_ins_wind     = excluded.max_ins_wind,
                    snow_depth       = excluded.snow_depth,
                    humidity_avg     = excluded.humidity_avg,
                    sunshine_hours   = excluded.sunshine_hours,
                    ground_temp      = excluded.ground_temp,
                    evaporation      = excluded.evaporation,
                    pressure         = excluded.pressure,
                    is_rain_day      = excluded.is_rain_day,
                    is_wind_day      = excluded.is_wind_day,
                    is_wind_crane    = excluded.is_wind_crane,
                    is_snow_day      = excluded.is_snow_day,
                    is_heat_day      = excluded.is_heat_day,
                    is_cold_day      = excluded.is_cold_day,
                    is_no_sunshine   = excluded.is_no_sunshine,
                    is_freeze_day    = excluded.is_freeze_day,
                    is_high_evap_day = excluded.is_high_evap_day,
                    rain_yn          = excluded.rain_yn,
                    fog_yn           = excluded.fog_yn
            """), r)
        conn.commit()
    print(f"[DB] {len(records)}건 저장 완료")
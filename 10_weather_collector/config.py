import os
import sys
from pathlib import Path
from datetime import date
from dotenv import load_dotenv

# exe(frozen)로 실행될 때는 exe 파일이 있는 폴더 기준, 스크립트 실행 시는 상위 폴더 기준
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
    load_dotenv(dotenv_path=BASE_DIR.parent / ".env")  # 개발 환경: 상위 폴더의 .env

load_dotenv(dotenv_path=BASE_DIR / ".env")  # exe 환경: exe 옆의 .env (덮어쓰기)

KMA_API_KEY = os.getenv("KMA_API_KEY")


def save_api_key(key: str) -> None:
    """API 키를 .env 파일에 저장하고 현재 프로세스에도 반영한다."""
    env_path = BASE_DIR / ".env"
    env_path.write_text(f"KMA_API_KEY={key.strip()}\n", encoding="utf-8")
    os.environ["KMA_API_KEY"] = key.strip()
    # 모듈 수준 변수도 갱신
    import config as _self
    _self.KMA_API_KEY = key.strip()

DB_PATH = BASE_DIR / "weather.db"

# 현장 목록 (위도/경도 기반)
SITES = [
    {
        "id":           "SITE001",
        "name":         "수원 장안구 파장동~송죽동 현장",
        "lat":          37.2723,
        "lon":          126.9853,
        "start":        "2024-01-01",
        "end":          "2025-12-31",

        # 공종별 작업 기간 정의
        "works": [
            {
                "name":  "토공사",
                "start": "2024-01-01",
                "end":   "2024-03-31",
                "flags": ["is_rain_day", "is_snow_day", "is_freeze_day",
                          "is_cold_day"]
            },
            {
                "name":  "철근콘크리트공사",
                "start": "2024-04-01",
                "end":   "2024-09-30",
                "flags": ["is_rain_day", "is_heat_day", "is_cold_day",
                          "is_freeze_day", "is_wind_day"]
            },
            {
                "name":  "타워크레인작업",
                "start": "2024-04-01",
                "end":   "2024-12-31",
                "flags": ["is_wind_crane", "is_wind_day", "fog_yn"]
            },
            {
                "name":  "도장·방수공사",
                "start": "2025-01-01",
                "end":   "2025-06-30",
                "flags": ["is_rain_day", "rain_yn", "is_no_sunshine",
                          "is_cold_day", "is_freeze_day"]
            },
        ]
    },
]
import os
from pathlib import Path
from dotenv import load_dotenv

# 상위 폴더의 .env 로드
load_dotenv(dotenv_path=Path(__file__).parent.parent / ".env")

KMA_API_KEY = os.getenv("KMA_API_KEY")
DB_URL = os.getenv("DB_URL", "sqlite:///weather.db")

# 현장 목록 (위도/경도 기반)
SITES = [
    {"id": "SITE001", "name": "서울 강남 현장", "lat": 37.5172, "lon": 127.0473, "start": "2024-01-01", "end": "2025-06-30"},
    {"id": "SITE002", "name": "부산 해운대 현장", "lat": 35.1631, "lon": 129.1639, "start": "2024-03-01", "end": "2025-12-31"},
]
import json
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

# 현장 목록 — sites.json 에서 로드 (git 비공개)
# 설정 방법: sites.example.json 을 복사해 sites.json 으로 저장 후 실제 값 입력
_sites_file = BASE_DIR / "sites.json"
if _sites_file.exists():
    SITES = json.loads(_sites_file.read_text(encoding="utf-8"))
else:
    SITES = []
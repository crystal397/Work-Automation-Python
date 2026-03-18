"""
LH 임대주택 도로명주소 기반
국토교통부 실거래가 API 조회 스크립트 v4.0
=====================================================
사전 준비:
  pip install pandas openpyxl requests tqdm

사용법:
  python lh_realestate_api.py        # 실거래가 조회 (행안부+Kakao 지번변환 + 매칭 통합 실행)

[v4.0 수정 목록]
  ㉟  clean_address_for_jibun() 추가: 지번변환 전 주소 정제
      괄호 뒤 부가텍스트(외N필지·건물명) 제거 → not_found 감소
  ㉞  extract_dong_from_address 개선: 동명+건물명 붙은 주소 동명 분리
      (덕진동2가영무예다음 → 덕진동2가 / 효자동효자코아루 → 효자동 등)
      여러 괄호 마지막 우선, 지 접미어 오탐 방지, 영문 건물명 분리
  ㉓  Semaphore가 _increment_call 감싸도록 이동 (카운트 선행 방지)
  ㉔  match_trade_to_address 내부 정렬 제거, run()에서 1회만 정렬
  ㉕  exhausted 시 threading.Event + future.cancel()로 배치 즉시 중단
  ㉖  trade_cache → SQLite(TradeCacheDB)로 교체, 메모리 O(1그룹) 유지
  ㉗  SQLite 캐시에 fetched_at + TTL(30일) 적용
  ㉘  _SORTED_HOUSING_KEYS 모듈 상수로 캐싱 (sorted 반복 호출 제거)
  ㉙  jibun_cache에 _meta 버전 태그 → 이미 변환된 캐시 migration 스킵
  ㉚  API 키를 .env 파일에서 로드 (소스코드 하드코딩 제거)
  ㉛  fetch_one 모듈 레벨 함수로 분리 (테스트·재사용 가능)
  ㉜  Kakao 429 지수 백오프 재시도 (QPS 제한 대응)
  ㉝  매칭 실패 케이스를 no_match_debug.xlsx로 저장

[v3.0 수정 목록]
  ①  trade_cache 영구 저장 → ㉖에서 SQLite로 발전
  ②  지역 우선순위 배치 처리
  ③  jibun_cache 구버전 호환 (migrate_jibun_cache)
  ④  DAILY_LIMIT MAX_WORKERS 안전 마진
  ⑥  exhausted 결과 cache 미저장
  ⑦  부산 서구/동구/남구/북구 접두사 키로 교체
  ⑧  페이지네이션 page_no attempt마다 리셋
  ⑨  make_dedup_key() 단일 함수
  ⑩  API별 Semaphore
  ⑪  api_url 컬럼 drop
  ⑫  apt Grade3 과잉 매칭 수정
  ⑬  을지로2가 등 숫자+가 파싱 수정
  ⑭  future.result() 예외 처리
  ⑮  get_api_info 완전일치 우선 + 긴 키 우선
  ⑯  Kakao similar 재시도 지연 추가
  ⑰  jibun 공백 구분자 처리
  ⑱  전용면적 NaN dedup_key 일관성
  ⑲  정렬 중복 제거
  ⑳  세종특별자치시 코드 누락
  ㉑  괄호 없는 주소 법정동 토큰 폴백

[v2.0 수정 목록]
  - Kakao 다필지 최대 3후보 반환
  - jibun_cache status(ok/not_found/network_error) 구분
  - grade_match_non_apt jibun_exact/jibun_bonbun 2단계
  - match_trade_to_address 후보 순회 최선 등급 채택
"""

import re
import time
import json
import logging
import sqlite3
import threading
import concurrent.futures
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timedelta
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

# ──────────────────────────────────────────────
# ① .env 로드  [㉚]
# ──────────────────────────────────────────────
def _load_env(path: str = ".env") -> dict:
    """
    .env 파일을 수동 파싱. python-dotenv 미설치 환경에서도 동작.
    처리 규칙:
      - 줄 단위 # 주석·빈 줄 무시
      - 등호 앞뒤 공백 제거
      - 값의 앞뒤 따옴표(작은/큰) 제거
      - 인라인 주석(공백+# ...) 제거  ← 추가
        예) KEY = "abc"   # 설명  →  abc
    """
    env: dict[str, str] = {}
    if not Path(path).exists():
        return env
    with open(path, encoding="utf-8-sig") as f:   # utf-8-sig: BOM 자동 제거
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            v = v.strip()
            # 따옴표로 감싸진 값: 닫는 따옴표 위치까지만 추출
            if v and v[0] in ('"', "'") and v[0] in v[1:]:
                v = v[1:v.index(v[0], 1)]
            else:
                # 따옴표 없으면 공백+# 이후 인라인 주석 제거 (URL 안 # 는 보호)
                v = re.sub(r'\s+#.*$', '', v).strip()
            env[k.strip()] = v
    return env

_ENV = _load_env()

# ──────────────────────────────────────────────
# ② 설정
# ──────────────────────────────────────────────
CONFIG = {
    # [㉚] 키는 .env에서 로드, 없으면 빈 문자열
    "SERVICE_KEY":       _ENV.get("MOLIT_SERVICE_KEY",   ""),
    "KAKAO_API_KEY":     _ENV.get("KAKAO_API_KEY",        ""),   # 폴백용 유지
    "JUSO_API_KEY":     _ENV.get("JUSO_API_KEY",         ""),   # ② 행안부 도로명주소 API
    "START_YMD":         "201603",   # ③ 확대: 2016.03부터 (10년, 기존 202001)
    "END_YMD":           "202602",
    "INPUT_FILE":        "LH_임대주택공급현황_251120.xlsx",
    "OUTPUT_DIR":        "output",
    # [㉖] SQLite 경로
    "TRADE_DB_FILE":     "output/trade_cache.db",
    "JIBUN_CACHE_FILE":  "output/jibun_cache.json",
    "CALL_INTERVAL":     0.5,
    "KAKAO_INTERVAL":    0.05,
    "MAX_RETRY":         3,
    "MAX_WORKERS":       3,
    # [④] DAILY_LIMIT: MAX_WORKERS만큼 안전 마진
    "DAILY_LIMIT":       9_500 - 3,
    # [㉗] 캐시 TTL (일)
    "CACHE_TTL_DAYS":    30,
}

# ──────────────────────────────────────────────
# ③ 지역 우선순위
# ──────────────────────────────────────────────
REGION_PRIORITY: dict[str, int] = {
    "서울특별시": 1, "경기도": 2, "인천광역시": 2,
    "부산광역시": 3, "대구광역시": 3, "광주광역시": 3,
    "대전광역시": 3, "울산광역시": 3, "세종특별자치시": 3,
}


def get_region_priority(sido: str) -> int:
    return REGION_PRIORITY.get(str(sido).strip(), 4)


# ──────────────────────────────────────────────
# ④ 법정동코드 매핑
#    [⑦] 부산 generic 키(서구/동구/남구/북구) → 접두사 키로 교체
#    [⑳] 세종 복수 키 추가
# ──────────────────────────────────────────────
SIGUNGU_CODE: dict[str, str] = {
    # 서울
    "종로구":"11110","중구":"11140","용산구":"11170","성동구":"11200",
    "광진구":"11215","동대문구":"11230","중랑구":"11260","성북구":"11290",
    "강북구":"11305","도봉구":"11320","노원구":"11350","은평구":"11380",
    "서대문구":"11410","마포구":"11440","양천구":"11470","강서구":"11500",
    "구로구":"11530","금천구":"11545","영등포구":"11560","동작구":"11590",
    "관악구":"11620","서초구":"11650","강남구":"11680","송파구":"11710","강동구":"11740",
    # 부산 [⑦] 접두사 키로 교체
    "부산중구":"26110","부산서구":"26140","부산동구":"26170","영도구":"26200",
    "부산진구":"26230","동래구":"26260","부산남구":"26290","부산북구":"26320",
    "해운대구":"26350","사하구":"26380","금정구":"26410","강서구_부산":"26440",
    "연제구":"26470","수영구":"26500","사상구":"26530","기장군":"26710",
    # 대구
    "대구중구":"27110","대구동구":"27140","대구서구":"27170","대구남구":"27200",
    "대구북구":"27230","수성구":"27260","달서구":"27290","달성군":"27710",
    # 인천
    "인천중구":"28110","인천동구":"28140","미추홀구":"28177","연수구":"28185",
    "남동구":"28200","부평구":"28237","계양구":"28245","인천서구":"28260",
    "강화군":"28710","옹진군":"28720",
    # 광주
    "광주동구":"29110","광주서구":"29140","광주남구":"29155","광주북구":"29170","광산구":"29200",
    # 대전
    "대전동구":"30110","대전중구":"30140","대전서구":"30170","대전유성구":"30200","대전대덕구":"30230",
    # 울산
    "울산중구":"31110","울산남구":"31140","울산동구":"31170","울산북구":"31200","울주군":"31710",
    # 세종 [⑳] 복수 키
    "세종시":"36110","세종특별자치시":"36110","세종":"36110",
    # 경기
    "수원시장안구":"41111","수원시권선구":"41113","수원시팔달구":"41115","수원시영통구":"41117",
    "성남시수정구":"41131","성남시중원구":"41133","성남시분당구":"41135",
    "의정부시":"41150","안양시만안구":"41171","안양시동안구":"41173",
    "부천시":"41190","광명시":"41210","평택시":"41220","동두천시":"41250",
    "안산시상록구":"41271","안산시단원구":"41273","고양시덕양구":"41281",
    "고양시일산동구":"41285","고양시일산서구":"41287","과천시":"41290",
    "구리시":"41310","남양주시":"41360","오산시":"41370","시흥시":"41390",
    "군포시":"41410","의왕시":"41430","하남시":"41450","용인시처인구":"41461",
    "용인시기흥구":"41463","용인시수지구":"41465","파주시":"41480",
    "이천시":"41500","안성시":"41550","김포시":"41570","화성시":"41590",
    "광주시":"41610","양주시":"41630","포천시":"41650","여주시":"41670",
    "연천군":"41800","가평군":"41820","양평군":"41830",
    # 강원
    "춘천시":"42110","원주시":"42130","강릉시":"42150","동해시":"42170",
    "태백시":"42190","속초시":"42210","삼척시":"42230","홍천군":"42720",
    "횡성군":"42730","영월군":"42750","평창군":"42760","정선군":"42770",
    "철원군":"42780","화천군":"42790","양구군":"42800","인제군":"42810",
    "고성군_강원":"42820","양양군":"42830",
    # 충북
    "청주시상당구":"43111","청주시서원구":"43112","청주시흥덕구":"43113","청주시청원구":"43114",
    "충주시":"43130","제천시":"43150","보은군":"43720","옥천군":"43730","영동군":"43740",
    "증평군":"43745","진천군":"43750","괴산군":"43760","음성군":"43770","단양군":"43800",
    # 충남
    "천안시동남구":"44131","천안시서북구":"44133","공주시":"44150","보령시":"44180",
    "아산시":"44200","서산시":"44210","논산시":"44230","계룡시":"44250","당진시":"44270",
    "금산군":"44710","부여군":"44760","서천군":"44770","청양군":"44790","홍성군":"44800",
    "예산군":"44810","태안군":"44825",
    # 전북
    "전주시완산구":"45111","전주시덕진구":"45113","군산시":"45130","익산시":"45140",
    "정읍시":"45180","남원시":"45190","김제시":"45210","완주군":"45710","진안군":"45720",
    "무주군":"45730","장수군":"45740","임실군":"45750","순창군":"45770","고창군":"45790","부안군":"45800",
    # 전남
    "목포시":"46110","여수시":"46130","순천시":"46150","나주시":"46170","광양시":"46230",
    "담양군":"46710","곡성군":"46720","구례군":"46730","고흥군":"46770","보성군":"46780",
    "화순군":"46790","장흥군":"46800","강진군":"46810","해남군":"46820","영암군":"46830",
    "무안군":"46840","함평군":"46860","영광군":"46870","장성군":"46880","완도군":"46890",
    "진도군":"46900","신안군":"46910",
    # 경북
    "포항시남구":"47111","포항시북구":"47113","경주시":"47130","김천시":"47150","안동시":"47170",
    "구미시":"47190","영주시":"47210","영천시":"47230","상주시":"47250","문경시":"47280",
    "경산시":"47290","군위군":"47720","의성군":"47730","청송군":"47750","영양군":"47760",
    "영덕군":"47770","청도군":"47820","고령군":"47830","성주군":"47840","칠곡군":"47850",
    "예천군":"47900","봉화군":"47920","울진군":"47930","울릉군":"47940",
    # 경남
    "창원시의창구":"48121","창원시성산구":"48123","창원시마산합포구":"48125",
    "창원시마산회원구":"48127","창원시진해구":"48129","진주시":"48170",
    "통영시":"48220","사천시":"48240","김해시":"48250","밀양시":"48270","거제시":"48310",
    "양산시":"48330","의령군":"48720","함안군":"48730","창녕군":"48740","고성군_경남":"48820",
    "남해군":"48840","하동군":"48850","산청군":"48860","함양군":"48870","거창군":"48880","합천군":"48890",
    # 제주
    "제주시":"50110","서귀포시":"50130",
}

# ──────────────────────────────────────────────
# ⑤ 주택유형 → API 매핑
#    [㉘] _SORTED_HOUSING_KEYS: 모듈 로드 시 1회만 정렬
# ──────────────────────────────────────────────
HOUSING_TYPE_API: dict[str, tuple[str, bool]] = {
    "아파트":                           ("https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev",  True),
    "오피스텔+아파트":                  ("https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev",  True),
    "주상복합":                         ("https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev",  True),
    "연립":                             ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "다세대":                           ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형생활주택(원룸형)":           ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형생활주택(투룸형)":           ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형생활주택(단지형 다세대주택)":("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형생활주택(단지형 연립주택)":  ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형생활주택(소형 주택)":        ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형 생활주택(아파트형 주택)":   ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "도시형생활주택+오피스텔":          ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "다세대주택+오피스텔":              ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "혼합형":                           ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",          False),
    "오피스텔":                         ("https://apis.data.go.kr/1613000/RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade",      False),
    "단독":                             ("https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",          False),
    "다가구":                           ("https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",          False),
    "다중주택":                         ("https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",          False),
    "기숙사":                           ("https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",          False),
    "다가구주택+오피스텔":              ("https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",          False),
}
DEFAULT_API = ("https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade", False)

# [㉘] 모듈 로드 시 1회만 정렬 → get_api_info 내 반복 sorted() 제거
_SORTED_HOUSING_KEYS: list[str] = sorted(HOUSING_TYPE_API, key=len, reverse=True)

API_META: dict[str, dict] = {
    "getRTMSDataSvcAptTradeDev": {"has_road": True,  "has_jibun": False, "has_floor": True,  "has_area": True},
    "getRTMSDataSvcRHTrade":     {"has_road": False, "has_jibun": True,  "has_floor": True,  "has_area": True},
    "getRTMSDataSvcOffiTrade":   {"has_road": False, "has_jibun": True,  "has_floor": True,  "has_area": True},
    "getRTMSDataSvcSHTrade":     {"has_road": False, "has_jibun": True,  "has_floor": False, "has_area": False},
}


# ──────────────────────────────────────────────
# ⑥ 로거
# ──────────────────────────────────────────────
Path("output").mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("output/api_query.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# ⑦ 국토부 API 일일 한도 관리
# ──────────────────────────────────────────────
_api_call_counts: dict[str, int] = {}
_api_exhausted:   set[str]       = set()
_count_lock = threading.Lock()


def _increment_call(api_url: str) -> bool:
    """카운트 증가 + 한도 초과 시 exhausted 등록. 호출 가능 여부 반환."""
    with _count_lock:
        _api_call_counts.setdefault(api_url, 0)
        if api_url in _api_exhausted:
            return False
        _api_call_counts[api_url] += 1
        if _api_call_counts[api_url] >= CONFIG["DAILY_LIMIT"]:
            _api_exhausted.add(api_url)
            logger.warning(
                f"[한도] {api_url.split('/')[-1]} → {_api_call_counts[api_url]:,}건 도달. 당일 중단."
            )
        return True


class APIExhaustedException(Exception):
    pass


# ──────────────────────────────────────────────
# ⑧ API별 Semaphore [⑩]
#    [㉓] Semaphore가 _increment_call을 감싸야
#         카운트 증가와 실제 요청이 원자적으로 진행됨
# ──────────────────────────────────────────────
_api_semaphores: dict[str, threading.Semaphore] = {}
_semaphore_lock = threading.Lock()


def get_semaphore(api_url: str) -> threading.Semaphore:
    with _semaphore_lock:
        if api_url not in _api_semaphores:
            _api_semaphores[api_url] = threading.Semaphore(1)
        return _api_semaphores[api_url]


# ──────────────────────────────────────────────
# ⑨ 배치 중단 이벤트 [㉕]
# ──────────────────────────────────────────────
_stop_event = threading.Event()


# ──────────────────────────────────────────────
# ⑩ TradeCacheDB — SQLite 기반 영구 캐시 [㉖ ㉗]
# ──────────────────────────────────────────────
class TradeCacheDB:
    """
    SQLite 기반 실거래 API 응답 캐시.
    [㉖] pickle dict 대비 메모리 사용량 대폭 절감
         → 매칭 시 (lawd_cd, api_url) 그룹 단위 로드
    [㉗] fetched_at 컬럼 + TTL로 오래된 캐시 자동 만료
    [최적화①] get_all_valid_keys(): 전체 유효 키를 1회 SQL로 로드
              → combos 구성 시 275만 번 개별 SQL 제거
    [최적화②] _cutoff를 run() 시작 시 1회 고정
              → datetime.now() 반복 호출 제거
    """

    def __init__(self, db_path: str, ttl_days: int = 30):
        self.db_path   = db_path
        self.ttl_days  = ttl_days
        self._conn     = sqlite3.connect(db_path, check_same_thread=False)
        self._conn.execute("PRAGMA journal_mode=WAL")  # 동시 읽기/쓰기 안전
        self._pending  = 0
        # [최적화②] cutoff를 인스턴스 생성 시 1회 계산해 고정
        #           TTL 30일 대비 실행 수 시간 오차는 실용적으로 무의미
        self._cutoff_str: str = (
            datetime.now() - timedelta(days=ttl_days)
        ).isoformat()
        self._init_db()

    def _cutoff(self) -> str:
        """고정된 cutoff 반환 (run() 시작 시점 기준)."""
        return self._cutoff_str

    def _init_db(self):
        self._conn.execute("""
            CREATE TABLE IF NOT EXISTS trade_cache (
                lawd_cd    TEXT NOT NULL,
                ym         TEXT NOT NULL,
                api_url    TEXT NOT NULL,
                items_json TEXT NOT NULL,
                fetched_at TEXT NOT NULL,
                PRIMARY KEY (lawd_cd, ym, api_url)
            )
        """)
        self._conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_lawd_api "
            "ON trade_cache (lawd_cd, api_url)"
        )
        self._conn.commit()

    def get_all_valid_keys(self) -> set[tuple]:
        """
        [최적화①] 유효한 (lawd_cd, ym, api_url) 전체를 1회 SQL로 로드.
        combos 구성 시 개별 exists() 대신 메모리 set O(1) 조회로 교체.
        """
        rows = self._conn.execute(
            "SELECT lawd_cd, ym, api_url FROM trade_cache WHERE fetched_at>?",
            (self._cutoff(),),
        ).fetchall()
        return {(r[0], r[1], r[2]) for r in rows}

    def get(self, key: tuple) -> "list | None":
        lawd_cd, ym, api_url = key
        row = self._conn.execute(
            "SELECT items_json FROM trade_cache "
            "WHERE lawd_cd=? AND ym=? AND api_url=? AND fetched_at>?",
            (lawd_cd, ym, api_url, self._cutoff()),
        ).fetchone()
        return json.loads(row[0]) if row else None

    def get_group(self, lawd_cd: str, api_url: str, ym_list: list[str]) -> dict[str, list]:
        """
        (lawd_cd, api_url)의 모든 ym을 단일 SQL로 조회.
        [㉖] 매칭 시 그룹 단위 일괄 로드 → 반복 쿼리 제거.
        """
        if not ym_list:
            return {}
        ph   = ",".join("?" * len(ym_list))
        rows = self._conn.execute(
            f"SELECT ym, items_json FROM trade_cache "
            f"WHERE lawd_cd=? AND api_url=? AND ym IN ({ph}) AND fetched_at>?",
            [lawd_cd, api_url] + ym_list + [self._cutoff()],
        ).fetchall()
        return {ym: json.loads(items_json) for ym, items_json in rows}

    def set(self, key: tuple, items: list):
        lawd_cd, ym, api_url = key
        self._conn.execute(
            "INSERT OR REPLACE INTO trade_cache "
            "(lawd_cd, ym, api_url, items_json, fetched_at) VALUES (?,?,?,?,?)",
            (lawd_cd, ym, api_url,
             json.dumps(items, ensure_ascii=False),
             datetime.now().isoformat()),
        )
        self._pending += 1
        if self._pending >= 50:
            self._conn.commit()
            self._pending = 0

    def commit(self):
        self._conn.commit()
        self._pending = 0

    def count_valid(self) -> int:
        return self._conn.execute(
            "SELECT COUNT(*) FROM trade_cache WHERE fetched_at>?", (self._cutoff(),)
        ).fetchone()[0]

    def count_expired(self) -> int:
        return self._conn.execute(
            "SELECT COUNT(*) FROM trade_cache WHERE fetched_at<=?", (self._cutoff(),)
        ).fetchone()[0]

    def purge_expired(self):
        """만료된 항목 정리 → DB 파일 크기 절감."""
        deleted = self._conn.execute(
            "DELETE FROM trade_cache WHERE fetched_at<=?", (self._cutoff(),)
        ).rowcount
        self._conn.commit()
        if deleted:
            logger.info(f"[Trade Cache] 만료 항목 {deleted:,}건 삭제")

    def close(self):
        self.commit()
        self._conn.close()


# ──────────────────────────────────────────────
# ⑪ Kakao 주소 API [㉒ ㉜]
# ──────────────────────────────────────────────
KAKAO_URL    = "https://dapi.kakao.com/v2/local/search/address.json"
_KAKAO_RETRY = {429, 500, 502, 503, 504}   # [㉒] 재시도 가능한 HTTP 코드
JUSO_URL     = "https://www.juso.go.kr/addrlink/addrLinkApi.do"  # ② 행안부 도로명주소 검색 API


def _call_juso_once(address: str, juso_key: str) -> list[dict]:
    """
    ② 행안부 도로명주소 검색 API → 지번 후보 반환.
    Kakao 대비 공식 지번 DB 직접 조회 → 정확도 높음.
    응답 필드: lnbrMnnm(본번), lnbrSlno(부번), jibunAddr, bdNm
    실패 시 빈 리스트 반환 → 호출자에서 Kakao 폴백
    """
    for attempt in range(3):
        try:
            resp = requests.get(JUSO_URL, params={
                "confmKey":     juso_key,
                "keyword":      address,
                "resultType":   "json",
                "countPerPage": "3",
                "currentPage":  "1",
            }, timeout=10)
            if resp.status_code != 200:
                time.sleep(2 ** attempt)
                continue
            data    = resp.json()
            err_cd  = data.get("results", {}).get("common", {}).get("errorCode", "")
            if err_cd != "0":
                return []   # 키 오류 등 → 폴백
            jusos = data.get("results", {}).get("juso", []) or []
            results = []
            for j in jusos[:3]:
                bon  = str(j.get("lnbrMnnm", "0")).strip().lstrip("0") or "0"
                bu   = str(j.get("lnbrSlno", "0")).strip().lstrip("0") or "0"
                jibun = bon if bu == "0" else f"{bon}-{bu}"
                dong  = str(j.get("emdNm", "")).strip()   # 읍면동명
                # sigunguCd: admCd 앞 5자리
                adm   = str(j.get("admCd", "")).strip()
                sigungu_cd = adm[:5] if len(adm) >= 5 else ""
                bjdong_cd  = adm[5:9] if len(adm) >= 9 else ""
                if jibun and jibun != "0":
                    results.append({
                        "dong":       dong,
                        "bonbun":     bon,
                        "bubun":      bu,
                        "jibun":      jibun,
                        "sigungu_cd": sigungu_cd,
                        "bjdong_cd":  bjdong_cd,
                    })
            return results
        except requests.RequestException:
            time.sleep(2 ** attempt)
    return []


def _extract_kakao_candidates(docs: list) -> list[dict]:
    """
    Kakao API 응답 docs → 최대 3개 후보 지번 추출.
    b_code(법정동코드 10자리) 추가 저장:
      앞 5자리 = sigunguCd, 다음 4자리 = bjdongCd
    """
    results = []
    for doc in docs[:3]:
        addr_obj = doc.get("address") or {}
        dong     = str(addr_obj.get("region_3depth_name", "")).strip()
        bonbun   = str(addr_obj.get("main_address_no",   "")).strip().lstrip("0") or "0"
        bubun    = str(addr_obj.get("sub_address_no",    "")).strip().lstrip("0") or "0"
        jibun    = bonbun if bubun == "0" else f"{bonbun}-{bubun}"
        b_code   = str(addr_obj.get("b_code", "")).strip()   # 법정동코드 10자리
        # b_code 예: "1165010100" → sigunguCd="11650", bjdongCd="1010"
        sigungu_cd = b_code[:5]  if len(b_code) >= 5 else ""
        bjdong_cd  = b_code[5:9] if len(b_code) >= 9 else ""
        if dong:
            results.append({
                "dong":       dong,
                "bonbun":     bonbun,
                "bubun":      bubun,
                "jibun":      jibun,
                "sigungu_cd": sigungu_cd,
                "bjdong_cd":  bjdong_cd,
            })
    return results


def _call_kakao_once(address: str, kakao_key: str) -> list[dict]:
    """
    Kakao 주소 검색 → 최대 3개 후보 반환.
    [㉒] 503/429 등 일시적 오류 → requests.RequestException 발생
         → 호출자(build_jibun_cache)에서 network_error로 저장 → 재시도
    [㉜] 429 발생 시 지수 백오프
    """
    for attempt in range(CONFIG["MAX_RETRY"]):
        resp = requests.get(
            KAKAO_URL,
            headers={"Authorization": f"KakaoAK {kakao_key}"},
            params={"query": address, "analyze_type": "exact"},
            timeout=10,
        )

        # [㉒] 일시적 오류 → 지수 백오프 후 재시도
        if resp.status_code in _KAKAO_RETRY:
            wait = 2 ** attempt
            logger.warning(f"[Kakao] HTTP {resp.status_code} — {wait}초 대기 후 재시도")
            time.sleep(wait)
            if attempt == CONFIG["MAX_RETRY"] - 1:
                raise requests.RequestException(f"Kakao HTTP {resp.status_code} 최대 재시도 초과")
            continue

        if resp.status_code == 401:
            logger.error("[Kakao] 인증 실패 — KAKAO_API_KEY 확인 필요")
            return []

        if resp.status_code != 200:
            return []   # 영구적 오류 → not_found

        docs = resp.json().get("documents", [])
        if not docs:
            # [⑯] similar 재시도 시 KAKAO_INTERVAL 적용
            time.sleep(CONFIG["KAKAO_INTERVAL"])
            resp2 = requests.get(
                KAKAO_URL,
                headers={"Authorization": f"KakaoAK {kakao_key}"},
                params={"query": address, "analyze_type": "similar"},
                timeout=10,
            )
            # [㉒] similar도 일시적 오류면 network_error로
            if resp2.status_code in _KAKAO_RETRY:
                raise requests.RequestException(f"Kakao similar HTTP {resp2.status_code}")
            docs = resp2.json().get("documents", []) if resp2.status_code == 200 else []

        return _extract_kakao_candidates(docs)

    return []  # 루프 정상 종료 (도달 불가 — 방어용)


# ──────────────────────────────────────────────
# ⑫ Jibun 캐시 (버전 관리) [③ ㉙]
# ──────────────────────────────────────────────
JIBUN_CACHE_VERSION = 3   # v3: candidates에 sigungu_cd·bjdong_cd 추가


def migrate_jibun_cache(data: dict) -> dict:
    """
    구버전 {addr: {dong, jibun, ...}} →
    신버전 {addr: {status, candidates}} 변환.
    """
    migrated, cnt = {}, 0
    for addr, val in data.items():
        if isinstance(val, dict) and "status" not in val:
            if val.get("dong"):
                migrated[addr] = {"status": "ok", "candidates": [val]}
            else:
                migrated[addr] = {"status": "not_found", "candidates": []}
            cnt += 1
        else:
            migrated[addr] = val
    if cnt:
        logger.info(f"[Kakao] 구버전 캐시 {cnt:,}건 자동 변환")
    return migrated


def load_jibun_cache(path: str) -> dict:
    if not Path(path).exists():
        return {}
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    meta = raw.get("_meta", {})
    data = {k: v for k, v in raw.items() if k != "_meta"}
    if meta.get("version") == JIBUN_CACHE_VERSION:
        return data
    # 버전 미달 → migrate 후 반환 (재수집은 build_jibun_cache에서 처리)
    return migrate_jibun_cache(data)


def _needs_bcode_refresh(cache: dict) -> bool:
    """
    기존 캐시 중 ok 항목에 bjdong_cd가 없으면 True → 재수집 필요.
    v3 이전 캐시는 b_code를 저장하지 않았으므로 재수집 필요.
    """
    for val in cache.values():
        if val.get("status") == "ok":
            cands = val.get("candidates", [])
            if cands and "bjdong_cd" not in cands[0]:
                return True
            return False  # 첫 ok 항목에 bjdong_cd 있으면 최신 버전
    return False


def save_jibun_cache(path: str, cache: dict):
    out = {
        "_meta": {"version": JIBUN_CACHE_VERSION, "saved_at": datetime.now().isoformat()},
        **cache,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def clean_address_for_jibun(addr: str) -> str:
    """
    [㉟] 지번 변환 API 검색용 주소 정제.
    행안부/Kakao API가 처리 못하는 부가 텍스트를 제거해 검색 성공률 향상.
    원본 주소(캐시 키)는 변경하지 않고 검색 쿼리에만 사용.
      1) 괄호가 있으면 마지막 괄호까지만 유지
         "숙선옹주로5길 10-23(묵동) 103-18외 1필지" → "숙선옹주로5길 10-23(묵동)"
         "한천로120길 24(번동) 해날하이빌"           → "한천로120길 24(번동)"
      2) 괄호 없으면 시도+도로명+번지까지만 유지
         "시루봉로15라길 8-6 A동"                   → "시루봉로15라길 8-6"
    """
    s = str(addr).strip()

    # 1) 괄호가 있으면 마지막 괄호까지만 유지
    m = re.search(r'\([^)]+\)', s)
    if m:
        return s[:m.end()].strip()

    # 2) 괄호 없으면 시도+도로명+번지까지만 유지
    m = re.search(r'^(.*?)((?:[가-힣a-zA-Z0-9]+(로|길|대로))\s+[\d-]+)', s)
    if m:
        return m.group(0).strip()

    return s


def build_jibun_cache(addresses: list[str], kakao_key: str, cache_path: str) -> dict:
    """
    ② 비아파트 주소 → 지번 변환 (행안부 우선, Kakao 폴백).
    행안부: 공식 주소 DB 직접 조회 → 정확도 높음
    Kakao:  행안부 실패 시 폴백
    [㉟] 검색 전 주소 정제(clean_address_for_jibun) → not_found 감소
    캐시 구조: {주소: {"status": ok|not_found|network_error, "candidates": [...]}}
    """
    juso_key = CONFIG.get("JUSO_API_KEY", "")
    cache    = load_jibun_cache(cache_path)

    need = [
        a for a in addresses
        if a not in cache or cache[a].get("status") == "network_error"
    ]

    if not need:
        ok  = sum(1 for v in cache.values() if v.get("status") == "ok")
        nf  = sum(1 for v in cache.values() if v.get("status") == "not_found")
        err = sum(1 for v in cache.values() if v.get("status") == "network_error")
        src = "행안부+Kakao" if juso_key else "Kakao"
        logger.info(f"[{src}] 캐시 전량 사용 — 성공:{ok:,} / 미발견:{nf:,} / 오류:{err:,}")
        return cache

    use_juso = bool(juso_key)
    logger.info(f"[지번변환] {len(need):,}건 — {'행안부 우선 + Kakao 폴백' if use_juso else 'Kakao 단독'}")
    ok_juso = ok_kakao = nf_cnt = err_cnt = 0

    for addr in tqdm(need, desc="지번 변환"):
        candidates = []
        source     = ""
        # [㉟] 검색용 주소 정제 (원본 캐시 키는 addr 그대로 유지)
        query = clean_address_for_jibun(addr)
        try:
            # ① 행안부 먼저 시도
            if use_juso:
                candidates = _call_juso_once(query, juso_key)
                if candidates:
                    source = "juso"

            # ② 행안부 실패 시 Kakao 폴백
            if not candidates and kakao_key:
                candidates = _call_kakao_once(query, kakao_key)
                if candidates:
                    source = "kakao"

            if candidates:
                cache[addr] = {"status": "ok", "candidates": candidates, "source": source}
                if source == "juso": ok_juso += 1
                else:                ok_kakao += 1
            else:
                cache[addr] = {"status": "not_found", "candidates": []}
                nf_cnt += 1
        except requests.RequestException:
            cache[addr] = {"status": "network_error", "candidates": []}
            err_cnt += 1

        time.sleep(CONFIG["KAKAO_INTERVAL"])
        if (ok_juso + ok_kakao + nf_cnt + err_cnt) % 100 == 0:
            save_jibun_cache(cache_path, cache)

    save_jibun_cache(cache_path, cache)
    logger.info(
        f"[지번변환] 완료: 행안부 {ok_juso:,} / Kakao폴백 {ok_kakao:,} / "
        f"미발견 {nf_cnt:,} / 오류 {err_cnt:,}"
    )
    return cache


# ──────────────────────────────────────────────
# ⑬ 유틸리티
# ──────────────────────────────────────────────
def get_ym_range(start: str, end: str) -> list[str]:
    result, cur = [], datetime.strptime(start, "%Y%m")
    end_dt = datetime.strptime(end, "%Y%m")
    while cur <= end_dt:
        result.append(cur.strftime("%Y%m"))
        cur = (cur + timedelta(days=32)).replace(day=1)
    return result


def make_dedup_key(df: pd.DataFrame) -> pd.Series:
    """
    [⑨ ⑱] dedup_key 단일 함수 정의.
    전용면적 NaN → "NA" (float 정밀도 불일치 방지).
    """
    def _area_str(x):
        try:
            v = float(x)
            return "NA" if pd.isna(v) else str(round(v, 4))
        except Exception:
            return "NA"

    return (
        df["도로명주소"].fillna("").str.strip() + "||"
        + df["전용면적"].apply(_area_str) + "||"
        + df["주택유형"].fillna("").str.strip()
    )


def get_lawd_cd(row: pd.Series) -> "str | None":
    sido    = str(row.get("광역시도", ""))
    sigungu = str(row.get("시군구",   ""))
    address = str(row.get("도로명주소", ""))

    # [⑳] 세종 특수 처리
    if "세종" in sido or "세종" in sigungu:
        return "36110"

    code = SIGUNGU_CODE.get(sigungu)
    if code:
        return code

    prefix_map = [
        ("부산", "부산{}"), ("대구", "대구{}"), ("인천", "인천{}"),
        ("광주", "광주{}"), ("대전", "대전{}"), ("울산", "울산{}"),
        ("강원특별자치도", "{}_강원"), ("강원", "{}_강원"), ("경상남도", "{}_경남"),
    ]
    for kw, fmt in prefix_map:
        if kw in sido:
            code = SIGUNGU_CODE.get(fmt.format(sigungu))
            if code:
                return code

    m = re.search(r"([가-힣]+구)\s", address)
    if m:
        gu = m.group(1)
        code = SIGUNGU_CODE.get(gu)
        if code:
            return code
        for kw, fmt in prefix_map:
            if kw in sido:
                code = SIGUNGU_CODE.get(fmt.format(gu))
                if code:
                    return code
        code = SIGUNGU_CODE.get(f"{sigungu}{gu}")
        if code:
            return code
    return None


def get_api_info(housing_type: str) -> tuple[str, bool]:
    """
    [⑮ ㉘] 완전일치 우선 → 긴 키 우선 부분일치.
    sorted()는 _SORTED_HOUSING_KEYS에 캐싱되어 반복 호출 없음.
    """
    ht = str(housing_type).strip()
    if ht in HOUSING_TYPE_API:
        return HOUSING_TYPE_API[ht]
    for key in _SORTED_HOUSING_KEYS:
        if key in ht:
            return HOUSING_TYPE_API[key]
    return DEFAULT_API


def build_unique_rows(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["dedup_key"]       = make_dedup_key(df)
    df["region_priority"] = df["광역시도"].apply(get_region_priority)
    unique_df = (
        df.sort_values("region_priority")
          .drop_duplicates(subset="dedup_key", keep="first")
          .sort_values("region_priority")
          .reset_index(drop=True)
    )
    logger.info(
        f"[중복제거] {len(df):,}건 → {len(unique_df):,}건 "
        f"(API 호출 {(1-len(unique_df)/len(df))*100:.1f}% 절감)"
    )
    labels = {1: "서울", 2: "수도권(경기·인천)", 3: "광역시·세종", 4: "기타"}
    for p in sorted(labels):
        cnt = (unique_df["region_priority"] == p).sum()
        logger.info(f"  우선순위{p} {labels[p]}: {cnt:,}건")
    return unique_df


# ──────────────────────────────────────────────
# ⑭ 국토부 API 호출
#    [⑧] page_no attempt마다 리셋
#    [㉓] Semaphore가 _increment_call + 요청을 원자적으로 감쌈
# ──────────────────────────────────────────────
def xml_to_items(xml_text: str) -> list[dict]:
    try:
        root = ET.fromstring(xml_text)
        return [{c.tag: c.text for c in item} for item in root.findall(".//item")]
    except ET.ParseError as e:
        logger.warning(f"XML 파싱 오류: {e}")
        return []


def call_trade_api(lawd_cd: str, deal_ymd: str, api_url: str,
                   service_key: str, max_retry: int = 3,
                   interval: float = 0.2) -> list[dict]:
    if api_url in _api_exhausted:
        return []

    sem = get_semaphore(api_url)

    for attempt in range(max_retry):
        all_items, page_no = [], 1   # [⑧] attempt마다 리셋
        try:
            while True:
                # [㉓] Semaphore 안에서 카운트 + 요청 → 원자적
                with sem:
                    if not _increment_call(api_url):
                        raise APIExhaustedException(api_url)
                    resp = requests.get(api_url, params={
                        "serviceKey": service_key, "LAWD_CD": lawd_cd,
                        "DEAL_YMD": deal_ymd, "pageNo": page_no, "numOfRows": 1000,
                    }, timeout=30)

                if resp.status_code == 429:
                    with _count_lock:
                        _api_exhausted.add(api_url)
                    logger.warning(f"[429] {api_url.split('/')[-1]} 한도 초과")
                    raise APIExhaustedException(api_url)

                resp.raise_for_status()
                items = xml_to_items(resp.text)
                all_items.extend(items)
                if len(items) < 1000:
                    break
                page_no += 1
                time.sleep(interval)

            return all_items   # [⑧] 성공 시 즉시 반환

        except APIExhaustedException:
            return []
        except requests.RequestException as e:
            logger.warning(f"API 오류 ({lawd_cd}/{deal_ymd}) 시도 {attempt+1}: {e}")
            time.sleep(2 ** attempt)

    logger.error(f"API 최대 재시도 초과: {lawd_cd}/{deal_ymd}")
    return []


# ──────────────────────────────────────────────
# ⑮ 주소 파싱 유틸
# ──────────────────────────────────────────────
def parse_floor(v) -> "int | None":
    if v is None:
        return None
    s = str(v).replace("층", "").strip()
    if not s:
        return None
    try:
        if "지하" in s or s.upper().startswith("B"):
            n = re.sub(r"[^0-9]", "", s)
            return -int(n) if n else None
        n = re.sub(r"[^0-9-]", "", s).strip("-")
        return int(n.split("-")[0]) if n else None
    except (ValueError, IndexError):
        return None


def parse_area(v) -> "float | None":
    try:
        return round(float(str(v).replace(",", "").strip()), 4)
    except (ValueError, TypeError):
        return None


def normalize_jibun(bonbun: str, bubun: str) -> str:
    b = str(bonbun).strip().lstrip("0") or "0"
    s = str(bubun ).strip().lstrip("0") or "0"
    return b if s == "0" else f"{b}-{s}"


def parse_jibun_field(jibun_str: str) -> str:
    """
    [⑰] 국토부 API jibun 필드 정규화.
    공백("123 45")·하이픈("123-45") 구분자 모두 처리.
    """
    s = str(jibun_str or "").strip()
    if not s:
        return ""
    m = re.match(r"(\d+)(?:[-\s]+(\d+))?", s)
    if not m:
        return ""
    bonbun = str(int(m.group(1)))
    bubun  = str(int(m.group(2))) if m.group(2) else "0"
    return normalize_jibun(bonbun, bubun)


def extract_dong_from_address(address: str) -> str:
    """
    [⑬ ㉑ ㉞] 법정동명 추출.
    여러 괄호: 마지막 괄호부터 시도 (건물명 괄호가 먼저 오는 주소 대응).
    괄호 안 매칭 우선순위 (첫 토큰 기준):
      1) 숫자+가 패턴: 을지로2가, 보수동1가, 동소문동3가, 덕진동2가
      2) 동/리/읍/면 + 뒤에 한글(지 제외)/영문이 붙은 경우 → 동명만 추출
         예) 효자동효자코아루웰라움 → 효자동 / 송내동LOFT → 송내동
         '지' 제외: 나너울빌리지 등에서 '빌리' 오탐 방지
      3) 순수하게 동/리/읍/면으로 끝나는 경우
    괄호 없음: 토큰 기반 폴백 (숫자+가 포함)
    """
    parens = re.findall(r'\(([가-힣0-9a-zA-Z ]+)\)', str(address))

    for inner in reversed(parens):
        inner = inner.strip()
        first_token = inner.split()[0] if inner.split() else inner

        # 1) 숫자+가 패턴
        m = re.match(r'^([가-힣]+\d+가)', first_token)
        if m:
            return m.group(1)

        # 2) 동/리/읍/면 + 뒤에 건물명(한글/영문) 붙은 경우 → 동명만
        #    '지' 제외: 빌리지·타운지 등에서 '리/동' 오탐 방지
        m = re.match(r'^([가-힣]{2,}(?:동|리|읍|면))(?=[가-힣A-Za-z])(?!지)', first_token)
        if m:
            return m.group(1)

        # 3) 순수 동명 (동/리/읍/면으로 끝나는 경우)
        m = re.match(r'^([가-힣0-9]{2,}(?:동|리|읍|면))$', first_token)
        if m:
            return m.group(1)

    # [㉑] 괄호 없음 → 토큰에서 동/가/리 추출
    for token in str(address).split():
        if re.search(r'[가-힣]+(\d+가|동|리|읍|면)$', token):
            return token
    return ""


def extract_road_parts(address: str) -> tuple[str, str]:
    norm = re.sub(r'\(.*?\)', '', str(address)).strip()
    m    = re.search(r'([가-힣a-zA-Z0-9]+(로|길|대로))\s+(\d+(?:-\d+)?)', norm)
    return (m.group(1), m.group(3)) if m else ("", "")


# ──────────────────────────────────────────────
# ⑯ 매칭 등급 판정
# ──────────────────────────────────────────────
AREA_TOLERANCE  = 3.0   # ④ 완화: 3㎡ (기존 1.0㎡)
FLOOR_TOLERANCE = 2     # ④ 완화: ±2층 (기존 ±1층)

GRADE_LABEL = {
    1: "Grade1_정밀(도로명+층+면적 또는 지번일치)",
    2: "Grade2_근접(지번+면적,층±1)",
    3: "Grade3_참고(법정동+층+면적)",
    4: "Grade4_낮음(법정동만)",
}


def grade_match_apt(
    item: dict,
    lh_road:     str,
    lh_road_num: str,
    lh_dong:     str,
    target_floor: "int | None",
    target_area:  "float | None",
) -> "int | None":
    """
    아파트 매칭.
    [⑫] road_match라도 면적·층 모두 불일치 시 None 반환.
    """
    item_road   = (item.get("roadNm")       or "").strip()
    item_bonbun = (item.get("roadNmBonbun") or "").strip().lstrip("0")

    road_match = False
    if lh_road and item_road and lh_road == item_road:
        road_match = (
            (lh_road_num.split("-")[0] == item_bonbun)
            if (lh_road_num and item_bonbun)
            else True
        )

    item_umd   = (item.get("umdNm") or "").strip()
    # [양방향 dong 매칭] 추출된 dong이 API값보다 길어도 처리
    # 예) lh_dong='공덕동공덕헤리' / item_umd='공덕동' → item_umd in lh_dong → True
    dong_match = bool(lh_dong and item_umd and (
        lh_dong in item_umd or item_umd in lh_dong
    ))

    if not (road_match or dong_match):
        return None

    item_floor  = parse_floor(item.get("floor"))
    item_area   = parse_area(item.get("excluUseAr"))
    floor_exact = target_floor is not None and item_floor is not None and abs(target_floor - item_floor) == 0
    floor_near  = target_floor is not None and item_floor is not None and abs(target_floor - item_floor) <= FLOOR_TOLERANCE
    area_ok     = target_area  is not None and item_area  is not None and abs(target_area  - item_area)  <= AREA_TOLERANCE

    if road_match:
        if floor_exact and area_ok: return 1
        if floor_near  and area_ok: return 2
        if area_ok:                 return 2
        return None   # [⑫] 면적·층 모두 불일치 → 매칭 제외
    else:
        if (floor_exact or floor_near) and area_ok: return 3
        return 4


def grade_match_non_apt(
    item: dict,
    lh_dong:      str,
    lh_jibun:     str,
    target_floor: "int | None",
    target_area:  "float | None",
    api_key:      str,
    lh_road:      str = "",
    lh_road_num:  str = "",
) -> "int | None":
    """
    비아파트 매칭 등급 판정.

    [SHTrade - 다가구/단독]
      API 응답에 floor·area·roadNm 없음 → jibun 일치만으로 최선 판단
      jibun_exact  → G1 (본번+부번 완전일치)
      jibun_bonbun → G2 (본번만 일치)
      else         → G4

    [RHTrade/OffiTrade - 다세대/연립/오피스텔/도시형]
      jibun + floor + area 3중 검증
      jibun_exact + floor_exact + area_ok → G1
      jibun_exact + floor_near  + area_ok → G2
      jibun_exact + area_ok               → G2
      jibun_exact                         → G3
      jibun_bonbun + floor + area         → G2~G3
      else                                → G3~G4
    """
    meta = API_META.get(
        api_key,
        {"has_road": False, "has_jibun": True, "has_floor": True, "has_area": True},
    )

    item_umd = (item.get("umdNm") or "").strip()
    # 양방향 dong 매칭: lh_dong='덕진동2가' / item_umd='덕진동' 모두 허용
    dong_match = bool(lh_dong and item_umd and (
        lh_dong in item_umd or item_umd in lh_dong
    ))
    if not dong_match:
        return None

    # ── 지번 비교 ──────────────────────────────────────────────────────
    jibun_exact = jibun_bonbun = False
    if meta["has_jibun"] and lh_jibun:
        # OffiTrade API는 jibun 대신 ji 필드명 사용
        item_jibun = parse_jibun_field(item.get("jibun") or item.get("ji", ""))
        if item_jibun:
            jibun_exact  = (lh_jibun == item_jibun)
            jibun_bonbun = (lh_jibun.split("-")[0] == item_jibun.split("-")[0])

    # ── [SHTrade] floor·area 없음 → jibun 일치만으로 등급 결정 ─────────
    if not meta["has_floor"] and not meta["has_area"]:
        if jibun_exact:  return 1  # 지번 완전일치 → G1 (최선)
        if jibun_bonbun: return 2  # 본번만 일치   → G2
        return 4                   # 미일치         → G4

    # ── [RHTrade / OffiTrade] floor + area 있음 → 3중 검증 ────────────
    item_floor  = parse_floor(item.get("floor"))
    item_area   = parse_area(item.get("excluUseAr"))
    floor_exact = target_floor is not None and item_floor is not None and abs(target_floor - item_floor) == 0
    floor_near  = target_floor is not None and item_floor is not None and abs(target_floor - item_floor) <= FLOOR_TOLERANCE
    area_ok     = target_area  is not None and item_area  is not None and abs(target_area  - item_area)  <= AREA_TOLERANCE

    if jibun_exact:
        if floor_exact and area_ok: return 1
        if floor_near  and area_ok: return 2
        if area_ok:                 return 2
        return 3
    elif jibun_bonbun:
        if floor_exact and area_ok: return 2
        if floor_near  and area_ok: return 3
        return 4
    else:
        if (floor_exact or floor_near) and area_ok: return 3
        return 4


def match_trade_to_address(
    trade_items:       list,
    target_floor:      "int | None",
    target_area:       "float | None",
    api_url:           str,
    is_apt:            bool,
    jibun_cache_entry: dict,
    # [최적화③] 주소 파싱 결과를 사전 계산해서 전달 (38개월 반복 파싱 제거)
    lh_road:          str = "",
    lh_road_num:      str = "",
    lh_dong:          str = "",
    candidates:       "list | None" = None,
) -> list:
    """
    실거래 목록 매칭.
    [㉔] 정렬 제거 → run()에서 월별 통합 후 1회만 정렬.
    [최적화③] lh_road·lh_road_num·lh_dong·candidates를 호출자가 사전 계산해 전달.
             → ym 루프 내 반복 파싱(2,673,620회) 완전 제거.
    다필지 후보 순회 → 각 item 최선 등급 채택.
    """
    api_key = api_url.rstrip("/").split("/")[-1]

    # candidates가 None이면 이전 호환 방식으로 폴백
    if candidates is None:
        if not is_apt:
            candidates = jibun_cache_entry.get("candidates", []) if jibun_cache_entry else []
            if not candidates:
                candidates = [{"dong": lh_dong, "jibun": ""}]
        else:
            candidates = []

    results = []
    for item in trade_items:
        if is_apt:
            grade = grade_match_apt(
                item, lh_road, lh_road_num, lh_dong,
                target_floor, target_area,
            )
            if grade is not None:
                item = dict(item)
                item["_match_grade"] = grade
                results.append(item)
        else:
            best_grade: "int | None" = None
            for cand in candidates:
                grade = grade_match_non_apt(
                    item,
                    cand.get("dong", ""),
                    cand.get("jibun", ""),
                    target_floor, target_area,
                    api_key,
                    lh_road=lh_road,
                    lh_road_num=lh_road_num,
                )
                if grade is not None:
                    best_grade = min(best_grade, grade) if best_grade is not None else grade

            if best_grade is not None:
                item = dict(item)
                item["_match_grade"] = best_grade
                results.append(item)

    # [㉔] 정렬 없이 반환 (run()에서 통합 정렬)
    return results


# ──────────────────────────────────────────────
# ⑰ 모듈 레벨 fetch_one [㉛ ㉕]
# ──────────────────────────────────────────────
def fetch_one(key: tuple) -> tuple[tuple, "list | None"]:
    """
    [㉛] 모듈 레벨 함수 → 테스트·재사용 가능.
    [㉕] _stop_event 체크 → exhausted 신호 시 즉시 반환.
    반환: (key, items)  items=None → exhausted 또는 중단
    """
    if _stop_event.is_set():
        return key, None

    lawd_cd, ym, api_url = key
    if api_url in _api_exhausted:
        return key, None

    items = call_trade_api(
        lawd_cd, ym, api_url,
        CONFIG["SERVICE_KEY"], CONFIG["MAX_RETRY"], CONFIG["CALL_INTERVAL"],
    )
    # [⑥] exhausted로 인한 빈 결과는 None으로 반환 (DB 저장 안 함)
    if api_url in _api_exhausted:
        return key, None

    return key, items


# ──────────────────────────────────────────────
# ⑱ 메인 처리 로직
# ──────────────────────────────────────────────
def run():
    from collections import defaultdict, Counter

    Path(CONFIG["OUTPUT_DIR"]).mkdir(exist_ok=True)

    # 키 설정 검증
    if not CONFIG["SERVICE_KEY"]:
        logger.error("MOLIT_SERVICE_KEY가 .env에 없습니다. 종료합니다.")
        return
    if not CONFIG["JUSO_API_KEY"] and not CONFIG["KAKAO_API_KEY"]:
        logger.warning("JUSO_API_KEY·KAKAO_API_KEY 모두 없습니다. 법정동명 폴백으로 진행합니다.")

    # [㉕] 배치 중단 이벤트 초기화
    _stop_event.clear()

    # ── 데이터 로드 ───────────────────────────────
    logger.info("데이터 로드 중...")
    df_orig = pd.read_excel(CONFIG["INPUT_FILE"])
    logger.info(f"총 {len(df_orig):,}건 로드 완료")

    # [최적화⑦] get_api_info apply 1회 → api_url·is_apt 컬럼 동시 생성
    df_orig[["api_url", "is_apt"]] = pd.DataFrame(
        df_orig["주택유형"].apply(get_api_info).tolist(),
        index=df_orig.index,
        columns=["api_url", "is_apt"],
    )
    df_orig["lawd_cd"] = df_orig.apply(get_lawd_cd, axis=1)

    missing = df_orig["lawd_cd"].isna().sum()
    if missing:
        logger.warning(f"법정동코드 미매핑 {missing:,}건")
        df_orig[df_orig["lawd_cd"].isna()].to_excel(
            f"{CONFIG['OUTPUT_DIR']}/unmapped_addresses.xlsx", index=False
        )
    df_orig = df_orig.dropna(subset=["lawd_cd"]).copy()

    # ── 중복 제거 & 우선순위 정렬 ────────────────
    df_unique = build_unique_rows(df_orig)

    # ── Kakao 지번 변환 (다가구·단독 포함 전체 비아파트) ─
    non_apt_addrs = (
        df_unique[~df_unique["is_apt"]]["도로명주소"]
        .dropna().unique().tolist()
    )
    logger.info(f"지번변환 대상: {len(non_apt_addrs):,}개 (비아파트, 행안부+Kakao폴백)")

    jibun_cache = (
        build_jibun_cache(non_apt_addrs, CONFIG["KAKAO_API_KEY"], CONFIG["JIBUN_CACHE_FILE"])
        if CONFIG["JUSO_API_KEY"] or CONFIG["KAKAO_API_KEY"]
        else {}
    )

    # ── [㉖ ㉗] Trade Cache DB 초기화 ────────────
    trade_db = TradeCacheDB(CONFIG["TRADE_DB_FILE"], CONFIG["CACHE_TTL_DAYS"])
    trade_db.purge_expired()
    logger.info(f"[Trade Cache] 유효: {trade_db.count_valid():,}건 (만료 삭제 후)")

    # ── 신규 호출 조합 구성 (캐시에 없는 것만) ────
    ym_list = get_ym_range(CONFIG["START_YMD"], CONFIG["END_YMD"])
    logger.info(f"조회 기간: {ym_list[0]} ~ {ym_list[-1]} ({len(ym_list)}개월)")

    # [최적화①] 유효 키를 1회 SQL로 전체 로드 → 275만 번 개별 exists() 제거
    cached_keys: set[tuple] = trade_db.get_all_valid_keys()
    logger.info(f"[Trade Cache] 유효 키 {len(cached_keys):,}건 메모리 로드 완료")

    # [최적화④] defaultdict로 priority 그룹 사전 구성 → 1,098만 번 순회 제거
    batches: dict[int, list] = defaultdict(list)
    seen:    set[tuple]      = set()

    # [최적화⑤] combos 구성: lawd_cd·api_url·region_priority 영문 컬럼만 → itertuples
    for row in df_unique[["lawd_cd", "api_url", "region_priority"]].itertuples(index=False):
        for ym in ym_list:
            key = (row.lawd_cd, ym, row.api_url)
            if key not in cached_keys and key not in seen:
                batches[int(row.region_priority)].append(key)
                seen.add(key)

    total_combos = sum(len(v) for v in batches.values())
    logger.info(f"신규 API 호출 조합: {total_combos:,}건")

    # ── [② ㉕] 우선순위 배치 처리 ────────────────
    logger.info("실거래가 API 조회 시작 (서울 → 수도권 → 광역시 → 기타)...")
    skipped = 0

    for priority in sorted(batches):
        if _stop_event.is_set():
            logger.warning(f"한도 소진 — 우선순위 {priority}+ 배치 건너뜀")
            break

        batch = batches[priority]
        logger.info(f"  우선순위 {priority}: {len(batch):,}건")

        with ThreadPoolExecutor(max_workers=CONFIG["MAX_WORKERS"]) as executor:
            futures = {executor.submit(fetch_one, k): k for k in batch}

            for future in tqdm(as_completed(futures), total=len(futures),
                               desc=f"P{priority} 호출"):
                try:
                    key, items = future.result()

                    if items is None:
                        skipped += 1
                        # [㉕] 중단 신호 + 미시작 future 취소
                        if not _stop_event.is_set():
                            _stop_event.set()
                            for f in futures:
                                f.cancel()
                    else:
                        trade_db.set(key, items)

                except concurrent.futures.CancelledError:
                    skipped += 1
                except Exception as e:
                    logger.error(f"[⑭] fetch_one 예외: {e}")
                    trade_db.commit()  # 즉시 저장 (진행분 보호)

        trade_db.commit()
        logger.info(f"  → Trade Cache 저장 완료 (유효: {trade_db.count_valid():,}건)")

    logger.info("=== API별 일일 호출 현황 ===")
    for url, cnt in _api_call_counts.items():
        status = "⛔ 소진" if url in _api_exhausted else "✅ 정상"
        logger.info(f"  {status} | {url.split('/')[-1]}: {cnt:,}건")
    if skipped:
        logger.warning(f"미완료: {skipped:,}건 → 재실행 시 자동 재시도")

    # ── [㉖] 주소 매칭 — 그룹 단위 로드로 메모리 절감 ──
    logger.info("주소 매칭 시작...")
    match_cache: dict[str, dict] = {}

    # [최적화⑧] df_unique.copy() 제거 — groupby는 원본 변형하지 않음
    groups = df_unique.groupby(["lawd_cd", "api_url"], sort=False)

    for (lawd_cd, api_url), group_df in tqdm(groups, desc="그룹 매칭",
                                              total=groups.ngroups):
        # 단일 SQL로 해당 그룹의 모든 월 데이터 로드 [㉖]
        group_trade = trade_db.get_group(lawd_cd, api_url, ym_list)

        for _, row in group_df.iterrows():   # 한국어 컬럼 있어서 iterrows 유지
            address      = str(row.get("도로명주소", ""))
            dedup_key    = row["dedup_key"]
            is_apt       = row["is_apt"]
            housing_type = str(row.get("주택유형", ""))
            target_floor = parse_floor(row.get("층"))
            target_area  = parse_area(row.get("전용면적"))
            jibun_entry  = jibun_cache.get(address, {}) if not is_apt else {}

            # [최적화③] 주소 파싱 1회 (38개월 반복 제거)
            lh_dong = extract_dong_from_address(address)
            if is_apt:
                lh_road, lh_road_num = extract_road_parts(address)
                pre_candidates       = None  # grade_match_apt는 roadNm 직접 사용
            else:
                lh_road = lh_road_num = ""
                # 비아파트: 행안부+Kakao 지번 사용
                pre_candidates = jibun_entry.get("candidates", []) if jibun_entry else []
                if not pre_candidates:
                    pre_candidates = [{"dong": lh_dong, "jibun": ""}]

            # 로그용 변환 상태
            first_jibun    = ""
            cand_count     = 0
            convert_status = ""
            if not is_apt and jibun_entry:
                cands          = jibun_entry.get("candidates", [])
                first_jibun    = cands[0].get("jibun", "") if cands else ""
                cand_count     = len(cands)
                convert_status = jibun_entry.get("status", "")

            all_matches = []
            for ym in ym_list:
                trade_items = group_trade.get(ym, [])

                # [최적화⑨] 빈 리스트 early return → 불필요한 함수 진입 제거
                if not trade_items:
                    continue

                for m in match_trade_to_address(
                    trade_items,
                    target_floor, target_area,
                    api_url, is_apt, jibun_entry,
                    lh_road=lh_road,
                    lh_road_num=lh_road_num,
                    lh_dong=lh_dong,
                    candidates=pre_candidates,
                ):
                    m = dict(m)
                    m["조회연월"] = ym
                    all_matches.append(m)

            # [㉔] 통합 후 1회만 정렬
            all_matches.sort(key=lambda x: (
                x["_match_grade"],
                -(int(x.get("dealYear",  0) or 0) * 10_000
                  + int(x.get("dealMonth", 0) or 0) * 100
                  + int(x.get("dealDay",   0) or 0)),
            ))

            if all_matches:
                best = all_matches[0]
                g    = best["_match_grade"]

                # [최적화⑥] Counter로 grade 카운트 1회 순회로 통합
                grade_counts = Counter(m["_match_grade"] for m in all_matches)
                g1_cnt       = grade_counts[1]
                g2_cnt       = grade_counts[2]

                match_cache[dedup_key] = {
                    "매칭등급":          GRADE_LABEL[g],
                    "매칭_Grade1건수":   g1_cnt,
                    "매칭_Grade2건수":   g2_cnt,
                    "총매칭건수":        len(all_matches),
                    "거래금액(만원)":    best.get("dealAmount", ""),
                    "거래년도":          best.get("dealYear",   ""),
                    "거래월":            best.get("dealMonth",  ""),
                    "거래일":            best.get("dealDay",    ""),
                    "전용면적(㎡)_거래": best.get("excluUseAr", ""),
                    "층_거래":           best.get("floor",      ""),
                    "건축년도_거래":     best.get("buildYear",  ""),
                    "건물명_거래":       best.get("aptNm", best.get("mhouseNm", best.get("offiNm", ""))),
                    "변환지번":          first_jibun,
                    "변환지번_후보수":   cand_count,
                    "변환상태":          convert_status,
                    "조회연월범위":      f"{ym_list[0]}~{ym_list[-1]}",
                    "Grade1_거래금액":   all_matches[0].get("dealAmount", "") if g1_cnt > 0 else "",
                    "Grade1_거래년월":   (
                        f"{all_matches[0].get('dealYear','')}.{all_matches[0].get('dealMonth','')}"
                        if g1_cnt > 0 else ""
                    ),
                }
            else:
                match_cache[dedup_key] = {
                    "매칭등급": "매칭없음", "매칭_Grade1건수": 0, "매칭_Grade2건수": 0,
                    "총매칭건수": 0, "거래금액(만원)": "", "거래년도": "", "거래월": "", "거래일": "",
                    "전용면적(㎡)_거래": "", "층_거래": "", "건축년도_거래": "", "건물명_거래": "",
                    "변환지번": first_jibun, "변환지번_후보수": cand_count,
                    "변환상태": convert_status,
                    "조회연월범위": f"{ym_list[0]}~{ym_list[-1]}",
                    "Grade1_거래금액": "", "Grade1_거래년월": "",
                }

    trade_db.close()

    # ── 원본 전체에 조인 ──────────────────────────
    logger.info("원본 데이터 조인 중...")
    df_orig["dedup_key"] = make_dedup_key(df_orig)   # [⑨] 동일 함수

    match_df = pd.DataFrame.from_dict(match_cache, orient="index")
    match_df.index.name = "dedup_key"
    match_df = match_df.reset_index()

    out_df = df_orig.merge(match_df, on="dedup_key", how="left")

    # ── [㉝] 매칭 실패 디버그 파일 저장 ──────────
    no_match = out_df[out_df["매칭등급"].fillna("매칭없음") == "매칭없음"][
        ["도로명주소", "주택유형", "전용면적", "층",
         "광역시도", "시군구", "lawd_cd",
         "변환상태", "변환지번"]
    ]
    if len(no_match):
        debug_path = f"{CONFIG['OUTPUT_DIR']}/no_match_debug.xlsx"
        no_match.to_excel(debug_path, index=False)
        logger.info(f"[㉝] 매칭 실패 {len(no_match):,}건 → {debug_path}")

    # ── [⑪] 내부 컬럼 제거 ───────────────────────
    out_df = out_df.drop(
        columns=["dedup_key", "region_priority", "is_apt", "api_url", "lawd_cd"],
        errors="ignore",
    )

    logger.info("결과 저장 중...")
    out_path = f"{CONFIG['OUTPUT_DIR']}/result_realestate_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    out_df.to_excel(out_path, index=False)
    logger.info(f"완료: {out_path}")

    total       = max(len(out_df), 1)   # ZeroDivisionError 방어
    matched_cnt = out_df[out_df["총매칭건수"] > 0].shape[0]
    g1_cnt      = out_df[out_df["매칭_Grade1건수"] > 0].shape[0]
    g2_cnt      = out_df[out_df["매칭_Grade2건수"] > 0].shape[0]
    logger.info(f"전체:           {total:,}건")
    logger.info(f"매칭 성공:      {matched_cnt:,}건 ({matched_cnt/total*100:.1f}%)")
    logger.info(f"  Grade1(정밀): {g1_cnt:,}건 ({g1_cnt/total*100:.1f}%)")
    logger.info(f"  Grade2(근접): {g2_cnt:,}건 ({g2_cnt/total*100:.1f}%)")

    if jibun_cache:
        ok    = sum(1 for v in jibun_cache.values() if v.get("status") == "ok")
        nf    = sum(1 for v in jibun_cache.values() if v.get("status") == "not_found")
        err   = sum(1 for v in jibun_cache.values() if v.get("status") == "network_error")
        multi = sum(
            1 for v in jibun_cache.values()
            if v.get("status") == "ok" and len(v.get("candidates", [])) > 1
        )
        logger.info(
            f"[Kakao] 성공:{ok:,} (다필지:{multi:,}) / 미발견:{nf:,} / 오류:{err:,}"
        )


# ──────────────────────────────────────────────
# 실행 진입점
# ──────────────────────────────────────────────
if __name__ == "__main__":
    run()
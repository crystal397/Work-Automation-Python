#!/usr/bin/env python3
"""
국토교통부 실거래가 API 전수 수집기
=====================================
서울(25구) + 경기(43 시군구) × 최근 10년 × 8개 거래유형

사용법:
  python bulk_collector.py           # 수집 실행 (이어하기 지원)
  python bulk_collector.py --status  # 진행 현황만 출력

API 키 설정:
  .env 파일에  MOLIT_SERVICE_KEY=발급받은키  추가 (기존 lh_realestate_api.py 공유)
  또는 이 파일 상단 MOLIT_SERVICE_KEY 변수에 직접 입력
"""

import os
import sys
import re
import json
import time
import sqlite3
import logging
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, date
from pathlib import Path

# =============================================================
#  .env 로드
# =============================================================
def _load_env() -> dict:
    env: dict = {}
    parent = Path(__file__).parent
    p = None
    for folder in [parent, parent.parent]:
        candidate = folder / ".env"
        if candidate.exists():
            p = candidate
            break
    if p is None:
        return env
    with open(p, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            v = v.strip()
            if v and v[0] in ('"', "'") and v[0] in v[1:]:
                v = v[1:v.index(v[0], 1)]
            else:
                v = re.sub(r'\s+#.*$', '', v).strip()
            env[k.strip()] = v
    return env

_ENV = _load_env()

# =============================================================
#  ★ 설정
# =============================================================
MOLIT_SERVICE_KEY = _ENV.get("MOLIT_SERVICE_KEY", "")
#  .env 없이 직접 입력하려면 아래 줄 사용:
#  MOLIT_SERVICE_KEY = "여기에_API_키_입력"

DAILY_LIMIT   = 10000  # 하루 최대 API 호출 수 (100% 사용)
CALL_DELAY    = 0.5    # 호출 간격(초)
ROWS_PER_PAGE = 1000   # 페이지당 최대 행 수
YEARS_BACK    = 10     # 수집 기간(년)

BASE_DIR      = Path(__file__).parent
DB_PATH       = BASE_DIR / "bulk_data.db"
LOG_DIR       = BASE_DIR / "logs"
PROGRESS_PATH = LOG_DIR  / "progress.json"
ZERO_LOG_PATH = LOG_DIR  / "zero_records.log"  # 0건 수집 전용 로그

# =============================================================
#  API 엔드포인트 (공공데이터포털 apis.data.go.kr)
# =============================================================
_B = "https://apis.data.go.kr/1613000"

API_TYPES = {
    # 매매
    "apt_trade":  ("아파트 매매",       _B + "/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev"),
    "rh_trade":   ("연립다세대 매매",   _B + "/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade"),
    "sh_trade":   ("단독다가구 매매",   _B + "/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade"),
    "offi_trade": ("오피스텔 매매",     _B + "/RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade"),
}

# =============================================================
#  지역 코드 (서울 25구 + 경기 43 시군구)
# =============================================================
REGIONS = {
    # 서울
    "11110": "서울 종로구",     "11140": "서울 중구",
    "11170": "서울 용산구",     "11200": "서울 성동구",
    "11215": "서울 광진구",     "11230": "서울 동대문구",
    "11260": "서울 중랑구",     "11290": "서울 성북구",
    "11305": "서울 강북구",     "11320": "서울 도봉구",
    "11350": "서울 노원구",     "11380": "서울 은평구",
    "11410": "서울 서대문구",   "11440": "서울 마포구",
    "11470": "서울 양천구",     "11500": "서울 강서구",
    "11530": "서울 구로구",     "11545": "서울 금천구",
    "11560": "서울 영등포구",   "11590": "서울 동작구",
    "11620": "서울 관악구",     "11650": "서울 서초구",
    "11680": "서울 강남구",     "11710": "서울 송파구",
    "11740": "서울 강동구",
    # 경기
    "41111": "경기 수원시장안구",    "41113": "경기 수원시권선구",
    "41115": "경기 수원시팔달구",    "41117": "경기 수원시영통구",
    "41131": "경기 성남시수정구",    "41133": "경기 성남시중원구",
    "41135": "경기 성남시분당구",    "41150": "경기 의정부시",
    "41171": "경기 안양시만안구",    "41173": "경기 안양시동안구",
    "41192": "경기 부천시원미구",    "41194": "경기 부천시소사구",
    "41196": "경기 부천시오정구",    "41210": "경기 광명시",
    "41220": "경기 평택시",          "41250": "경기 동두천시",
    "41271": "경기 안산시상록구",    "41273": "경기 안산시단원구",
    "41281": "경기 고양시덕양구",    "41285": "경기 고양시일산동구",
    "41287": "경기 고양시일산서구",  "41290": "경기 과천시",
    "41310": "경기 구리시",          "41360": "경기 남양주시",
    "41370": "경기 오산시",          "41390": "경기 시흥시",
    "41410": "경기 군포시",          "41430": "경기 의왕시",
    "41450": "경기 하남시",          "41461": "경기 용인시처인구",
    "41463": "경기 용인시기흥구",    "41465": "경기 용인시수지구",
    "41480": "경기 파주시",          "41500": "경기 이천시",
    "41550": "경기 안성시",          "41570": "경기 김포시",
    "41591": "경기 화성시(향남)",    "41593": "경기 화성시(봉담)",
    "41595": "경기 화성시(병점)",    "41610": "경기 광주시",
    "41630": "경기 양주시",          "41650": "경기 포천시",
    "41670": "경기 여주시",          "41800": "경기 연천군",
    "41820": "경기 가평군",          "41830": "경기 양평군",
}

# =============================================================
#  예외
# =============================================================
class DailyLimitReached(Exception):
    pass

# =============================================================
#  시간 포맷 헬퍼
# =============================================================
def fmt_duration(seconds: float) -> str:
    seconds = int(seconds)
    h, r = divmod(seconds, 3600)
    m, s = divmod(r, 60)
    if h:
        return f"{h}시간 {m}분 {s}초"
    if m:
        return f"{m}분 {s}초"
    return f"{s}초"

# =============================================================
#  로깅
# =============================================================
def setup_logging():
    LOG_DIR.mkdir(exist_ok=True)
    today = date.today().strftime("%Y%m%d")
    log_file = LOG_DIR / ("bulk_collector_" + today + ".log")
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("bulk_collector")

def log_zero_record(api_type, api_name, region_code, region_name, year_month):
    """0건 수집 항목을 전용 로그 파일에 기록 (정상 여부 점검용)"""
    LOG_DIR.mkdir(exist_ok=True)
    with open(ZERO_LOG_PATH, "a", encoding="utf-8") as f:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"{ts}\t{api_type}\t{api_name}\t{region_code}\t{region_name}\t{year_month}\n")

# =============================================================
#  진행 상황 관리
# =============================================================
def load_progress():
    old_path = BASE_DIR / "progress.json"
    if old_path.exists() and not PROGRESS_PATH.exists():
        LOG_DIR.mkdir(exist_ok=True)
        old_path.rename(PROGRESS_PATH)

    if PROGRESS_PATH.exists():
        with open(PROGRESS_PATH, encoding="utf-8") as f:
            content = f.read().strip()
        if not content.replace('\x00', ''):
            data = {}
        else:
            data = json.loads(content)
        data.setdefault("today_date",          "")
        data.setdefault("today_calls",         0)
        data.setdefault("today_calls_by_type", {})
        data.setdefault("total_calls",         0)
        data.setdefault("total_records",       0)
        data.setdefault("completed",           [])
        return data
    return {
        "completed":           [],
        "today_date":          "",
        "today_calls":         0,
        "today_calls_by_type": {},
        "total_calls":         0,
        "total_records":       0,
    }

def save_progress(progress):
    tmp = PROGRESS_PATH.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)
    os.replace(tmp, PROGRESS_PATH)

def reset_daily_if_new_day(progress):
    today_str = date.today().strftime("%Y-%m-%d")
    if progress["today_date"] != today_str:
        progress["today_date"]          = today_str
        progress["today_calls"]         = 0
        progress["today_calls_by_type"] = {}

# =============================================================
#  SQLite 초기화
# =============================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            api_type     TEXT NOT NULL,
            region_code  TEXT NOT NULL,
            year_month   TEXT NOT NULL,
            data         TEXT NOT NULL,
            collected_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_main
        ON transactions (api_type, region_code, year_month)
    """)
    conn.commit()
    return conn

# =============================================================
#  API 호출
# =============================================================
def fetch_page(url, region_code, year_month, page, log):
    params = {
        "serviceKey": MOLIT_SERVICE_KEY,
        "LAWD_CD":    region_code,
        "DEAL_YMD":   year_month,
        "numOfRows":  ROWS_PER_PAGE,
        "pageNo":     page,
    }
    for attempt in range(1, 4):
        try:
            resp = requests.get(url, params=params, timeout=30)
            resp.raise_for_status()
            root = ET.fromstring(resp.content)

            result_code = (root.findtext(".//resultCode") or "00").strip()
            if result_code not in ("00", "000", "0000", ""):
                result_msg = root.findtext(".//resultMsg") or ""
                log.warning("API 오류 코드=%s 메시지=%s", result_code, result_msg)
                return [], 0

            total_count = int(root.findtext(".//totalCount") or "0")
            records = [
                {child.tag: (child.text or "").strip() for child in item}
                for item in root.findall(".//item")
            ]
            return records, total_count

        except requests.Timeout:
            log.warning("타임아웃 (시도 %d/3)", attempt)
        except requests.RequestException as e:
            log.warning("네트워크 오류 (시도 %d/3): %s", attempt, e)
        except ET.ParseError as e:
            log.warning("XML 파싱 오류: %s", e)
            return [], 0

        time.sleep(2 ** attempt)

    log.error("3회 재시도 실패, 해당 작업 건너뜀")
    return [], 0

# =============================================================
#  단위 수집
# =============================================================
def collect_one(conn, api_type, api_url, region_code, year_month, progress, log):
    page = 1
    total_saved = 0
    now = datetime.now().isoformat(timespec="seconds")
    by_type = progress["today_calls_by_type"]

    while True:
        if by_type.get(api_type, 0) >= DAILY_LIMIT:
            raise DailyLimitReached()

        records, total_count = fetch_page(api_url, region_code, year_month, page, log)
        by_type[api_type] = by_type.get(api_type, 0) + 1
        progress["today_calls"] += 1
        progress["total_calls"] += 1
        time.sleep(CALL_DELAY)

        if records:
            rows = [
                (api_type, region_code, year_month,
                 json.dumps(rec, ensure_ascii=False), now)
                for rec in records
            ]
            conn.executemany(
                "INSERT INTO transactions "
                "(api_type, region_code, year_month, data, collected_at) "
                "VALUES (?,?,?,?,?)",
                rows,
            )
            conn.commit()
            total_saved += len(records)

        if len(records) < ROWS_PER_PAGE or total_saved >= total_count:
            break
        page += 1

    progress["total_records"] += total_saved
    return total_saved

# =============================================================
#  연월 목록
# =============================================================
def generate_year_months(years_back=10):
    today = date.today()
    end_y, end_m = today.year, today.month
    start_y, start_m = 2016, 1

    months = []
    y, m = end_y, end_m
    while (y, m) >= (start_y, start_m):
        months.append("%d%02d" % (y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return months

# =============================================================
#  현황 출력
# =============================================================
def print_status(progress, total_tasks):
    done      = len(progress["completed"])
    pct       = done / total_tasks * 100 if total_tasks else 0
    remaining = total_tasks - done
    est_days  = (remaining // DAILY_LIMIT + 1) if remaining > 0 else 0
    print()
    print("=" * 55)
    print("  진행률     : %7d / %d (%.1f%%)" % (done, total_tasks, pct))
    print("  남은 작업  : %7d 건" % remaining)
    print("  예상 잔여  : 약 %d일 (일일 %d회 기준)" % (est_days, DAILY_LIMIT))
    print("  누적 호출  : %7d 회" % progress["total_calls"])
    print("  누적 레코드: %7d 건" % progress["total_records"])
    print("  오늘 날짜  : %s" % progress["today_date"])
    print("  오늘 호출  : %d / %d 회" % (progress["today_calls"], DAILY_LIMIT))
    print("=" * 55)
    print()

# =============================================================
#  메인 수집 루프
# =============================================================
def run():
    log = setup_logging()

    progress = load_progress()
    reset_daily_if_new_day(progress)

    year_months  = generate_year_months(YEARS_BACK)
    total_tasks  = len(API_TYPES) * len(REGIONS) * len(year_months)
    completed_set = set(progress["completed"])

    # 오늘 이미 소진된 호출 수 기준으로 잔여 한도 계산
    today_calls_at_start = progress["today_calls"]
    today_budget         = DAILY_LIMIT * len(API_TYPES) - today_calls_at_start

    session_start  = time.time()
    session_tasks  = 0
    session_calls  = 0
    records_before = progress["total_records"]

    log.info("=" * 60)
    log.info("국토교통부 실거래가 전수 수집 시작")
    log.info("전체 작업: %d건 | 완료: %d건 | 남음: %d건",
             total_tasks, len(completed_set), total_tasks - len(completed_set))
    log.info("오늘 호출: %d회 소진 / 한도 %d회(유형별 %d × %d종) | 잔여 %d회",
             today_calls_at_start, DAILY_LIMIT * len(API_TYPES),
             DAILY_LIMIT, len(API_TYPES), today_budget)
    log.info("=" * 60)

    conn = init_db()

    try:
        for api_type, (api_name, api_url) in API_TYPES.items():
            for region_code, region_name in REGIONS.items():
                for ym in year_months:
                    key = "%s|%s|%s" % (api_type, region_code, ym)
                    if key in completed_set:
                        continue

                    calls_before = progress["today_calls"]

                    saved = collect_one(
                        conn, api_type, api_url,
                        region_code, ym, progress, log,
                    )

                    calls_used    = progress["today_calls"] - calls_before
                    session_calls += calls_used

                    # 0건이면 전용 로그에만 기록 (정상 수집은 print 안 함)
                    if saved == 0:
                        log_zero_record(api_type, api_name, region_code, region_name, ym)

                    # 진행률 계산
                    completed_set.add(key)
                    done_total = len(completed_set)
                    total_pct  = done_total / total_tasks * 100

                    # 오늘 가용 한도 대비 소진 %
                    today_pct = (
                        session_calls / today_budget * 100
                        if today_budget > 0 else 100.0
                    )

                    # 예상 잔여 시간 (이번 세션 속도 기준)
                    elapsed    = time.time() - session_start
                    speed      = (session_tasks + 1) / elapsed if elapsed > 0 else 0
                    remaining  = total_tasks - done_total
                    eta_sec    = remaining / speed if speed > 0 else 0

                    # 콘솔: \r 덮어쓰기 (같은 줄 갱신)
                    status_line = (
                        f"[오늘 {today_pct:5.1f}% | 전체 {total_pct:5.1f}%]  "
                        f"소진 {session_calls}/{today_budget}회  "
                        f"경과 {fmt_duration(elapsed)}  "
                        f"잔여예상 {fmt_duration(eta_sec)}"
                    )
                    print(f"\r{status_line:<100}", end="", flush=True)



                    progress["completed"] = list(completed_set)
                    save_progress(progress)
                    session_tasks += 1

        print()
        log.info("★ 모든 수집 완료!")

    except DailyLimitReached:
        print()
        log.info("일일 호출 한도 %d회 도달 → 오늘 수집 종료", DAILY_LIMIT)
        log.info("내일 다시 실행하면 이어서 수집합니다.")

    except KeyboardInterrupt:
        print()
        log.info("사용자 중단 (Ctrl+C)")

    finally:
        conn.close()
        save_progress(progress)

        elapsed         = time.time() - session_start
        done            = len(completed_set)
        pct             = done / total_tasks * 100
        session_records = progress["total_records"] - records_before
        remaining       = total_tasks - done
        est_days        = (remaining // DAILY_LIMIT + 1) if remaining > 0 else 0

        log.info("-" * 60)
        log.info("세션 소요 시간 : %s", fmt_duration(elapsed))
        log.info("이번 세션      : 작업 %d건 | 저장 %d건 | 호출 %d회",
                 session_tasks, session_records, session_calls)
        log.info("누적 진행      : %d / %d (%.1f%%)", done, total_tasks, pct)
        if remaining > 0:
            log.info("예상 잔여      : 약 %d일 (일일 %d회 기준)", est_days, DAILY_LIMIT)
        log.info("=" * 60)

# =============================================================
#  엔트리포인트
# =============================================================
if __name__ == "__main__":
    if "--status" in sys.argv:
        _prog = load_progress()
        _ym   = generate_year_months(YEARS_BACK)
        _tot  = len(API_TYPES) * len(REGIONS) * len(_ym)
        print_status(_prog, _tot)
        sys.exit(0)

    if not MOLIT_SERVICE_KEY:
        print("오류: MOLIT_SERVICE_KEY 가 설정되지 않았습니다.")
        print("  방법 1) .env 파일에  MOLIT_SERVICE_KEY=발급받은키  추가")
        print("  방법 2) 이 파일 상단 MOLIT_SERVICE_KEY 변수에 직접 입력")
        sys.exit(1)

    run()
#!/usr/bin/env python3
"""
zero_records.log 점검 도구
===========================
사용법:
  python check_zero_records.py --analyze          # 패턴 분석만
  python check_zero_records.py --verify           # API 재조회 (샘플 10건)
  python check_zero_records.py --verify --all     # API 재조회 (전체)
  python check_zero_records.py --analyze --verify # 둘 다
"""

import re
import sys
import time
import argparse
import requests
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

# =============================================================
#  .env 로드 (bulk_collector.py 와 동일 방식)
# =============================================================
def _load_env() -> dict:
    env: dict = {}
    parent = Path(__file__).parent
    for folder in [parent, parent.parent]:
        p = folder / ".env"
        if p.exists():
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
            break
    return env

SERVICE_KEY = _load_env().get("MOLIT_SERVICE_KEY", "")

ZERO_LOG = Path(__file__).parent / "logs" / "zero_records.log"

API_URLS = {
    "apt_trade":  "https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev",
    "rh_trade":   "https://apis.data.go.kr/1613000/RTMSDataSvcRHTrade/getRTMSDataSvcRHTrade",
    "sh_trade":   "https://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade",
    "offi_trade": "https://apis.data.go.kr/1613000/RTMSDataSvcOffiTrade/getRTMSDataSvcOffiTrade",
    "apt_rent":   "https://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent",
    "rh_rent":    "https://apis.data.go.kr/1613000/RTMSDataSvcRHRent/getRTMSDataSvcRHRent",
    "sh_rent":    "https://apis.data.go.kr/1613000/RTMSDataSvcSHRent/getRTMSDataSvcSHRent",
    "offi_rent":  "https://apis.data.go.kr/1613000/RTMSDataSvcOffiRent/getRTMSDataSvcOffiRent",
}

# =============================================================
#  로그 파싱
# =============================================================
def load_zero_records():
    if not ZERO_LOG.exists():
        print(f"파일 없음: {ZERO_LOG}")
        sys.exit(1)
    records = []
    with open(ZERO_LOG, encoding="utf-8") as f:
        for line in f:
            parts = line.strip().split("\t")
            if len(parts) < 6:
                continue
            records.append({
                "ts":          parts[0],
                "api_type":    parts[1],
                "api_name":    parts[2],
                "region_code": parts[3],
                "region_name": parts[4],
                "year_month":  parts[5],
            })
    return records

# =============================================================
#  패턴 분석
# =============================================================
def analyze(records):
    print("\n" + "=" * 60)
    print(f"  zero_records.log 패턴 분석  (총 {len(records)}건)")
    print("=" * 60)

    # 거래 유형별
    by_type = Counter(r["api_name"] for r in records)
    print("\n[거래 유형별]")
    for name, cnt in by_type.most_common():
        bar = "█" * (cnt * 30 // max(by_type.values()))
        print(f"  {name:<18} {cnt:4d}건  {bar}")

    # 지역별 Top 10
    by_region = Counter(r["region_name"] for r in records)
    print("\n[지역별 Top 10]")
    for name, cnt in by_region.most_common(10):
        bar = "█" * (cnt * 30 // max(by_region.values()))
        print(f"  {name:<18} {cnt:4d}건  {bar}")

    # 연도별
    by_year = Counter(r["year_month"][:4] for r in records)
    print("\n[연도별]")
    for year in sorted(by_year):
        cnt = by_year[year]
        bar = "█" * (cnt * 30 // max(by_year.values()))
        print(f"  {year}년  {cnt:4d}건  {bar}")

    # 연월별 (최근 12개월)
    by_ym = Counter(r["year_month"] for r in records)
    recent = sorted(by_ym.keys(), reverse=True)[:12]
    print("\n[최근 12개월별]")
    for ym in recent:
        cnt = by_ym[ym]
        bar = "█" * (cnt * 30 // max(by_ym.values()))
        print(f"  {ym[:4]}-{ym[4:]}  {cnt:4d}건  {bar}")

    # 진단 의견
    print("\n[진단]")
    type_kinds = set(r["api_type"] for r in records)
    if len(type_kinds) == 1:
        print(f"  → 0건이 {list(type_kinds)[0]} 한 유형에만 집중 → 해당 유형의 데이터 자체가 희박하거나 API 엔드포인트 확인 필요")
    else:
        print(f"  → {len(type_kinds)}개 유형에 걸쳐 분포")

    current_ym = None
    from datetime import date
    d = date.today()
    current_ym = f"{d.year}{d.month:02d}"
    this_month = sum(1 for r in records if r["year_month"] == current_ym)
    if this_month > 0:
        print(f"  → 이번 달({current_ym}) {this_month}건: 아직 신고 기간이라 실제로 없을 수 있음 (정상)")

    print()

# =============================================================
#  API 재조회
# =============================================================
def verify(records, verify_all=False):
    if not SERVICE_KEY:
        print("오류: MOLIT_SERVICE_KEY 미설정 → .env 확인")
        sys.exit(1)

    targets = records if verify_all else records[:10]
    print(f"\n{'=' * 60}")
    print(f"  API 재조회  ({len(targets)}건{'  ← 전체' if verify_all else '  ← 샘플 10건'})")
    print("=" * 60)

    results = {"진짜_없음": 0, "API_오류": 0, "데이터_있음": 0}

    for r in targets:
        url = API_URLS.get(r["api_type"])
        if not url:
            continue

        params = {
            "serviceKey": SERVICE_KEY,
            "LAWD_CD":    r["region_code"],
            "DEAL_YMD":   r["year_month"],
            "numOfRows":  10,
            "pageNo":     1,
        }
        try:
            resp = requests.get(url, params=params, timeout=15)
            root = ET.fromstring(resp.content)
            result_code  = (root.findtext(".//resultCode") or "00").strip()
            result_msg   = (root.findtext(".//resultMsg") or "").strip()
            total_count  = int(root.findtext(".//totalCount") or "0")

            if result_code not in ("00", "000", "0000", ""):
                status = f"API 오류  코드={result_code} {result_msg}"
                results["API_오류"] += 1
            elif total_count == 0:
                status = "진짜 없음 (totalCount=0)"
                results["진짜_없음"] += 1
            else:
                status = f"데이터 있음! ({total_count}건) ← 수집 로직 확인 필요"
                results["데이터_있음"] += 1

        except Exception as e:
            status = f"요청 실패: {e}"
            results["API_오류"] += 1

        print(f"  {r['region_name']} {r['year_month'][:4]}-{r['year_month'][4:]}  {r['api_name']}")
        print(f"    └ {status}")
        time.sleep(0.3)

    print(f"\n[재조회 결과 요약]")
    print(f"  진짜 없음(정상) : {results['진짜_없음']}건")
    print(f"  API 오류        : {results['API_오류']}건")
    print(f"  데이터 있음(!)  : {results['데이터_있음']}건  ← 있다면 수집 코드 점검 필요")
    print()

# =============================================================
#  엔트리포인트
# =============================================================
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--analyze", action="store_true", help="패턴 분석")
    parser.add_argument("--verify",  action="store_true", help="API 재조회")
    parser.add_argument("--all",     action="store_true", help="재조회 전체 실행 (기본: 샘플 10건)")
    args = parser.parse_args()

    if not args.analyze and not args.verify:
        parser.print_help()
        sys.exit(0)

    records = load_zero_records()

    if args.analyze:
        analyze(records)

    if args.verify:
        verify(records, verify_all=args.all)
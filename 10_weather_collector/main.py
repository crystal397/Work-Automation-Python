"""
건설현장 기상데이터 수집 & 작업불가일 산정 통합 실행기
─────────────────────────────────────────────────────
실행: python main.py

흐름:
  1) 현장 정보 입력 (이름, 좌표)
  2) 관측소 선택 (자동추천 or 직접검색/지정)
  3) 수집 기간 설정
  4) 공종별 작업 기간·플래그 설정
  5) 데이터 수집 (기상청 API)
  6) 작업불가일 산정 & 엑셀 출력
"""

import sys
import re
from datetime import date, datetime

from station_mapper import ASOS_STATIONS, find_nearest_station, haversine
from kma_client import fetch_daily_weather, parse_weather, validate_station
from storage import init_db, upsert_weather
from analyzer import summarize
from flags import FLAG_DEFS, FLAG_BY_ID, WORK_PRESET_ITEMS

_OP_SYM = {">=": "≥", "<=": "≤", "<": "<", ">": ">"}


# ── 유틸리티 ─────────────────────────────────────────────────

def input_strip(prompt: str) -> str:
    return input(prompt).strip()


def validate_date(s: str) -> str | None:
    """YYYY-MM-DD 형식 검증. 유효하면 문자열 반환, 아니면 None"""
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return s
    except ValueError:
        return None


def input_date(prompt: str) -> str:
    while True:
        s = input_strip(prompt)
        if validate_date(s):
            return s
        print("  ※ YYYY-MM-DD 형식으로 입력해 주세요. (예: 2024-01-01)")


# ── 관측소 선택 ──────────────────────────────────────────────

def search_stations_by_name(keyword: str) -> list[dict]:
    """관측소 이름으로 검색 (부분 일치)"""
    keyword = keyword.strip()
    return [s for s in ASOS_STATIONS if keyword in s["name"]]


def show_station_list(stations: list[dict], lat=None, lon=None, max_show=20):
    """관측소 목록 표시 (좌표가 있으면 거리도 표시)"""
    for i, s in enumerate(stations[:max_show], 1):
        dist_str = ""
        if lat is not None and lon is not None:
            dist = haversine(lat, lon, s["lat"], s["lon"])
            dist_str = f"  ({dist:.1f}km)"
        print(f"  {i:3d}) [{s['code']:>4s}] {s['name']:<12s}"
              f"  ({s['lat']:.4f}, {s['lon']:.4f}){dist_str}")
    if len(stations) > max_show:
        print(f"  ... 외 {len(stations) - max_show}개")


def select_station(lat: float = None, lon: float = None) -> dict:
    """
    관측소 선택 인터페이스
    - 좌표가 있으면 자동추천 + 주변 5개 표시
    - 이름 검색, 코드 직접 입력 가능
    """
    print("\n" + "─" * 55)
    print("  관측소 선택")
    print("─" * 55)

    # 좌표 기반 추천
    if lat is not None and lon is not None:
        nearest = find_nearest_station(lat, lon)
        # 주변 5개
        ranked = sorted(
            ASOS_STATIONS,
            key=lambda s: haversine(lat, lon, s["lat"], s["lon"])
        )[:5]

        dist = haversine(lat, lon, nearest["lat"], nearest["lon"])
        print(f"\n  현장 좌표 ({lat}, {lon}) 기준 가장 가까운 관측소:")
        print(f"  → [{nearest['code']}] {nearest['name']} ({dist:.1f}km)\n")
        print("  주변 관측소 목록:")
        show_station_list(ranked, lat, lon)

    while True:
        print("\n  선택 방법:")
        if lat is not None and lon is not None:
            print("    Enter) 추천 관측소 사용")
            print("    번호)  위 목록에서 선택")
        print("    이름)  관측소 이름으로 검색 (예: 수원, 서울)")
        print("    코드)  관측소 코드 직접 입력 (예: 119)")

        choice = input_strip("\n  입력 > ")

        # 1) Enter → 추천 사용
        if choice == "" and lat is not None:
            print(f"  ✓ [{nearest['code']}] {nearest['name']} 선택")
            return nearest

        # 2) 번호 선택
        if choice.isdigit() and lat is not None:
            idx = int(choice) - 1
            if 0 <= idx < len(ranked):
                selected = ranked[idx]
                print(f"  ✓ [{selected['code']}] {selected['name']} 선택")
                return selected

        # 3) 이름 검색
        if not choice.isdigit() and choice != "":
            results = search_stations_by_name(choice)
            if not results:
                print(f"  ※ '{choice}' 검색 결과가 없습니다. 다시 시도해 주세요.")
                continue

            print(f"\n  '{choice}' 검색 결과 ({len(results)}건):")
            show_station_list(results, lat, lon)

            num = input_strip("  번호 선택 > ")
            if num.isdigit():
                idx = int(num) - 1
                if 0 <= idx < len(results):
                    selected = results[idx]
                    print(f"  ✓ [{selected['code']}] {selected['name']} 선택")
                    return selected

        # 4) 코드 직접 입력 체크
        code_match = [s for s in ASOS_STATIONS if s["code"] == choice]
        if code_match:
            selected = code_match[0]
            print(f"  ✓ [{selected['code']}] {selected['name']} 선택")
            return selected

        print("  ※ 올바른 값을 입력해 주세요.")


# ── 공종 설정 ────────────────────────────────────────────────

def select_works(site_start: str, site_end: str) -> list[dict]:
    """공종별 작업 기간·플래그 설정"""
    works = []

    print("\n" + "─" * 55)
    print("  공종별 작업 기간 설정")
    print("─" * 55)
    print(f"  수집 기간: {site_start} ~ {site_end}")

    while True:
        print(f"\n  ── 공종 {len(works) + 1} ──")
        print("  프리셋 선택:")
        for i, (name, flags) in enumerate(WORK_PRESET_ITEMS, 1):
            flags_str = ", ".join(flags[:3])
            if len(flags) > 3:
                flags_str += " ..."
            print(f"    {i}) {name}  [{flags_str}]")
        print(f"    0) 직접 입력")

        preset_choice = input_strip("\n  프리셋 번호 (Enter=직접입력) > ")

        work_thresholds: dict = {}

        preset_idx = (int(preset_choice) - 1
                      if preset_choice.isdigit()
                      and 1 <= int(preset_choice) <= len(WORK_PRESET_ITEMS)
                      else None)

        if preset_idx is not None:
            work_name, work_flags = WORK_PRESET_ITEMS[preset_idx]
            work_flags = work_flags[:]
            print(f"  → {work_name} 선택됨")

            # 플래그 수정 여부
            modify = input_strip("  플래그를 수정하시겠습니까? (y/N) > ")
            if modify.lower() == "y":
                work_flags, work_thresholds = select_flags(work_flags)
            else:
                # 플래그는 유지, 기준값만 확인
                print("\n  판정 기준값 설정 (Enter=기본값 유지):")
                for flag_id in work_flags:
                    if flag_id not in FLAG_BY_ID:
                        continue
                    _, label, col, op, default, unit = FLAG_BY_ID[flag_id]
                    if col is None:
                        continue
                    val_str = input_strip(
                        f"    {label:<10s} {_OP_SYM[op]} ? "
                        f"(기본 {default}{unit}, Enter=유지) > "
                    )
                    try:
                        work_thresholds[flag_id] = float(val_str)
                    except ValueError:
                        work_thresholds[flag_id] = default
        else:
            work_name = input_strip("  공종명 > ")
            if not work_name:
                work_name = f"공종{len(works) + 1}"
            work_flags, work_thresholds = select_flags()

        # 기간 입력
        print(f"\n  '{work_name}' 작업 기간:")
        ws_input = input_strip(f"    시작일 (YYYY-MM-DD, Enter={site_start}) > ")
        work_start = ws_input if validate_date(ws_input) else site_start

        we_input = input_strip(f"    종료일 (YYYY-MM-DD, Enter={site_end}) > ")
        work_end = we_input if validate_date(we_input) else site_end

        work = {
            "name":       work_name,
            "start":      work_start,
            "end":        work_end,
            "flags":      work_flags,
            "thresholds": work_thresholds,
        }
        works.append(work)

        # 기준값 요약 출력
        threshold_parts = []
        for flag_id in work_flags:
            if flag_id not in FLAG_BY_ID:
                continue
            _, label, col, op, default, unit = FLAG_BY_ID[flag_id]
            if col is None:
                threshold_parts.append(label)
            else:
                t = work_thresholds.get(flag_id, default)
                threshold_parts.append(f"{label}({_OP_SYM[op]}{t}{unit})")

        print(f"\n  ✓ [{work_name}] {work_start} ~ {work_end}")
        print(f"    기준: {', '.join(threshold_parts)}")

        more = input_strip("\n  공종을 추가하시겠습니까? (Y/n) > ")
        if more.lower() == "n":
            break

    return works


def select_flags(defaults: list[str] = None) -> tuple[list[str], dict]:
    """플래그 선택 + 수치 기준값 입력. (flags, thresholds) 반환"""
    print("\n  작업불가일 판정 플래그 선택:")
    for i, (flag_id, label, col, op, default, unit) in enumerate(FLAG_DEFS, 1):
        marker = " ✓" if defaults and flag_id in defaults else ""
        detail = f" ({_OP_SYM[op]}{default}{unit})" if col is not None else ""
        print(f"    {i:>2d}) {label:<14s}{detail:<18s} ({flag_id}){marker}")

    def _parse_nums(sel: str) -> list[str]:
        result = []
        for n in sel.split(","):
            n = n.strip()
            if n.isdigit() and 1 <= int(n) <= len(FLAG_DEFS):
                result.append(FLAG_DEFS[int(n) - 1][0])
        return result

    if defaults:
        print(f"\n  현재 선택: {', '.join(defaults)}")
        sel = input_strip("  번호를 쉼표로 입력 (Enter=현재 유지) > ")
        selected = defaults[:] if not sel else _parse_nums(sel)
    else:
        sel = input_strip("  번호를 쉼표로 입력 (예: 1,2,4,6) > ")
        selected = _parse_nums(sel)

    if not selected:
        print("  ※ 선택된 플래그가 없어 기본값(우천, 강풍)을 사용합니다.")
        selected = ["is_rain_day", "is_wind_day"]

    # ── 수치 플래그 기준값 설정 ──────────────────────
    print("\n  판정 기준값 설정 (Enter=기본값 유지):")
    thresholds: dict[str, float] = {}
    for flag_id in selected:
        if flag_id not in FLAG_BY_ID:
            continue
        _, label, col, op, default, unit = FLAG_BY_ID[flag_id]
        if col is None:
            continue  # 범주형 플래그는 수치 기준값 없음
        val_str = input_strip(
            f"    {label:<10s} {_OP_SYM[op]} ? "
            f"(기본 {default}{unit}, Enter=유지) > "
        )
        try:
            thresholds[flag_id] = float(val_str)
        except ValueError:
            thresholds[flag_id] = default

    return selected, thresholds


# ── 메인 흐름 ────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  건설현장 기상데이터 수집 & 작업불가일 산정")
    print("=" * 55)

    # ── 1. 현장 정보 ──
    print("\n[1/5] 현장 정보 입력")
    site_name = input_strip("  현장명 > ")
    if not site_name:
        site_name = "현장001"

    site_id = input_strip(f"  현장 ID (Enter={site_name[:8]}) > ")
    if not site_id:
        # 영문/숫자 기반 ID 생성
        site_id = re.sub(r"[^a-zA-Z0-9가-힣]", "", site_name)[:10] or "SITE001"

    lat_str = input_strip("  위도 (예: 37.2723, 없으면 Enter) > ")
    lon_str = input_strip("  경도 (예: 126.9853, 없으면 Enter) > ") if lat_str else ""

    lat, lon = None, None
    try:
        if lat_str:
            lat = float(lat_str)
        if lon_str:
            lon = float(lon_str)
    except ValueError:
        print("  ※ 좌표값이 올바르지 않아 좌표 없이 진행합니다.")
        lat, lon = None, None

    # ── 2. 관측소 선택 ──
    print("\n[2/5] 관측소 선택")
    while True:
        station = select_station(lat, lon)
        station_code = station["code"]
        print("  관측소 유효성 확인 중...", end=" ", flush=True)
        if validate_station(station_code):
            print("✓ ASOS 데이터 확인됨")
            break
        print("✗")
        print("  ※ 해당 관측소는 ASOS API에서 데이터를 제공하지 않습니다.")
        print("  다른 관측소를 선택해 주세요.\n")

    # ── 3. 수집 기간 ──
    print("\n[3/5] 수집 기간 설정")
    site_start = input_date("  시작일 (YYYY-MM-DD) > ")
    site_end = input_date("  종료일 (YYYY-MM-DD) > ")

    # ── 4. 공종 설정 ──
    print("\n[4/5] 공종 설정")
    use_works = input_strip("  공종별 작업불가일을 산정하시겠습니까? (Y/n) > ")

    works = []
    if use_works.lower() != "n":
        works = select_works(site_start, site_end)

    # ── 설정 확인 ──
    print("\n" + "=" * 55)
    print("  설정 확인")
    print("=" * 55)
    print(f"  현장명    : {site_name}")
    print(f"  현장 ID   : {site_id}")
    if lat and lon:
        print(f"  좌표      : ({lat}, {lon})")
    print(f"  관측소    : [{station_code}] {station['name']}")
    print(f"  수집 기간 : {site_start} ~ {site_end}")
    if works:
        print(f"  공종 수   : {len(works)}개")
        for w in works:
            print(f"    - {w['name']}: {w['start']} ~ {w['end']}")
            parts = []
            for flag_id in w["flags"]:
                if flag_id not in FLAG_BY_ID:
                    continue
                _, label, col, op, default, unit = FLAG_BY_ID[flag_id]
                if col is None:
                    parts.append(label)
                else:
                    t = w.get("thresholds", {}).get(flag_id, default)
                    parts.append(f"{label}({_OP_SYM[op]}{t}{unit})")
            print(f"      기준: {', '.join(parts)}")
    print("=" * 55)

    confirm = input_strip("\n  진행하시겠습니까? (Y/n) > ")
    if confirm.lower() == "n":
        print("  취소되었습니다.")
        sys.exit(0)

    # ── 5. 데이터 수집 ──
    print("\n[5/5] 데이터 수집 중...")
    init_db()

    start_fmt = site_start.replace("-", "")
    end_fmt = site_end.replace("-", "")

    print(f"  관측소 [{station_code}] {station['name']}에서 데이터 수집 중...")
    raw_records = fetch_daily_weather(station_code, start_fmt, end_fmt)

    if not raw_records:
        print("  ※ 수집된 데이터가 없습니다. API 키와 기간을 확인해 주세요.")
        sys.exit(1)

    parsed = [parse_weather(r, site_id) for r in raw_records]
    upsert_weather(parsed)
    print(f"  ✓ {len(parsed)}일치 데이터 수집 완료")

    # ── 6. 분석 & 엑셀 출력 ──
    if works:
        print("\n  작업불가일 산정 중...")
        site_config = {
            "id": site_id,
            "name": site_name,
            "lat": lat or 0,
            "lon": lon or 0,
            "start": site_start,
            "end": site_end,
            "works": works,
        }
        summarize(site_config)
    else:
        print("\n  ✓ 데이터 수집이 완료되었습니다.")
        print(f"    DB 위치: weather.db")
        print(f"    분석이 필요하면 config.py에 공종을 설정 후 analyzer.py를 실행하세요.")

    print("\n  모든 작업이 완료되었습니다.")


if __name__ == "__main__":
    main()
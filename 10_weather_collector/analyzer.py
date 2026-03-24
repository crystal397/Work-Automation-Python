import pandas as pd
from sqlalchemy import create_engine
from config import DB_PATH, SITES

engine = create_engine(f"sqlite:///{DB_PATH}")

# 컬럼 한글 매핑
COLUMN_LABELS = {
    "date":             "날짜",
    "station_code":     "관측소코드",
    "temp_max":         "최고기온(℃)",
    "temp_min":         "최저기온(℃)",
    "precipitation":    "일강수량(mm)",
    "wind_avg":         "평균풍속(m/s)",
    "wind_max":         "최대풍속(m/s)",
    "max_ins_wind":     "순간최대풍속(m/s)",
    "snow_depth":       "최대적설(cm)",
    "humidity_avg":     "평균습도(%)",
    "sunshine_hours":   "일조시간(hr)",
    "ground_temp":      "지면온도(℃)",
    "evaporation":      "증발량(mm)",
    "pressure":         "평균기압(hPa)",
    "is_rain_day":      "우천(10mm↑)",
    "is_wind_day":      "강풍(14m/s↑)",
    "is_wind_crane":    "크레인제한(순간10m/s↑)",
    "is_snow_day":      "적설(1cm↑)",
    "is_heat_day":      "폭염(35℃↑)",
    "is_cold_day":      "한파(-10℃↓)",
    "is_no_sunshine":   "일조부족(2hr미만)",
    "is_freeze_day":    "지면동결(0℃↓)",
    "is_high_evap_day": "증발과다(10mm↑)",
    "rain_yn":          "강수유무",
    "snow_yn":          "강설유무",
    "fog_yn":           "안개유무",
}

# 플래그 한글명
FLAG_NAMES = {
    "is_rain_day":      "우천 (10mm 이상)",
    "is_wind_day":      "강풍 (14m/s 이상)",
    "is_wind_crane":    "크레인 제한 (순간 10m/s 이상)",
    "is_snow_day":      "적설 (1cm 이상)",
    "is_heat_day":      "폭염 (35℃ 이상)",
    "is_cold_day":      "한파 (-10℃ 이하)",
    "is_no_sunshine":   "일조 부족 (2시간 미만)",
    "is_freeze_day":    "지면 동결 (0℃ 이하)",
    "is_high_evap_day": "증발 과다 (10mm 이상)",
    "rain_yn":          "강수 유무",
    "snow_yn":          "강설 유무",
    "fog_yn":           "안개",
}


def get_weather_df(site_id: str, start: str, end: str) -> pd.DataFrame:
    """기간 내 기상 데이터 조회"""
    query = """
        SELECT *
        FROM weather_daily
        WHERE site_id = :site_id
          AND date BETWEEN :start AND :end
        ORDER BY date
    """
    return pd.read_sql(query, engine, params={
        "site_id": site_id,
        "start":   start,
        "end":     end,
    })


def analyze_work(df: pd.DataFrame, work: dict) -> dict:
    """
    공종별 작업불가일 산정
    - work: {"name": ..., "start": ..., "end": ..., "flags": [...]}
    """
    # 공종 기간으로 필터
    mask = (df["date"] >= work["start"]) & (df["date"] <= work["end"])
    wdf  = df[mask].copy()

    if wdf.empty:
        return None

    flags = work["flags"]

    # 작업불가일: 해당 공종의 플래그 중 하나라도 True인 날
    wdf["is_work_impossible"] = wdf[flags].any(axis=1)

    total_days       = len(wdf)
    impossible_days  = int(wdf["is_work_impossible"].sum())
    workable_days    = total_days - impossible_days

    # 사유별 집계
    flag_counts = {
        flag: int(wdf[flag].sum())
        for flag in flags
        if flag in wdf.columns
    }

    return {
        "name":            work["name"],
        "start":           work["start"],
        "end":             work["end"],
        "total_days":      total_days,
        "workable_days":   workable_days,
        "impossible_days": impossible_days,
        "flag_counts":     flag_counts,
        "df":              wdf,
    }


def build_summary_sheet(site: dict, results: list) -> pd.DataFrame:
    """요약 시트 데이터 구성"""
    rows = []

    rows.append(("현장명",   site["name"]))
    rows.append(("수집기간", f"{site['start']} ~ {site['end']}"))
    rows.append(("",         ""))

    for work, r in zip(site.get("works", []), results):
        rows.append((f"[ {work['name']} ]", f"{work['start']} ~ {work['end']}"))

        if r is None:
            rows.append(("", "해당 기간 수집된 데이터가 없습니다."))
        else:
            rows.append(("총 일수",     f"{r['total_days']}일"))
            rows.append(("작업가능일",  f"{r['workable_days']}일"))
            rows.append(("작업불가일",  f"{r['impossible_days']}일"))
            rows.append(("",           "[ 사유별 집계 ]"))
            for flag, cnt in r["flag_counts"].items():
                label = FLAG_NAMES.get(flag, flag)
                rows.append((f"  {label}", f"{cnt}일"))

        rows.append(("", ""))

    rows.append(("※ 비고", "동일 날짜에 여러 사유가 겹쳐도 작업불가일은 1일로 산정"))

    return pd.DataFrame(rows, columns=["항목", "내용"])


def build_detail_sheet(df: pd.DataFrame, work_flags: list) -> pd.DataFrame:
    """일별 상세 시트 데이터 구성"""

    # 기본 기상 관측값 컬럼
    base_cols = [
        "date", "station_code",
        "temp_max", "temp_min", "precipitation",
        "wind_avg", "wind_max", "max_ins_wind",
        "snow_depth", "humidity_avg", "sunshine_hours",
        "ground_temp", "evaporation", "pressure",
    ]

    # 해당 공종의 플래그만
    flag_cols = [f for f in work_flags if f in df.columns]

    # 작업불가일 컬럼
    imp_col = ["is_work_impossible"] if "is_work_impossible" in df.columns else []

    # 중복 없이 컬럼 선택
    cols   = [c for c in base_cols if c in df.columns] + flag_cols + imp_col
    df_out = df[cols].copy()

    # Boolean 컬럼만 O/X 변환 (숫자 컬럼은 그대로 유지)
    for col in flag_cols + imp_col:
        df_out[col] = df_out[col].map(lambda x: "O" if x else "")

    # 컬럼명 한글 변환
    label_map = {**COLUMN_LABELS, "is_work_impossible": "작업불가일"}
    df_out = df_out.rename(columns=label_map)

    return df_out


def summarize(site: dict):
    site_id = site["id"]
    start   = site["start"]
    end     = site["end"]
    name    = site["name"]
    works   = site.get("works", [])

    # 전체 기간 기상 데이터 한 번에 조회
    df = get_weather_df(site_id, start, end)

    if df.empty:
        print(f"[{name}] 수집된 데이터가 없습니다.")
        return

    # 공종별 분석
    results = [analyze_work(df, work) for work in works]

    # 엑셀 저장
    output_path = f"{site_id}_작업불가일.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # 요약 시트
        df_summary = build_summary_sheet(site, results)
        df_summary.to_excel(writer, sheet_name="요약", index=False)

        # 공종별 상세 시트
        for work, r in zip(works, results):
            sheet_name = work["name"][:31]
            if r is None:
                pd.DataFrame(
                    [["해당 기간 수집된 데이터가 없습니다."]]
                ).to_excel(
                    writer, sheet_name=sheet_name, index=False, header=False
                )
            else:
                df_detail = build_detail_sheet(r["df"], work["flags"])
                df_detail.to_excel(writer, sheet_name=sheet_name, index=False)

    # 콘솔 출력
    print(f"\n{'='*55}")
    print(f"  {name}")
    print(f"  작업불가일 산정 결과 ({start} ~ {end})")
    print(f"{'='*55}")

    for work, r in zip(works, results):
        print(f"\n  [ {work['name']} ]  {work['start']} ~ {work['end']}")
        if r is None:
            print(f"    해당 기간 수집된 데이터가 없습니다.")
        else:
            print(f"    총 일수     : {r['total_days']}일")
            print(f"    작업가능일  : {r['workable_days']}일")
            print(f"    작업불가일  : {r['impossible_days']}일")
            print(f"    사유별 집계 :")
            for flag, cnt in r["flag_counts"].items():
                label = FLAG_NAMES.get(flag, flag)
                print(f"      {label:<30} : {cnt}일")

    print(f"\n  엑셀 저장 완료 → {output_path}")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    for site in SITES:
        summarize(site)
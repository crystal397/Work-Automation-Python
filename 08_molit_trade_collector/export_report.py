#!/usr/bin/env python3
"""
실거래가 DB → 경영진 보고용 Excel 추출기
=========================================
사용법:
  python export_report.py                        # 기본 (전체 최근 12개월)
  python export_report.py --region 강남구        # 특정 지역
  python export_report.py --type apt_trade       # 특정 거래유형
  python export_report.py --months 24            # 최근 24개월
  python export_report.py --min-price 50000      # 최소 거래금액 (만원)
  python export_report.py --max-price 200000     # 최대 거래금액 (만원)
"""

import argparse
import json
import sqlite3
import sys
from datetime import date
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import (
        Alignment, Border, Font, GradientFill, PatternFill, Side
    )
    from openpyxl.utils import get_column_letter
except ImportError:
    print("openpyxl 설치 필요: pip install openpyxl")
    sys.exit(1)

# =============================================================
#  설정
# =============================================================
DB_PATH = Path(__file__).parent / "bulk_data.db"

API_TYPE_LABELS = {
    "apt_trade":  "아파트 매매",
    "rh_trade":   "연립다세대 매매",
    "sh_trade":   "단독다가구 매매",
    "offi_trade": "오피스텔 매매",
    "apt_rent":   "아파트 전월세",
    "rh_rent":    "연립다세대 전월세",
    "sh_rent":    "단독다가구 전월세",
    "offi_rent":  "오피스텔 전월세",
}

REGION_CODES = {
    "종로구": "11110", "중구": "11140", "용산구": "11170",
    "성동구": "11200", "광진구": "11215", "동대문구": "11230",
    "중랑구": "11260", "성북구": "11290", "강북구": "11305",
    "도봉구": "11320", "노원구": "11350", "은평구": "11380",
    "서대문구": "11410", "마포구": "11440", "양천구": "11470",
    "강서구": "11500", "구로구": "11530", "금천구": "11545",
    "영등포구": "11560", "동작구": "11590", "관악구": "11620",
    "서초구": "11650", "강남구": "11680", "송파구": "11710",
    "강동구": "11740",
    "수원시": ["41110","41113","41115","41117"],
    "성남시": ["41130","41131","41135"],
    "의정부시": "41150", "안양시": ["41170","41171"],
    "부천시": "41190", "광명시": "41210", "평택시": "41220",
    "안산시": ["41270","41271"], "고양시": ["41280","41285","41287"],
    "과천시": "41290", "구리시": "41310", "남양주시": "41360",
    "오산시": "41370", "시흥시": "41390", "군포시": "41410",
    "의왕시": "41430", "하남시": "41450",
    "용인시": ["41460","41461","41463"],
    "파주시": "41480", "이천시": "41500", "안성시": "41550",
    "김포시": "41570", "화성시": "41590", "광주시": "41610",
    "양주시": "41630",
}

# =============================================================
#  컬럼 매핑 (API 필드 → 한글 컬럼명)
# =============================================================
TRADE_COLUMNS = [
    ("dealYear",      "거래연도"),
    ("dealMonth",     "거래월"),
    ("dealDay",       "거래일"),
    ("umdNm",         "법정동"),
    ("roadNm",        "도로명"),
    ("aptNm",         "건물명"),
    ("excluUseAr",    "전용면적(㎡)"),
    ("floor",         "층"),
    ("buildYear",     "건축연도"),
    ("dealAmount",    "거래금액(만원)"),
    ("dealingGbn",    "거래유형"),
    ("estateAgentSggNm", "중개사소재지"),
]

RENT_COLUMNS = [
    ("dealYear",      "거래연도"),
    ("dealMonth",     "거래월"),
    ("dealDay",       "거래일"),
    ("umdNm",         "법정동"),
    ("roadNm",        "도로명"),
    ("aptNm",         "건물명"),
    ("excluUseAr",    "전용면적(㎡)"),
    ("floor",         "층"),
    ("buildYear",     "건축연도"),
    ("deposit",       "보증금(만원)"),
    ("monthlyRent",   "월세(만원)"),
    ("contractTerm",  "계약기간"),
    ("contractType",  "계약구분"),
]

# =============================================================
#  스타일 헬퍼
# =============================================================
HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
SUBHEAD_FILL  = PatternFill("solid", fgColor="2E75B6")
STRIPE_FILL   = PatternFill("solid", fgColor="EBF3FB")
SUMMARY_FILL  = PatternFill("solid", fgColor="D6E4F0")
WHITE_FILL    = PatternFill("solid", fgColor="FFFFFF")

THIN_SIDE  = Side(style="thin",   color="BFBFBF")
MED_SIDE   = Side(style="medium", color="2E75B6")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE,
                     top=THIN_SIDE,  bottom=THIN_SIDE)
MED_BORDER  = Border(left=MED_SIDE,  right=MED_SIDE,
                     top=MED_SIDE,   bottom=MED_SIDE)

def h1(bold=True, size=12, color="FFFFFF"):
    return Font(name="Arial", bold=bold, size=size, color=color)

def h2(bold=True, size=10, color="FFFFFF"):
    return Font(name="Arial", bold=bold, size=size, color=color)

def body(size=9, color="000000", bold=False):
    return Font(name="Arial", size=size, color=color, bold=bold)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_mid():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def style_header_row(ws, row, col_start, col_end, text=None, fill=None):
    fill = fill or HEADER_FILL
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill   = fill
        cell.font   = h1()
        cell.border = THIN_BORDER
        cell.alignment = center()
    if text:
        ws.cell(row=row, column=col_start).value = text

# =============================================================
#  데이터 로드
# =============================================================
def load_data(api_types, region_codes, months, min_price, max_price):
    if not DB_PATH.exists():
        print(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)

    today = date.today()
    ym_list = []
    y, m = today.year, today.month
    for _ in range(months):
        ym_list.append("%d%02d" % (y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1

    placeholders_type   = ",".join("?" * len(api_types))
    placeholders_region = ",".join("?" * len(region_codes))
    placeholders_ym     = ",".join("?" * len(ym_list))

    query = f"""
        SELECT api_type, region_code, year_month, data
        FROM transactions
        WHERE api_type   IN ({placeholders_type})
          AND region_code IN ({placeholders_region})
          AND year_month  IN ({placeholders_ym})
        ORDER BY year_month DESC
    """
    params = api_types + region_codes + ym_list
    rows = conn.execute(query, params).fetchall()
    conn.close()

    records = []
    for api_type, region_code, year_month, data_str in rows:
        rec = json.loads(data_str)
        rec["_api_type"]    = api_type
        rec["_region_code"] = region_code
        rec["_year_month"]  = year_month

        # 가격 필터
        is_trade = "trade" in api_type
        price_str = rec.get("dealAmount", "").replace(",", "").strip() if is_trade else None
        if price_str:
            try:
                price = int(price_str)
                if min_price and price < min_price:
                    continue
                if max_price and price > max_price:
                    continue
            except ValueError:
                pass

        records.append(rec)

    return records

# =============================================================
#  시트 1 : 요약 대시보드
# =============================================================
def write_summary_sheet(wb, records, api_types, region_filter, months):
    ws = wb.active
    ws.title = "📊 요약"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.row_dimensions[1].height = 36

    today_str = date.today().strftime("%Y년 %m월 %d일")

    # ── 타이틀
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value     = f"부동산 실거래가 분석 보고서  |  기준일: {today_str}"
    title_cell.font      = h1(size=14)
    title_cell.fill      = HEADER_FILL
    title_cell.alignment = center()
    title_cell.border    = MED_BORDER

    # ── 조건 요약
    ws.merge_cells("A2:E2")
    cond_text = f"수집기간: 최근 {months}개월  /  대상유형: {', '.join(API_TYPE_LABELS[t] for t in api_types)}  /  지역: {region_filter}"
    c = ws["A2"]
    c.value = cond_text
    c.font  = body(size=9, color="404040")
    c.fill  = PatternFill("solid", fgColor="D9E2F3")
    c.alignment = left_mid()

    ws.row_dimensions[3].height = 10  # 여백

    # ── 섹션: 거래유형별 건수
    ws.merge_cells("A4:E4")
    ws["A4"].value     = "▌ 거래유형별 거래 현황"
    ws["A4"].font      = h2(size=10, color="1F4E79")
    ws["A4"].fill      = SUMMARY_FILL
    ws["A4"].alignment = left_mid()

    headers = ["거래유형", "총 건수", "평균금액(만원)", "최고금액(만원)", "최저금액(만원)"]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=ci)
        cell.value     = h
        cell.font      = h2()
        cell.fill      = SUBHEAD_FILL
        cell.alignment = center()
        cell.border    = THIN_BORDER

    type_stats: dict = {}
    for rec in records:
        at = rec["_api_type"]
        if at not in type_stats:
            type_stats[at] = {"count": 0, "prices": []}
        type_stats[at]["count"] += 1
        if "trade" in at:
            try:
                p = int(rec.get("dealAmount", "0").replace(",", ""))
                type_stats[at]["prices"].append(p)
            except ValueError:
                pass

    row = 6
    for at in api_types:
        stat = type_stats.get(at, {"count": 0, "prices": []})
        prices = stat["prices"]
        avg_p  = int(sum(prices) / len(prices)) if prices else "-"
        max_p  = max(prices) if prices else "-"
        min_p  = min(prices) if prices else "-"

        vals = [API_TYPE_LABELS[at], stat["count"],
                f"{avg_p:,}" if isinstance(avg_p, int) else avg_p,
                f"{max_p:,}" if isinstance(max_p, int) else max_p,
                f"{min_p:,}" if isinstance(min_p, int) else min_p]
        fill = STRIPE_FILL if row % 2 == 0 else WHITE_FILL
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=ci)
            cell.value     = v
            cell.font      = body()
            cell.fill      = fill
            cell.alignment = center() if ci > 1 else left_mid()
            cell.border    = THIN_BORDER
        row += 1

    ws.row_dimensions[row].height = 10
    row += 1

    # ── 섹션: 월별 거래량
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"].value     = "▌ 월별 거래량 추이"
    ws[f"A{row}"].font      = h2(size=10, color="1F4E79")
    ws[f"A{row}"].fill      = SUMMARY_FILL
    ws[f"A{row}"].alignment = left_mid()
    row += 1

    month_headers = ["연월", "총 건수", "매매 건수", "전월세 건수", "전월 대비"]
    for ci, h in enumerate(month_headers, 1):
        cell = ws.cell(row=row, column=ci)
        cell.value     = h
        cell.font      = h2()
        cell.fill      = SUBHEAD_FILL
        cell.alignment = center()
        cell.border    = THIN_BORDER
    row += 1

    month_data: dict = {}
    for rec in records:
        ym = rec["_year_month"]
        at = rec["_api_type"]
        if ym not in month_data:
            month_data[ym] = {"total": 0, "trade": 0, "rent": 0}
        month_data[ym]["total"] += 1
        if "trade" in at:
            month_data[ym]["trade"] += 1
        else:
            month_data[ym]["rent"] += 1

    sorted_months = sorted(month_data.keys(), reverse=True)[:12]
    first_data_row = row
    for i, ym in enumerate(sorted_months):
        d   = month_data[ym]
        ym_label = f"{ym[:4]}.{ym[4:]}"
        fill = STRIPE_FILL if row % 2 == 0 else WHITE_FILL

        ws.cell(row=row, column=1).value = ym_label
        ws.cell(row=row, column=2).value = d["total"]
        ws.cell(row=row, column=3).value = d["trade"]
        ws.cell(row=row, column=4).value = d["rent"]

        if i < len(sorted_months) - 1:
            prev_total = month_data[sorted_months[i + 1]]["total"]
            ws.cell(row=row, column=5).value = (
                f"=({get_column_letter(2)}{row}-{prev_total})/{prev_total}"
                if prev_total else "-"
            )
            ws.cell(row=row, column=5).number_format = "0.0%"
        else:
            ws.cell(row=row, column=5).value = "-"

        for ci in range(1, 6):
            cell = ws.cell(row=row, column=ci)
            cell.font      = body()
            cell.fill      = fill
            cell.alignment = center()
            cell.border    = THIN_BORDER
        row += 1

    # 합계 행
    ws.cell(row=row, column=1).value = "합 계"
    ws.cell(row=row, column=2).value = f"=SUM(B{first_data_row}:B{row-1})"
    ws.cell(row=row, column=3).value = f"=SUM(C{first_data_row}:C{row-1})"
    ws.cell(row=row, column=4).value = f"=SUM(D{first_data_row}:D{row-1})"
    ws.cell(row=row, column=5).value = "-"
    for ci in range(1, 6):
        cell = ws.cell(row=row, column=ci)
        cell.font      = body(bold=True)
        cell.fill      = SUMMARY_FILL
        cell.alignment = center()
        cell.border    = THIN_BORDER

# =============================================================
#  시트 2~ : 유형별 상세 데이터
# =============================================================
def write_detail_sheet(wb, records, api_type):
    label  = API_TYPE_LABELS[api_type]
    is_trade = "trade" in api_type
    cols   = TRADE_COLUMNS if is_trade else RENT_COLUMNS

    ws = wb.create_sheet(title=f"{'💰' if is_trade else '🏠'} {label}")
    ws.sheet_view.showGridLines = False

    # 헤더
    for ci, (_, col_label) in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci)
        cell.value     = col_label
        cell.font      = h2()
        cell.fill      = SUBHEAD_FILL
        cell.alignment = center()
        cell.border    = THIN_BORDER
        ws.column_dimensions[get_column_letter(ci)].width = 14

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 6
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 20
    ws.row_dimensions[1].height = 22

    type_records = [r for r in records if r["_api_type"] == api_type]

    for ri, rec in enumerate(type_records, 2):
        fill = STRIPE_FILL if ri % 2 == 0 else WHITE_FILL
        for ci, (field, _) in enumerate(cols, 1):
            val  = rec.get(field, "")
            cell = ws.cell(row=ri, column=ci)
            # 숫자형 변환 시도
            if field in ("excluUseAr", "floor", "buildYear",
                         "dealYear", "dealMonth", "dealDay",
                         "deposit", "monthlyRent"):
                try:
                    val = float(val) if "." in str(val) else int(val)
                except (ValueError, TypeError):
                    pass
            elif field == "dealAmount":
                try:
                    val = int(str(val).replace(",", ""))
                except (ValueError, TypeError):
                    pass
            cell.value     = val
            cell.font      = body()
            cell.fill      = fill
            cell.alignment = center()
            cell.border    = THIN_BORDER

    # 자동 필터
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
    ws.freeze_panes    = "A2"

    return len(type_records)

# =============================================================
#  메인
# =============================================================
def main():
    parser = argparse.ArgumentParser(description="실거래가 Excel 보고서 생성")
    parser.add_argument("--region",    nargs="+", default=[], help="지역 이름 (예: 강남구 서초구)")
    parser.add_argument("--type",      nargs="+",
                        default=list(API_TYPE_LABELS.keys()),
                        choices=list(API_TYPE_LABELS.keys()),
                        help="거래유형 키")
    parser.add_argument("--months",    type=int, default=12,  help="최근 N개월")
    parser.add_argument("--min-price", type=int, default=None, help="최소 거래금액 (만원, 매매만)")
    parser.add_argument("--max-price", type=int, default=None, help="최대 거래금액 (만원, 매매만)")
    parser.add_argument("--output",    default=None, help="출력 파일명")
    args = parser.parse_args()

    # 지역코드 변환
    if args.region:
        codes = []
        for r in args.region:
            v = REGION_CODES.get(r)
            if v is None:
                print(f"알 수 없는 지역: {r}")
                sys.exit(1)
            codes.extend(v if isinstance(v, list) else [v])
        region_filter = " / ".join(args.region)
    else:
        codes = list({v for vs in REGION_CODES.values()
                      for v in (vs if isinstance(vs, list) else [vs])})
        region_filter = "서울 + 수도권 전체"

    print(f"▶ 데이터 조회 중... (유형: {len(args.type)}종, 최근 {args.months}개월)")
    records = load_data(args.type, codes, args.months, args.min_price, args.max_price)
    print(f"  총 {len(records):,}건 로드됨")

    if not records:
        print("조건에 맞는 데이터가 없습니다.")
        sys.exit(0)

    # Excel 생성
    wb = openpyxl.Workbook()
    write_summary_sheet(wb, records, args.type, region_filter, args.months)

    for at in args.type:
        cnt = write_detail_sheet(wb, records, at)
        print(f"  └ {API_TYPE_LABELS[at]}: {cnt:,}건")

    # 저장
    today_str = date.today().strftime("%Y%m%d")
    out_path  = args.output or f"realestate_report_{today_str}.xlsx"
    wb.save(out_path)
    print(f"\n✅ 저장 완료: {out_path}")

if __name__ == "__main__":
    main()
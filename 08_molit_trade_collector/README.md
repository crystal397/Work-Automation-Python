# 국토교통부 실거래가 전수 수집기

국토교통부 실거래가 공공 API를 활용하여 **서울 25구 + 경기 43 시군구**의 부동산 **매매** 거래 데이터를 자동으로 수집하고, 연도별 Excel 보고서로 추출하는 Python 스크립트입니다.

---

## 개요

| 항목 | 내용 |
|------|------|
| 수집 대상 | 서울 25구 + 경기 43 시군구 (총 68개 지역) |
| 수집 기간 | 2016년 1월 ~ 현재 |
| 거래 유형 | 아파트·연립다세대·단독다가구·오피스텔 **매매** (4종) |
| API 출처 | 국토교통부 실거래가 공공데이터 (apis.data.go.kr) |
| 저장 방식 | SQLite (`bulk_data.db`) |
| 이어하기 | 지원 (중단 후 재실행 시 자동으로 이어서 수집) |

---

## 수집 대상 API (4종)

| 키 | 유형 |
|----|------|
| `apt_trade` | 아파트 매매 |
| `rh_trade` | 연립다세대 매매 |
| `sh_trade` | 단독다가구 매매 |
| `offi_trade` | 오피스텔 매매 |

---

## 파일 구조

```
08_molit_trade_collector/
├── bulk_collector.py        # 메인 수집 스크립트
├── export_report.py         # 연도별 Excel 보고서 추출
├── check_zero_records.py    # 0건 수집 항목 점검 도구
├── bulk_collector.bat       # Windows 스케줄러용 배치 파일
├── bulk_data.db             # 수집된 데이터 (SQLite, 자동 생성)
└── logs/
    ├── progress.json                   # 수집 진행 상황 (자동 생성)
    ├── bulk_collector_YYYYMMDD.log     # 일별 수집 로그 (자동 생성)
    ├── zero_records.log                # 0건 수집 항목 전용 로그 (자동 생성)
    └── cron.log                        # 배치 실행 로그 (자동 생성)
```

> `.env`는 현재 폴더에 없으면 상위 폴더(`python work/`)에서도 자동으로 탐색합니다.

---

## 사전 요구사항

- Python 3.8 이상
- 국토교통부 실거래가 API 키 ([공공데이터포털](https://www.data.go.kr) 에서 발급)
- 발급 후 아래 4개 API 별도 활용 신청 필요:
  - 아파트 매매 실거래가 자료
  - 연립다세대 매매 실거래 자료
  - 단독/다가구 매매 실거래 자료
  - 오피스텔 매매 신고 조회

```bash
pip install requests openpyxl
```

---

## API 키 설정

### 방법 1 — `.env` 파일 사용 (권장)

프로젝트 루트 또는 상위 폴더에 `.env` 파일을 만들고 아래 내용을 추가합니다.

```
MOLIT_SERVICE_KEY=발급받은_API_키
```

### 방법 2 — 스크립트 직접 수정

`bulk_collector.py` 상단의 변수를 직접 수정합니다.

```python
MOLIT_SERVICE_KEY = "여기에_API_키_입력"
```

---

## 사용법

### 1. 수집 (`bulk_collector.py`)

```bash
# 수집 실행 (중단 후 재실행하면 이어서 수집)
python bulk_collector.py

# 진행 현황만 확인
python bulk_collector.py --status
```

수집 중 터미널에서 한 줄이 실시간 갱신됩니다.

```
[오늘  42.3% | 전체  54.7%]  소진 423/39932회  경과 3분 31초  잔여예상 4시간 22분 18초
```

- **오늘 %**: 오늘 가용 한도(10,000 × 4종 = 40,000회) 대비 소진 비율
- **전체 %**: 전체 수집 작업 누적 완료 비율

`--status` 출력 예시:

```
=======================================================
  진행률     :   20000 / 34932 (57.3%)
  남은 작업  :   14932 건
  예상 잔여  : 약 2일 (일일 10000회 기준)
  누적 호출  :   20000 회
  누적 레코드:  600000 건
  오늘 날짜  : 2026-03-24
  오늘 호출  :  3000 / 10000 회
=======================================================
```

---

### 2. 보고서 추출 (`export_report.py`)

매매 데이터를 **연도별 시트**로 구성한 Excel 파일을 생성합니다.

```bash
# 기본 (전체 지역, 최근 10년)
python export_report.py

# 특정 지역만
python export_report.py --region 강남구 서초구

# 최근 N년
python export_report.py --years 5

# 거래금액 범위 필터 (만원)
python export_report.py --min-price 50000 --max-price 200000

# 출력 파일명 지정
python export_report.py --output my_report.xlsx
```

**생성되는 Excel 구조:**

| 시트 | 내용 |
|------|------|
| `요약` | 연도 × 거래유형 교차표 + 거래유형별 평균/최고/최저 금액 |
| `2016` ~ `2026` | 해당 연도의 모든 매매 거래 내역 (지역 → 거래유형 순 정렬) |

각 연도 시트 컬럼: 거래유형 / 지역 / 법정동 / 도로명 / 건물명 / 전용면적 / 층 / 건축연도 / 거래금액 / 거래월 / 거래일 / 거래유형상세 / 중개사소재지

---

### 3. 0건 점검 (`check_zero_records.py`)

수집 중 0건으로 기록된 항목의 원인을 분석합니다.

```bash
# 패턴 분석만 (거래유형별·지역별·연도별 분포)
python check_zero_records.py --analyze

# API 재조회로 실제 데이터 유무 확인 (샘플 10건)
python check_zero_records.py --verify

# API 재조회 전체 실행
python check_zero_records.py --verify --all

# 분석 + 재조회 동시 실행
python check_zero_records.py --analyze --verify
```

0건의 주요 원인:
- 해당 지역·연월에 실제로 거래가 없음 (정상)
- 이번 달처럼 아직 신고 기간 중인 경우 (정상)
- API 엔드포인트 오류 또는 키 권한 문제 → `--verify`로 확인

---

### 4. Windows 작업 스케줄러 등록

`bulk_collector.bat`을 스케줄러에 등록하면 매일 자동 수집됩니다.

```powershell
$action = New-ScheduledTaskAction `
  -Execute "C:\Users\CT Group\Desktop\python work\08_molit_trade_collector\.venv\Scripts\python.exe" `
  -Argument "bulk_collector.py" `
  -WorkingDirectory "C:\Users\CT Group\Desktop\python work\08_molit_trade_collector"

$trigger = New-ScheduledTaskTrigger -Daily -At "09:00"

Register-ScheduledTask -TaskName "molit_collector" -Action $action -Trigger $trigger
```

배치 실행 시 로그는 `logs/cron.log`에 쌓입니다.

---

## 주요 설정 값

`bulk_collector.py` 상단에서 조정할 수 있습니다.

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `DAILY_LIMIT` | `10000` | 하루 최대 API 호출 수 (유형당) |
| `CALL_DELAY` | `0.5` | API 호출 간격 (초) |
| `ROWS_PER_PAGE` | `1000` | 페이지당 최대 행 수 |
| `YEARS_BACK` | `10` | 수집 기간 (년, 2016년부터 고정) |

---

## 데이터베이스 구조

SQLite 파일(`bulk_data.db`) 내 `transactions` 테이블에 저장됩니다.

| 컬럼 | 타입 | 설명 |
|------|------|------|
| `id` | INTEGER | PK, 자동 증가 |
| `api_type` | TEXT | 거래 유형 (예: `apt_trade`) |
| `region_code` | TEXT | 법정동 코드 (예: `11680`) |
| `year_month` | TEXT | 거래 연월 (예: `202501`) |
| `data` | TEXT | API 응답 원문 (JSON) |
| `collected_at` | TEXT | 수집 일시 (ISO 8601) |

### 데이터 조회 예시

```python
import sqlite3, json

conn = sqlite3.connect("bulk_data.db")

# 강남구 아파트 매매 건수 (월별)
cur = conn.execute("""
    SELECT year_month, COUNT(*) AS cnt
    FROM transactions
    WHERE api_type = 'apt_trade' AND region_code = '11680'
    GROUP BY year_month
    ORDER BY year_month DESC
    LIMIT 12
""")
for row in cur:
    print(row)

conn.close()
```

---

## 로그 파일 안내

| 파일 | 설명 |
|------|------|
| `logs/bulk_collector_YYYYMMDD.log` | 세션 시작/종료 요약 + 경고/에러 기록 |
| `logs/zero_records.log` | 0건 수집 항목 전용 로그 (점검용) |
| `logs/cron.log` | 배치 파일 실행 시 stdout/stderr 전체 로그 |

로그 예시:
```
============================================================
국토교통부 실거래가 전수 수집 시작
전체 작업: 34932건 | 완료: 20000건 | 남음: 14932건
오늘 호출: 0회 소진 / 한도 40000회(유형별 10000 × 4종) | 잔여 40000회
============================================================

[WARNING] 네트워크 오류 (시도 1/3): ...
[ERROR]   3회 재시도 실패, 해당 작업 건너뜀

------------------------------------------------------------
세션 소요 시간 : 4시간 22분 18초
이번 세션      : 작업 8000건 | 저장 240000건 | 호출 8000회
누적 진행      : 28000 / 34932 (80.2%)
예상 잔여      : 약 1일 (일일 10000회 기준)
============================================================
```

---

## 동작 방식

```
실행
 │
 ├─ .env 로드 (현재 폴더 → 상위 폴더 순 탐색)
 │
 ├─ progress.json 로드 → 완료된 작업 목록 확인
 │
 ├─ [API 유형 4개] × [지역 68개] × [연월 ~122개] 순회
 │    │
 │    ├─ 이미 완료된 조합은 건너뜀
 │    │
 │    ├─ fetch_page() 호출 (페이지 단위, 최대 1,000건/페이지)
 │    │    └─ 실패 시 최대 3회 재시도 (지수 백오프: 2^n 초)
 │    │
 │    ├─ SQLite에 저장
 │    │
 │    ├─ 0건이면 zero_records.log에 기록
 │    │
 │    ├─ 콘솔 진행률 실시간 갱신
 │    │
 │    └─ progress.json 업데이트 (atomic replace)
 │
 └─ 일일 한도 도달 or 완료 → 종료 (소요 시간 출력)
```

- **이어하기**: 중단되어도 `progress.json`을 기반으로 완료 작업을 건너뛰고 재개
- **안전한 저장**: `progress.json`은 임시 파일에 쓴 뒤 교체(`os.replace`)하여 손상 방지
- **오류 복구**: 네트워크 오류/타임아웃 시 지수 백오프로 최대 3회 재시도

---

## 예상 소요 기간

전체 작업 수 = 4 유형 × 68 지역 × ~122 연월 ≈ **약 34,932 작업**

일일 한도 10,000회 기준으로 약 **3~4일** 소요됩니다.

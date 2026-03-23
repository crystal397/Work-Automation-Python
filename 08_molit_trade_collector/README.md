# 국토교통부 실거래가 전수 수집기

국토교통부 실거래가 공공 API를 활용하여 **서울 25구 + 경기 43 시군구**의 부동산 거래 데이터를 자동으로 수집하고, Excel 보고서로 추출하는 Python 스크립트입니다.

---

## 개요

| 항목 | 내용 |
|------|------|
| 수집 대상 | 서울 25구 + 경기 43 시군구 (총 68개 지역) |
| 수집 기간 | 2016년 1월 ~ 현재 |
| 거래 유형 | 아파트·연립다세대·단독다가구·오피스텔 × (매매 + 전월세) |
| API 출처 | 국토교통부 실거래가 공공데이터 (apis.data.go.kr) |
| 저장 방식 | SQLite (`bulk_data.db`) |
| 이어하기 | 지원 (중단 후 재실행 시 자동으로 이어서 수집) |

---

## 수집 대상 API (8종)

| 키 | 유형 |
|----|------|
| `apt_trade` | 아파트 매매 |
| `rh_trade` | 연립다세대 매매 |
| `sh_trade` | 단독다가구 매매 |
| `offi_trade` | 오피스텔 매매 |
| `apt_rent` | 아파트 전월세 |
| `rh_rent` | 연립다세대 전월세 |
| `sh_rent` | 단독다가구 전월세 |
| `offi_rent` | 오피스텔 전월세 |

---

## 사전 요구사항

- Python 3.8 이상
- 국토교통부 실거래가 API 키 ([공공데이터포털](https://www.data.go.kr) 에서 발급)

```bash
pip install requests openpyxl
```

---

## 설치 및 API 키 설정

### 방법 1 — `.env` 파일 사용 (권장)

프로젝트 루트 또는 **상위 폴더**에 `.env` 파일을 만들고 아래 내용을 추가합니다.  
(스크립트는 현재 폴더 → 상위 폴더 순으로 `.env`를 탐색합니다.)

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

### 수집 (`bulk_collector.py`)

```bash
# 수집 실행 (중단 후 재실행하면 이어서 수집)
python bulk_collector.py

# 진행 현황만 확인
python bulk_collector.py --status
```

### 보고서 추출 (`export_report.py`)

```bash
# 기본 (전체 지역, 최근 12개월)
python export_report.py

# 특정 지역
python export_report.py --region 강남구 서초구

# 특정 거래유형
python export_report.py --type apt_trade

# 최근 24개월
python export_report.py --months 24

# 거래금액 범위 필터 (만원, 매매만 적용)
python export_report.py --min-price 50000 --max-price 200000

# 출력 파일명 지정
python export_report.py --output my_report.xlsx
```

---

## 현황 출력 예시 (`--status`)

```
=======================================================
  진행률     :    1200 / 59840 (2.0%)
  남은 작업  :   58640 건
  예상 잔여  : 약 6일 (일일 10000회 기준)
  누적 호출  :    1210 회
  누적 레코드:  350000 건
  오늘 날짜  : 2025-06-01
  오늘 호출  : 1210 / 10000 회
=======================================================
```

### 수집 중 실시간 콘솔 표시

터미널에서 한 줄이 계속 덮어써지며 실시간으로 갱신됩니다.

```
[오늘  42.3% | 전체   1.8%]  소진 423/79349회  경과 3분 31초  잔여예상 4시간 22분 18초
```

- **오늘 %** : 오늘 가용 한도(잔여 호출 수) 대비 이번 세션 소진 비율 (전체 한도 = 10,000 × 8종 = 80,000회)
- **전체 %** : 전체 수집 작업 누적 완료 비율

---

## 주요 설정 값

`bulk_collector.py` 상단에서 아래 값을 조정할 수 있습니다.

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `DAILY_LIMIT` | `10000` | 하루 최대 API 호출 수 (100% 사용) |
| `CALL_DELAY` | `0.5` | API 호출 간격 (초) |
| `ROWS_PER_PAGE` | `1000` | 페이지당 최대 행 수 |
| `YEARS_BACK` | `10` | 수집 기간 (년) |

---

## 파일 구조

```
project/
├── bulk_collector.py   # 메인 수집 스크립트
├── export_report.py    # Excel 보고서 추출 스크립트
├── .env                # API 키 설정 (직접 생성)
├── bulk_data.db        # 수집된 데이터 (SQLite, 자동 생성)
└── logs/
    ├── progress.json                   # 수집 진행 상황 (자동 생성)
    ├── bulk_collector_YYYYMMDD.log     # 일별 수집 로그 (자동 생성)
    └── zero_records.log                # 0건 수집 항목 전용 로그 (자동 생성)
```

> `.env`는 현재 폴더에 없으면 상위 폴더에서도 자동으로 탐색합니다.

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

# 강남구 아파트 매매 건수 조회
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
| `logs/bulk_collector_YYYYMMDD.log` | 세션 시작/종료 요약 + 경고/에러만 기록 (진행률은 콘솔에만 표시) |
| `logs/zero_records.log` | 0건 수집된 항목만 별도 기록 — 코드 오류·API 이상 점검용 |

`bulk_collector_YYYYMMDD.log` 기록 예시:
```
# 시작
============================================================
국토교통부 실거래가 전수 수집 시작
전체 작업: 59840건 | 완료: 0건 | 남음: 59840건
오늘 호출: 0회 소진 / 한도 80000회(유형별 10000 × 8종) | 잔여 80000회
============================================================

# 경고/에러 발생 시만 기록
[WARNING] 타임아웃 (시도 1/3)

# 종료
------------------------------------------------------------
세션 소요 시간 : 4시간 22분 18초
이번 세션      : 작업 8000건 | 저장 240000건 | 호출 8000회
누적 진행      : 8000 / 59840 (13.4%)
예상 잔여      : 약 6일 (일일 10000회 기준)
============================================================
```

`zero_records.log` 형식 예시:
```
2026-03-23 09:15:02    offi_trade    오피스텔 매매    41800    경기 연천군    202301
```
0건이 많이 쌓인 경우, 해당 지역·유형·연월 조합의 API 응답을 직접 확인해보세요.

---

## 동작 방식

```
실행
 │
 ├─ .env 로드 (현재 폴더 → 상위 폴더 순 탐색)
 │
 ├─ progress.json 로드 → 완료된 작업 목록 확인
 │
 ├─ [API 유형 8개] × [지역 68개] × [연월 ~110개] 순회
 │    │
 │    ├─ 이미 완료된 조합은 건너뜀
 │    │
 │    ├─ fetch_page() 호출 (페이지 단위, 최대 1,000건/페이지)
 │    │    └─ 실패 시 최대 3회 재시도 (지수 백오프)
 │    │
 │    ├─ SQLite에 저장
 │    │
 │    ├─ 0건이면 zero_records.log에 기록
 │    │
 │    ├─ 진행률 로그 출력 (오늘 % / 전체 % / 경과 시간 / 잔여 예상)
 │    │
 │    └─ progress.json 업데이트 (atomic replace)
 │
 └─ 일일 한도 도달 or 완료 → 종료 (소요 시간 출력)
```

- **이어하기**: 수집 도중 중단되어도 `progress.json`을 기반으로 완료된 작업을 건너뛰고 재개합니다.
- **안전한 저장**: `progress.json`은 임시 파일에 쓴 뒤 교체하는 방식(`os.replace`)으로 손상을 방지합니다.
- **오류 복구**: 네트워크 오류 및 타임아웃 발생 시 지수 백오프(2^n초)로 최대 3회 재시도합니다.

---

## 예상 소요 기간

전체 작업 수 = 8 유형 × 68 지역 × ~110 연월 ≈ **약 59,840 작업**

일일 한도 10,000회 기준으로 약 **6~7일** 소요됩니다. (네트워크 및 API 응답 속도에 따라 달라질 수 있습니다.)

---

## crontab 등록을 위해 .bat 추가

해당 파일로 이동 -> 가상환경 사용

### Powershell 한 번에 등록

매일 오전 9시에 작업 시작

```powershell
$action = New-ScheduledTaskAction `
  -Execute "C:\Users\Desktop\08_molit_trade_collector\.venv\Scripts\python.exe" `
  -Argument "bulk_collector.py" `
  -WorkingDirectory "C:\Users\Desktop\08_molit_trade_collector"

$trigger = New-ScheduledTaskTrigger -Daily -At "09:00"

Register-ScheduledTask -TaskName "molit_collector" -Action $action -Trigger $trigger
```
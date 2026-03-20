# 국토교통부 실거래가 전수 수집기

국토교통부 실거래가 공공 API를 활용하여 **서울 25구 + 경기 43 시군구**의 부동산 거래 데이터를 자동으로 수집하는 Python 스크립트입니다.

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
pip install requests
```

---

## 설치 및 API 키 설정

### 방법 1 — `.env` 파일 사용 (권장)

프로젝트 루트에 `.env` 파일을 만들고 아래 내용을 추가합니다.

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

```bash
# 수집 실행 (중단 후 재실행하면 이어서 수집)
python bulk_collector.py

# 진행 현황만 확인
python bulk_collector.py --status
```

### 현황 출력 예시

```
=======================================================
  진행률     :    1200 / 98464 (1.2%)
  남은 작업  :   97264 건
  예상 잔여  : 약 11일 (일일 9000회 기준)
  누적 호출  :    1210 회
  누적 레코드:  350000 건
  오늘 날짜  : 2025-06-01
  오늘 호출  : 1210 / 9000 회
=======================================================
```

---

## 주요 설정 값

`bulk_collector.py` 상단에서 아래 값을 조정할 수 있습니다.

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `DAILY_LIMIT` | `9000` | 하루 최대 API 호출 수 (유형별 각각) |
| `CALL_DELAY` | `0.5` | API 호출 간격 (초) |
| `ROWS_PER_PAGE` | `1000` | 페이지당 최대 행 수 |
| `YEARS_BACK` | `10` | 수집 기간 (년) |

---

## 파일 구조

```
project/
├── bulk_collector.py   # 메인 수집 스크립트
├── .env                # API 키 설정 (직접 생성)
├── bulk_data.db        # 수집된 데이터 (SQLite, 자동 생성)
└── logs/
    ├── progress.json               # 수집 진행 상황 (자동 생성)
    └── bulk_collector_YYYYMMDD.log # 일별 로그 (자동 생성)
```

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

## 동작 방식

```
실행
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
 │    └─ progress.json 업데이트 (atomic replace)
 │
 └─ 일일 한도 도달 or 완료 → 종료
```

- **이어하기**: 수집 도중 중단되어도 `progress.json`을 기반으로 완료된 작업을 건너뛰고 재개합니다.
- **안전한 저장**: `progress.json`은 임시 파일에 쓴 뒤 교체하는 방식(`os.replace`)으로 손상을 방지합니다.
- **오류 복구**: 네트워크 오류 및 타임아웃 발생 시 지수 백오프(2^n초)로 최대 3회 재시도합니다.

---

## 예상 소요 기간

전체 작업 수 = 8 유형 × 68 지역 × ~110 연월 ≈ **약 59,840 작업**

일일 한도 9,000회 기준으로 약 **7~8일** 소요됩니다. (네트워크 및 API 응답 속도에 따라 달라질 수 있습니다.)

---

## 주의사항

- API 키는 `.env` 파일에 보관하고, `.gitignore`에 추가하여 저장소에 올라가지 않도록 주의하세요.
- 공공데이터포털의 일일 호출 제한 정책을 준수하여 `DAILY_LIMIT` 값을 설정하세요.
- `bulk_data.db`는 수집이 완료되면 수 GB에 달할 수 있습니다.

```gitignore
# .gitignore 권장 항목
.env
bulk_data.db
logs/
```

---

## 라이선스

MIT License

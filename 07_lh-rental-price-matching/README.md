# LH 임대주택 국토부 실거래가 매칭

LH 임대주택 195,238건에 대해 국토부 실거래가 API를 호출하여 각 주택의 매매 실거래가를 조회·매칭하는 Python 스크립트입니다.

---

## 주요 기능

- 행안부 도로명주소 API → Kakao API 폴백 방식으로 지번 자동 변환
- 국토부 실거래가 API 4종 (아파트·연립다세대·오피스텔·단독다가구) 자동 호출
- 주택유형별 매칭 전략 적용 및 G1~G4 4단계 등급 판정
- SQLite 기반 실거래 캐시 + JSON 기반 지번 캐시로 재실행 시 API 호출 최소화
- 지역 우선순위 배치 처리 (서울 → 수도권 → 광역시 → 기타)

---

## 매칭 결과 (2026-03-17 기준)

| 구분 | 건수 | 비율 |
|---|---|---|
| 전체 | 195,238건 | 100% |
| 매칭 성공 | 171,285건 | 87.7% |
| 미매칭 | 23,953건 | 12.3% |

### 등급별 분포

| 등급 | 기준 | 건수 | 비율 |
|---|---|---|---|
| G1 정밀 | 도로명/지번 + 층 + 면적 완전일치 | 59,994건 | 30.7% |
| G2 근접 | 지번 + 면적 또는 층 ±2 일치 | 11,638건 | 6.0% |
| G3 참고 | 법정동 + 층 + 면적 일치 | 22,505건 | 11.5% |
| G4 낮음 | 법정동만 일치 | 77,148건 | 39.5% |
| **G1+G2** | | **71,632건** | **36.7%** |

### 주택유형별 G1+G2 현황

| 유형 | 전체 | G1+G2 | 비율 |
|---|---|---|---|
| 아파트·주상복합 | 10,541건 | 8,859건 | 84.0% |
| 다세대·오피스텔·연립·도시형 | 98,818건 | 62,092건 | 62.8% |
| 다가구·단독 | 85,291건 | 512건 | 0.6% |
| **다가구·단독 제외 시** | **109,359건** | **70,951건** | **64.9%** |

> **다가구·단독 G1+G2 0.6%** : 국토부 개인정보보호 정책으로 SHTrade API의 지번이 `7*`, `3*` 형태로 마스킹되어 정밀 매칭이 구조적으로 불가능합니다.

---

## 설치 및 실행

### 요구사항

```bash
pip install pandas openpyxl requests tqdm
```

### .env 설정

프로젝트 루트에 `.env` 파일을 생성하고 아래 키를 입력합니다.

```
MOLIT_SERVICE_KEY=국토부실거래가API키
JUSO_API_KEY=행안부도로명주소API키
KAKAO_API_KEY=카카오API키
```

- `MOLIT_SERVICE_KEY` : [공공데이터포털](https://www.data.go.kr) 에서 발급
- `JUSO_API_KEY` : [주소정보 누리집](https://www.juso.go.kr) → 오픈API → 도로명주소 검색 API
- `KAKAO_API_KEY` : [Kakao Developers](https://developers.kakao.com) → 로컬 API

### 실행

```bash
python lh_realestate_api.py
```

입력 파일(`LH_임대주택공급현황_251120.xlsx`)이 동일 경로에 있어야 합니다.  
실행 시 `output/` 폴더가 자동 생성되며 결과 파일이 저장됩니다.

---

## 주요 설정값 (CONFIG)

```python
CONFIG = {
    "INPUT_FILE":        "LH_임대주택공급현황_251120.xlsx",
    "OUTPUT_DIR":        "output",
    "START_YMD":         "201603",   # 조회 시작일 (10년 치)
    "END_YMD":           "202602",   # 조회 종료일
    "AREA_TOLERANCE":    3.0,        # 전용면적 허용 오차 (㎡)
    "FLOOR_TOLERANCE":   2,          # 층 허용 오차
    "CALL_INTERVAL":     0.5,        # API 호출 간격 (초)
    "CACHE_TTL_DAYS":    30,         # 캐시 유효기간 (일)
    "MAX_WORKERS":       3,          # 동시 호출 스레드 수
    "DAILY_LIMIT":       9_497,      # API 일일 한도 (안전 마진 포함)
}
```

---

## 스크립트 구조

### 처리 흐름

```
1. 데이터 로드 및 전처리
   └─ LH 엑셀 파일 읽기 → 주택유형별 API 매핑 → 법정동코드 매핑 → 중복 제거

2. 지번 변환 (비아파트)
   └─ 주소 정제 (clean_address_for_jibun)
      → 행안부 도로명주소 API
      → 실패 시 Kakao API 폴백
      → jibun_cache.json 캐시

3. 실거래가 API 호출
   └─ 지역 우선순위 배치 (서울→수도권→광역시→기타)
      → ThreadPoolExecutor 병렬 호출
      → trade_cache.db (SQLite) 캐시
      → 일일 한도 도달 시 자동 중단

4. 주소 매칭 및 등급 판정
   └─ 아파트: 도로명 + 층 + 면적 → G1~G4
      비아파트: 지번 + 층 + 면적 → G1~G4
      다가구: 지번만 비교 (API 마스킹) → G1/G2/G4

5. 결과 저장
   └─ result_realestate_YYYYMMDD_HHMM.xlsx
      no_match_debug.xlsx (미매칭 디버그)
```

### 유형별 API 및 매칭 전략

| 주택유형 | 국토부 API | 매칭 필드 | 최대 등급 |
|---|---|---|---|
| 아파트·주상복합 | AptTradeDev | roadNm + floor + area | G1 |
| 다세대·연립·도시형 | RHTrade | jibun + floor + area | G1 |
| 오피스텔 | OffiTrade | ji + floor + area | G1 |
| 다가구·단독 | SHTrade | jibun만 (마스킹으로 사실상 G4) | G4 |

### 매칭 등급 판정 로직

```
[아파트]
  roadNm + floor_exact + area_ok → G1
  roadNm + floor_near  + area_ok → G2
  dong   + floor_near  + area_ok → G3
  dong만                         → G4

[다세대·오피스텔]
  jibun_exact + floor_exact + area_ok → G1
  jibun_exact + floor_near  + area_ok → G2
  jibun_exact + area_ok               → G2
  jibun_exact                         → G3
  jibun_bonbun + floor + area         → G2~G3
  else                                → G4

[다가구·단독]
  jibun_exact  → G1 (마스킹으로 사실상 도달 불가)
  jibun_bonbun → G2
  else         → G4
```

---

## 출력 파일

| 파일 | 설명 |
|---|---|
| `output/result_realestate_YYYYMMDD_HHMM.xlsx` | 전체 매칭 결과 |
| `output/no_match_debug.xlsx` | 미매칭 주소 디버그 |
| `output/trade_cache.db` | 실거래가 API 응답 캐시 (SQLite) |
| `output/jibun_cache.json` | 지번 변환 결과 캐시 |
| `output/api_query.log` | 실행 로그 |

---

## 알려진 한계

| 항목 | 원인 | 해결 가능 여부 |
|---|---|---|
| 다가구·단독 G4 집중 | 국토부 API 지번 마스킹 정책 | ❌ 불가 |
| 다세대 G3 잔존 | Kakao/행안부 지번 ≠ 국토부 등록 지번 (구조적 불일치) | ❌ 불가 |
| 오피스텔·다세대 미매칭 | 신규 LH 단지의 실거래 기록 부재 | ❌ 데이터 부재 |
| 건축물대장 API | Kakao 지번 ≠ 건축물대장 등록 지번 (4회 시도 모두 0건) | ❌ 코드 제거 완료 |

---

## 주요 변경 이력

| 버전 | 내용 |
|---|---|
| v4.0 | SQLite 캐시, Semaphore 원자적 처리, 배치 즉시 중단 등 |
| ① | 건축물대장 코드 완전 제거 → 실행시간 약 2시간 단축 |
| ② | 행안부 API 교체 (Kakao 단독 → 행안부 우선 + Kakao 폴백) |
| ③ | 조회기간 확대: 2023.01 → 2016.03 (38개월 → 120개월) |
| ④ | 오차 완화: 면적 ±1→3㎡, 층 ±1→±2 |
| ㉞ | `extract_dong_from_address()` 개선: 동명+건물명 혼합 주소 분리 |
| ㉟ | `clean_address_for_jibun()` 추가: 지번변환 전 주소 정제 (not_found 감소) |

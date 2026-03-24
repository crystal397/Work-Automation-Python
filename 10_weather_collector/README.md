# 10_weather_collector

기상청 ASOS(종관기상관측) API를 활용한 건설현장 기상데이터 Bulk 수집 및 작업불가일 산정 시스템

---

## 목적

건설 프로젝트 관리에서 기상 데이터는 다음 용도로 활용됩니다.

- **공정 관리** : 작업불가일(강수·강풍·적설·폭염·한파) 산정 및 공기 지연 근거 자료
- **계약·클레임 대응** : 불가항력 사유 발생 시 객관적 기상 기록 제출
- **비용 산정** : 기상 조건에 따른 간접비·장비 대기비 산정
- **안전 관리** : 극한 기상 시 작업 중지 기준 수립

---

## 프로젝트 구조

```
상위폴더/
├── .env                        ← API 키
└── 10_weather_collector/
    ├── README.md
    ├── config.py               ← 설정, API 키, 현장 목록 (공종별 작업 기간 포함)
    ├── station_mapper.py       ← 현장 좌표 → 기상관측소 매핑 (727개소)
    ├── kma_client.py           ← 기상청 API 호출 + 파싱
    ├── storage.py              ← DB 저장 (SQLite)
    ├── collector.py            ← 수집 오케스트레이터
    ├── scheduler.py            ← 매일 자동 실행 (APScheduler)
    └── analyzer.py             ← 공종별 작업불가일 산정 + 엑셀 출력
```

---

## 사전 준비

### 1. API 키 발급

[공공데이터포털](https://www.data.go.kr) 접속 후 아래 서비스 신청

- 서비스명 : `기상청_지상(종관, ASOS) 일자료 조회서비스`
- 승인 후 `일반 인증키(Encoding)` 복사

### 2. 패키지 설치

```bash
pip install requests pandas sqlalchemy apscheduler python-dotenv openpyxl
```

### 3. .env 파일 작성

프로젝트 **상위 폴더**에 `.env` 파일 생성

```
KMA_API_KEY=발급받은키를여기에입력
```

---

## 파일별 설명

### config.py

API 키, DB 경로, 수집 대상 현장 목록을 관리합니다.
각 현장에 **공종별 작업 기간(`works`)** 을 정의하면 `analyzer.py`에서 작업불가일을 자동으로 산정합니다.

```python
SITES = [
    {
        "id":    "SITE001",
        "name":  "수원 장안구 파장동~송죽동 현장",
        "lat":   37.2723,
        "lon":   126.9853,
        "start": "2024-01-01",
        "end":   "2025-12-31",

        # 공종별 작업 기간 정의
        "works": [
            {
                "name":  "토공사",
                "start": "2024-01-01",
                "end":   "2024-03-31",
                "flags": ["is_rain_day", "is_snow_day", "is_freeze_day", "is_cold_day"]
            },
            {
                "name":  "철근콘크리트공사",
                "start": "2024-04-01",
                "end":   "2024-09-30",
                "flags": ["is_rain_day", "is_heat_day", "is_cold_day",
                          "is_freeze_day", "is_wind_day"]
            },
            {
                "name":  "타워크레인작업",
                "start": "2024-04-01",
                "end":   "2024-12-31",
                "flags": ["is_wind_crane", "is_wind_day", "fog_yn"]
            },
            {
                "name":  "도장·방수공사",
                "start": "2025-01-01",
                "end":   "2025-06-30",
                "flags": ["is_rain_day", "rain_yn", "is_no_sunshine",
                          "is_cold_day", "is_freeze_day"]
            },
        ]
    },
]
```

현장 추가 시 `SITES` 리스트에 항목을 추가하면 됩니다.

---

### station_mapper.py

현장의 위도·경도를 기준으로 가장 가까운 ASOS 관측소를 자동으로 찾아줍니다.

- 기상청 공식 메타데이터 기준 **전국 727개소** 수록
- Haversine 공식으로 직선거리 계산

```python
from station_mapper import find_nearest_station

station = find_nearest_station(37.5172, 127.0473)
# {'code': '108', 'name': '서울', 'lat': 37.5714, 'lon': 126.9658}
```

---

### kma_client.py

기상청 ASOS 일자료 API를 호출하고 건설 관리에 필요한 항목을 파싱합니다.

수집 항목:

| 항목 | 설명 |
|---|---|
| `temp_max` / `temp_min` | 최고·최저 기온 (℃) |
| `precipitation` | 일강수량 (mm) |
| `wind_avg` / `wind_max` | 평균·최대 풍속 (m/s) |
| `max_ins_wind` | 순간최대풍속 (m/s) |
| `snow_depth` | 최대 적설 (cm) |
| `humidity_avg` | 평균 습도 (%) |
| `sunshine_hours` | 일조시간 (hr) |
| `ground_temp` | 지면온도 (℃) |
| `evaporation` | 증발량 (mm) |
| `pressure` | 평균기압 (hPa) |
| `is_rain_day` | 강수 10mm 이상 여부 |
| `is_wind_day` | 최대풍속 14m/s 이상 여부 |
| `is_wind_crane` | 순간최대풍속 10m/s 이상 여부 (크레인 작업 제한) |
| `is_snow_day` | 적설 1cm 이상 여부 |
| `is_heat_day` | 최고기온 35℃ 이상 여부 |
| `is_cold_day` | 최저기온 -10℃ 이하 여부 |
| `is_no_sunshine` | 일조시간 2시간 미만 여부 |
| `is_freeze_day` | 지면온도 0℃ 이하 여부 (지면 동결) |
| `is_high_evap_day` | 증발량 10mm 이상 여부 (증발 과다) |
| `rain_yn` | 강수 유무 (iscs 기반, 소량 포함) |
| `snow_yn` | 강설 유무 (iscs 기반, 눈·진눈깨비 포함) |
| `fog_yn` | 안개 유무 (지속시간 또는 iscs 기반) |

API 호출 제한(초당 1회)을 준수하며, 365일 단위로 청크 분할 요청합니다.

---

### storage.py

수집한 데이터를 DB에 저장합니다.

- `UNIQUE(site_id, date)` 제약으로 중복 저장 방지
- UPSERT 방식으로 재수집 시 최신 데이터로 갱신
- DB 파일은 최초 실행 시 `10_weather_collector/weather.db`로 자동 생성

테이블 구조:

```sql
CREATE TABLE weather_daily (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    site_id          TEXT NOT NULL,
    date             TEXT NOT NULL,
    station_code     TEXT,
    temp_max         REAL,
    temp_min         REAL,
    precipitation    REAL,
    wind_avg         REAL,
    wind_max         REAL,
    max_ins_wind     REAL,
    snow_depth       REAL,
    humidity_avg     REAL,
    sunshine_hours   REAL,
    ground_temp      REAL,
    evaporation      REAL,
    pressure         REAL,
    is_rain_day      BOOLEAN,
    is_wind_day      BOOLEAN,
    is_wind_crane    BOOLEAN,
    is_snow_day      BOOLEAN,
    is_heat_day      BOOLEAN,
    is_cold_day      BOOLEAN,
    is_no_sunshine   BOOLEAN,
    is_freeze_day    BOOLEAN,
    is_high_evap_day BOOLEAN,
    rain_yn          BOOLEAN,
    snow_yn          BOOLEAN,
    fog_yn           BOOLEAN,
    UNIQUE(site_id, date)
)
```

---

### collector.py

전체 수집 흐름을 조율하는 오케스트레이터입니다.

```
현장 목록 → 관측소 매핑 → API 호출 → 파싱 → DB 저장
```

직접 실행 시:

```bash
python collector.py
```

---

### scheduler.py

매일 오전 6시에 전날 데이터를 자동으로 갱신합니다.

```bash
python scheduler.py
```

---

### analyzer.py

DB에 저장된 기상 데이터를 바탕으로 **공종별 작업불가일을 산정**하고 엑셀 파일로 출력합니다.

- `config.py`의 `works` 정의를 읽어 공종별 기간·플래그 적용
- 동일 날짜에 여러 사유가 겹쳐도 작업불가일은 1일로 산정
- 출력 파일: `{site_id}_작업불가일.xlsx`

엑셀 구성:

| 시트 | 내용 |
|---|---|
| 요약 | 현장명, 공종별 총 일수 / 작업가능일 / 작업불가일, 사유별 집계 |
| 공종명 (개별 시트) | 일별 기상 관측값 + 플래그(O/X) + 작업불가일 여부 |

직접 실행 시:

```bash
python analyzer.py
```

---

## 실행 방법

### 최초 실행 (과거 데이터 bulk 수집)

```bash
python collector.py
```

`config.py`의 `SITES`에 설정된 `start` ~ `end` 기간 전체를 수집합니다.

### 작업불가일 산정 및 엑셀 출력

```bash
python analyzer.py
```

현장별 `{site_id}_작업불가일.xlsx` 파일이 생성됩니다.

### 일배치 실행 (매일 자동 갱신)

```bash
python scheduler.py
```

서버 환경에서는 백그라운드 실행을 권장합니다.

```bash
nohup python scheduler.py &
```

---

## 작업불가일 기준

| 구분 | 기준 | 플래그 | 비고 |
|---|---|---|---|
| 우천 | 일강수량 10mm 이상 | `is_rain_day` | |
| 강수 유무 | 소량 포함 강수 발생 | `rain_yn` | 지속성 강수 대응 |
| 강풍 | 최대풍속 14m/s 이상 | `is_wind_day` | |
| 크레인 제한 | 순간최대풍속 10m/s 이상 | `is_wind_crane` | 타워크레인 작업 기준 |
| 적설 | 최대적설 1cm 이상 | `is_snow_day` | |
| 강설 유무 | 눈·진눈깨비 발생 | `snow_yn` | iscs 기반 |
| 폭염 | 최고기온 35℃ 이상 | `is_heat_day` | |
| 한파 | 최저기온 -10℃ 이하 | `is_cold_day` | |
| 일조 부족 | 일조시간 2시간 미만 | `is_no_sunshine` | 도장·방수·양생 작업 기준 |
| 지면 동결 | 지면온도 0℃ 이하 | `is_freeze_day` | 동절기 터파기·다짐·콘크리트 타설 기준 |
| 증발 과다 | 증발량 10mm 이상 | `is_high_evap_day` | 콘크리트 양생 중 균열 방지 기준 |
| 안개 | 안개 발생 | `fog_yn` | 고층·크레인 작업 기준 |

기준값은 `kma_client.py`의 `parse_weather()` 함수에서 조정할 수 있습니다.

---

## 관측소 데이터 출처

- 출처 : 기상청 기상자료개방포털 (data.kma.go.kr)
- 파일 : `META_관측지점정보.csv`
- 기준 : 현재 운영 중인 관측소 (종료일 없음) / 국내 한정 (코드 < 1000)
- 관측소 수 : 727개소

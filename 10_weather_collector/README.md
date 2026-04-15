# 10_weather_collector

기상청 ASOS(종관기상관측) API를 활용한 건설현장 기상데이터 수집 및 작업불가일 산정 시스템

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
    ├── gui.py                  ← GUI 실행기 (customtkinter) ★ 주 실행 파일
    ├── main.py                 ← CLI 대화형 실행기
    ├── build.bat               ← exe 빌드 스크립트
    ├── .env.template           ← API 키 설정 템플릿
    ├── config.py               ← 설정, API 키, DB 경로 (exe/스크립트 모드 자동 전환)
    ├── station_mapper.py       ← 전국 관측소 목록 + 거리 계산
    ├── kma_client.py           ← 기상청 API 호출 + 파싱 + 관측소 유효성 검증
    ├── storage.py              ← DB 저장 (SQLite)
    ├── collector.py            ← 수집 오케스트레이터 (레거시 배치용)
    ├── scheduler.py            ← 매일 자동 실행 (APScheduler)
    └── analyzer.py             ← 공종별 작업불가일 산정 + 엑셀 출력
```

---

## 사전 준비

### 1. API 키 발급

[공공데이터포털](https://www.data.go.kr) 접속 후 아래 서비스 신청

- 서비스명 : `기상청_지상(종관, ASOS) 일자료 조회서비스`
- 승인 후 `일반 인증키(Decoding)` 복사

### 2. 패키지 설치

```bash
pip install requests pandas sqlalchemy python-dotenv openpyxl customtkinter
```

### 3. .env 파일 작성

프로젝트 **상위 폴더**에 `.env` 파일 생성

```
KMA_API_KEY=발급받은키를여기에입력
```

---

## 실행 방법

### GUI 실행 (권장)

```bash
python gui.py
```

단계별 화면으로 안내합니다.

| 단계 | 내용 |
|---|---|
| 1. 현장 정보 | 현장명, ID, 위도/경도 입력 |
| 2. 관측소 선택 | 이름 검색 또는 좌표 기반 추천 → ASOS 유효 관측소만 자동 필터링 |
| 3. 수집 기간 | 시작일 / 종료일 설정 |
| 4. 공종 설정 | 프리셋 선택 + 플래그 체크박스 + 기간 지정 |
| 5. 실행 | 설정 확인 → 진행바 + 실시간 로그 → 엑셀 자동 출력 |

### CLI 대화형 실행

```bash
python main.py
```

터미널 환경에서 질문에 답하며 순서대로 진행합니다.

### 배치 수집 (레거시)

```bash
python collector.py   # sites.json의 SITES 기준 일괄 수집
python analyzer.py    # 작업불가일 산정 + 엑셀 출력
python scheduler.py   # 매일 오전 6시 자동 갱신
```

---

## exe 배포

### 빌드

```
build.bat 더블클릭
```

빌드 결과물:

```
dist\기상데이터수집기\
  ├── 기상데이터수집기.exe   ← 실행 파일
  ├── .env                  ← API 키 (빌드 시 자동 복사)
  └── _internal\            ← 의존 라이브러리 (삭제 금지)
```

### 배포

`기상데이터수집기` 폴더 전체를 ZIP으로 압축하여 전달합니다.
받는 사람은 압축 해제 후 `기상데이터수집기.exe`를 실행합니다.

> `.env` 파일에 `KMA_API_KEY`가 반드시 포함되어 있어야 합니다.
> 실행 후 생성되는 `weather.db`와 `_작업불가일.xlsx`도 exe와 같은 폴더에 저장됩니다.

---

## 파일별 설명

### gui.py

customtkinter 기반 데스크톱 GUI입니다. 5단계 wizard 방식으로 구성되어 있습니다.

- 관측소 검색 시 ASOS API에 유효한 관측소만 병렬 자동 검증 후 표시
- 검색 결과가 없으면 인근 좌표 기준으로 자동 대체 검색
- 데이터 수집·관측소 유효성 확인은 백그라운드 스레드 처리 (UI 미 블로킹)

---

### config.py

API 키, DB 경로를 관리합니다. exe(frozen) 실행 시 자동으로 exe 폴더 기준 경로로 전환됩니다.

현장 목록은 `config.py`에 하드코딩하지 않고 `sites.json`에서 로드합니다. (`sites.example.json` 참고)

```python
# 스크립트 실행: 상위 폴더의 .env / 스크립트 폴더의 weather.db
# exe 실행:      exe 옆의 .env    / exe 폴더의 weather.db
```

---

### station_mapper.py

전국 관측소 목록과 거리 계산 유틸리티를 제공합니다.

> **참고**: 목록에는 ASOS·AWS·특수 관측소가 혼재합니다.
> GUI/CLI에서 관측소 선택 시 `validate_station()`으로 ASOS API 제공 여부를 자동 검증합니다.

```python
from station_mapper import find_nearest_station

station = find_nearest_station(37.5172, 127.0473)
# {'code': '108', 'name': '서울', 'lat': 37.5714, 'lon': 126.9658}
```

---

### kma_client.py

기상청 ASOS 일자료 API 호출, 파싱, 관측소 유효성 검증을 담당합니다.

**주요 함수:**

| 함수 | 설명 |
|---|---|
| `fetch_daily_weather(code, start, end)` | 일자료 수집 (365일 단위 분할 요청) |
| `parse_weather(raw, site_id)` | API 응답 → 건설 관리 항목 변환 |
| `validate_station(code)` | 관측소가 ASOS API에서 실제 데이터를 제공하는지 확인 |

**수집 항목:**

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
| `is_freeze_day` | 지면온도 0℃ 이하 여부 |
| `is_high_evap_day` | 증발량 10mm 이상 여부 |
| `rain_yn` | 강수 유무 (iscs 기반, 소량 포함) |
| `snow_yn` | 강설 유무 (iscs 기반, 눈·진눈깨비 포함) |
| `fog_yn` | 안개 유무 (지속시간 또는 iscs 기반) |

---

### storage.py

수집한 데이터를 SQLite DB에 저장합니다.

- `UNIQUE(site_id, date)` 제약으로 중복 저장 방지
- UPSERT 방식으로 재수집 시 최신 데이터로 갱신
- DB 파일 위치: 스크립트 실행 시 `10_weather_collector/weather.db`, exe 실행 시 exe 폴더

---

### analyzer.py

DB에 저장된 기상 데이터를 바탕으로 **공종별 작업불가일을 산정**하고 엑셀 파일로 출력합니다.

- 동일 날짜에 여러 사유가 겹쳐도 작업불가일은 1일로 산정
- 출력 파일: `{site_id}_작업불가일.xlsx` (exe/스크립트 실행 위치에 저장)

엑셀 구성:

| 시트 | 내용 |
|---|---|
| 요약 | 현장명, 공종별 총 일수 / 작업가능일 / 작업불가일, 사유별 집계 |
| 공종명 (개별 시트) | 일별 기상 관측값 + 플래그(O/X) + 작업불가일 여부 |

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
- 관측소 수 : 727개소 (ASOS·AWS·특수 관측소 포함)
- ASOS API 제공 여부는 실행 시 자동 검증

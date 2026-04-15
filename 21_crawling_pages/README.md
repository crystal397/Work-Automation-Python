# 21. 건설·부동산 관련 웹 크롤링 도구

건설협회·감정평가 관련 사이트 크롤링 및 공시지가 조회 도구 모음.

## 스크립트 목록

| 파일 | 기능 |
|------|------|
| `cak_crawler.py` | 대한건설협회(CAK) 건설업체 정보 수집 |
| `kosca_crawler.py` | 한국건설감정원(KOSCA) 데이터 수집 |
| `land_price_lookup.py` | V-World API 기반 공시지가 일괄 조회 |
| `csv_to_xlsx.py` | 수집 결과 CSV → Excel 변환 |
| `debug_api.py` | API 응답 디버깅 유틸리티 |

## 사용법

```bash
# 건설협회 업체 정보 수집
python cak_crawler.py

# 공시지가 조회 (토지 목록 Excel 필요)
python land_price_lookup.py

# CSV → Excel 변환
python csv_to_xlsx.py
```

## 환경 설정

`.env` 파일에 API 키 설정:

```env
VWORLD_API_KEY=브이월드_API_키
```

## 의존성

```bash
pip install requests openpyxl pandas python-dotenv
```

## 참고

- V-World API 키 발급: 공간정보 오픈플랫폼 (vworld.kr)
- 수집 결과 Excel 파일(`.xlsx`)은 `.gitignore`로 제외됨

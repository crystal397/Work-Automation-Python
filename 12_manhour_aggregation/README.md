# 12. 돌관공사비 노임 시트 공수 취합

업체별 일용노무비 자료를 취합하여 산출내역서 노임 시트에 자동 기입하는 도구.

## 주요 기능

- `pdf_to_excel.py`: PDF 노임 지급 명세서 → 페이지별 시트 Excel 변환 (병합 셀 복원)
- `aggregator.py`: 업체별 공수 집계
- `filler.py`: 노임 시트 자동 기입
- `formula_writer.py`: 합계 수식 자동 작성
- `readers/`: PDF·Excel 파일 형식별 읽기 모듈

## 사용법

```bash
# PDF → Excel 변환 (루트 폴더 지정)
python pdf_to_excel.py [폴더경로]

# .env 에 PDF_ROOT 설정 후 인자 생략 가능
python pdf_to_excel.py

# 공수 취합 및 시트 기입
python main.py
```

## 환경 설정

`.env.example`을 복사하여 `.env`로 저장 후 값 입력:

```env
PDF_ROOT=C:\작업폴더\노무비자료
```

## 의존성

```bash
pip install pdfplumber openpyxl python-dotenv
```

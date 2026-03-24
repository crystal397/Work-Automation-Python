# 세금계산서 PDF → 엑셀 자동 입력 도구

세금계산서 PDF 파일에서 핵심 정보를 자동 추출하여 엑셀 등록 양식에 일괄 입력하는 GUI 도구입니다.

---

## 주요 기능

- PDF 폴더 지정 또는 드래그앤드롭으로 일괄 처리
- **텍스트 추출** (pdfplumber) → 실패 시 **OCR 폴백** (Tesseract)
- 승인번호 5가지 패턴 자동 인식 (국세청 신고번호 / 파일명 / 일반 / 스캔형)
- 공급자·수취인 사업자번호, 작성일자, 공급가액 자동 추출
- 추출 결과를 엑셀 양식(`제3자발급사실 일괄등록양식`)에 자동 입력
- Tkinter GUI (드래그앤드롭 지원)

---

## 프로젝트 구조

```
09_pdf_invoice_extractor_and_excel_filler/
├── pdf_invoice_extractor_and_excel_filler.py   # 메인 실행 파일
└── README.md
```

---

## 설치

```bash
pip install pdfplumber pypdf openpyxl
```

### OCR 사용 시 추가 설치 (선택)

텍스트 추출이 불가한 스캔 PDF 처리에 필요합니다.

```bash
pip install pytesseract pdf2image pillow
```

- **Tesseract OCR** 별도 설치 필요: [https://github.com/tesseract-ocr/tesseract](https://github.com/tesseract-ocr/tesseract)
- 한국어 언어팩(`kor`) 추가 필요

---

## 실행

```bash
python pdf_invoice_extractor_and_excel_filler.py
```

1. GUI에서 **PDF 폴더**를 선택합니다.
2. 입력할 **엑셀 양식 파일**을 선택합니다.
3. 자동으로 추출 후 엑셀에 입력됩니다.

---

## 추출 항목

| 항목 | 설명 |
|---|---|
| `approval_no` | 승인번호 (국세청 신고번호 등 5가지 패턴) |
| `supplier_biz_no` | 공급자 사업자번호 |
| `receiver_biz_no` | 수취인 사업자번호 |
| `issue_date` | 작성일자 (YYYYMMDD) |
| `supply_amount` | 공급가액 |

---

## 주의사항

- OCR 미설치 시 스캔 PDF는 추출 불가 (텍스트 PDF만 처리)
- Tesseract 및 한국어 언어팩(`kor`) 미설치 시 OCR 자동 비활성화
- `pdf2image`는 내부적으로 Poppler를 사용하므로 별도 설치가 필요할 수 있음

---

## EXE 빌드

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "세금계산서추출기" pdf_invoice_extractor_and_excel_filler.py
```

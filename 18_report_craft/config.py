"""
시스템 설정값 — 경로, API, 모델 등
"""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env", override=False)
except ImportError:
    pass

# ── 기본 경로 ────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent

# 환경변수로 경로 재지정 가능. .env 또는 시스템 환경변수 모두 지원.
#   REPORT_INPUT_DIR  = 수신자료 폴더 경로 (기본: 프로젝트/input)
#   REPORT_OUTPUT_DIR = 결과물 저장 경로   (기본: 프로젝트/output)
INPUT_DIR     = Path(os.environ["REPORT_INPUT_DIR"])  if "REPORT_INPUT_DIR"  in os.environ else BASE_DIR / "input"
OUTPUT_DIR    = Path(os.environ["REPORT_OUTPUT_DIR"]) if "REPORT_OUTPUT_DIR" in os.environ else BASE_DIR / "output"

# classify --copy 결과 폴더 (extract --filtered 의 입력 경로)
# 설정하지 않으면 OUTPUT_DIR / "input_filtered" 를 자동 사용
FILTERED_DIR  = Path(os.environ["REPORT_FILTERED_DIR"]) if "REPORT_FILTERED_DIR" in os.environ else None

# 보고서 작성 회사명 (제출문에 표기)
# .env 또는 환경변수에서 REPORT_AUTHOR 로 설정
REPORT_AUTHOR = os.environ.get("REPORT_AUTHOR", "")

LAW_DIR       = BASE_DIR / "법령"
REFERENCE_DIR = BASE_DIR / "reference"
TEMPLATE_GUIDE = BASE_DIR / "report_template_guide.md"

# ── OCR ──────────────────────────────────────────────────────────────────────────
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSERACT_LANG = "kor+eng"

# ── 간접비 계산 상수 ──────────────────────────────────────────────────────────────
# 일반관리비 한도 (법령 기준)
GENERAL_ADMIN_RATE_LIMIT = {
    "A": 0.06,         # 지방계약법: 6%
    "B_large": 0.05,   # 국가계약법 300억 이상: 5%
    "B_small": 0.06,   # 국가계약법 300억 미만: 6%
    "C": 0.06,         # 민간: 6% (계약서 우선)
}
LARGE_CONTRACT_THRESHOLD = 30_000_000_000  # 300억

# 이윤 한도
PROFIT_RATE_LIMIT = 0.15   # 15%

# 퇴직급여충당금 기본 요율 (월급 / 12)
RETIREMENT_DEFAULT_RATE = 1 / 12

# 4대 보험 기본 요율 (직접계상 항목 — 사용자 부담분)
NATIONAL_PENSION_RATE   = 0.045     # 국민연금: 4.5%
HEALTH_INSURANCE_RATE   = 0.03545   # 건강보험: 3.545% (2025년 기준, 사용자 부담분)
LONG_TERM_CARE_RATE     = 0.1295    # 노인장기요양보험: 건강보험료 × 12.95% (2025년 기준)

# ── 품질 검증 임계값 ──────────────────────────────────────────────────────────────
MIN_CHARS_DEFAULT   = 30     # 최소 텍스트 길이
MIN_KOREAN_RATIO    = 0.05   # 한글 비율 최소 (OCR 대상)
MAX_GARBAGE_PATTERN = 3      # 깨진 문자 패턴 허용 건수
OCR_CONFIDENCE_WARN = 60     # OCR 신뢰도 경고 임계값 (%)
OCR_CONFIDENCE_FAIL = 45     # OCR 신뢰도 실패 임계값 (%)

# ── OCR 제한 ──────────────────────────────────────────────────────────────────
# 페이지 수가 이 값을 초과하면 OCR 시도 없이 텍스트 추출만 수행
# 스캔 PDF의 경우 FAIL로 표시되며 수동 확인 안내
PDF_OCR_MAX_PAGES   = 100    # 100페이지 초과 PDF는 OCR 생략
PDF_FAST_MAX_MB     = 20     # 20MB 초과 PDF는 pdfplumber 대신 pymupdf로만 빠르게 추출

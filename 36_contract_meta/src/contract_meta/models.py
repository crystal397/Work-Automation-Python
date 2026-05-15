"""계약 메타데이터의 단일 진실원.

모든 leaf 데이터는 `Sourced[T]` 로 감싸 출처 메타를 강제한다.
JSON Schema는 `ContractMeta.model_json_schema()` 로 자동 생성.
"""

from __future__ import annotations

from datetime import date
from enum import Enum
from typing import Generic, Literal, TypeVar

from pydantic import BaseModel, ConfigDict, Field, model_serializer

T = TypeVar("T")


# ─────────────────────────────────────────────────────────────────
# 출처(Source) 메타
# ─────────────────────────────────────────────────────────────────

class Source(BaseModel):
    """단일 데이터 추출 출처."""
    file: str = Field(description="원본 파일 경로(프로젝트 입력 디렉터리 기준 상대경로)")
    page: int | None = Field(default=None, description="PDF 페이지 번호(1-base). 엑셀이면 None")
    sheet: str | None = Field(default=None, description="엑셀 시트명")
    cell: str | None = Field(default=None, description="엑셀 셀 주소 (예: G42)")
    field_label: str | None = Field(default=None, description="원본 표/문서의 필드명(예: '공사명')")
    raw_text: str | None = Field(default=None, description="원본에서 추출한 원문 그대로")
    method: Literal["pdf_text", "ocr", "xlsx", "llm", "manual", "computed", "cross_check"] = Field(
        description="pdf_text=PDF 텍스트 레이어, ocr=OCR, xlsx=Excel 셀 추출, llm=LLM 보조, "
                    "manual=사용자 직접, computed=다른 필드에서 산출, "
                    "cross_check=의뢰처 요약본 등 보조 자료"
    )


class Sourced(BaseModel, Generic[T]):
    """값 + 출처 메타 묶음. 모든 leaf 데이터는 이 형태로 보관."""
    value: T
    src: Source = Field(alias="_source")

    model_config = ConfigDict(populate_by_name=True)


# ─────────────────────────────────────────────────────────────────
# 분기 enum (2장/3.2/4.2 템플릿 선택의 키)
# ─────────────────────────────────────────────────────────────────

class ClientType(str, Enum):
    NATIONAL = "국가"
    LOCAL = "지방"
    PRIVATE = "민간"


class ContractForm(str, Enum):
    LONG_TERM = "장기계속"
    CONTINUING_COST = "계속비"
    SINGLE_YEAR = "단년도"


class BiddingMethod(str, Enum):
    DETAILED = "내역입찰"
    TOTAL = "총액입찰"


class VATType(str, Enum):
    ZERO = "영세율"
    REGULAR = "일반과세"


# ─────────────────────────────────────────────────────────────────
# 도메인 모델
# ─────────────────────────────────────────────────────────────────

class Project(BaseModel):
    name: Sourced[str]
    name_short: Sourced[str] | None = None
    client_type: Sourced[ClientType]
    contract_form: Sourced[ContractForm]
    bidding_method: Sourced[BiddingMethod]
    vat_type: Sourced[VATType]
    bid_announcement_date: Sourced[date] | None = Field(
        default=None,
        description="입찰공고일 — 보고서 5.4 적용 법령 검색의 기준일. 22_laws_import 가 사용.",
    )


class Owner(BaseModel):
    """발주자."""
    name: Sourced[str]
    address: Sourced[str] | None = None
    representative: Sourced[str] | None = None
    contract_officer: Sourced[str] | None = Field(default=None, description="계약담당공무원/분임계약담당")


class Contractor(BaseModel):
    """계약상대자(원도급사 또는 하도급사)."""
    name_legal: Sourced[str] = Field(description="법인 정식명칭 (예: ㈜A건설)")
    name_display: Sourced[str] | None = Field(default=None, description="보고서 표기용 (예: ㈜A건설)")
    address: Sourced[str] | None = None
    representative: Sourced[str] | None = None
    business_no: Sourced[str] | None = None
    phone: Sourced[str] | None = None
    share_ratio_percent: Sourced[float] | None = None


class SupervisorMember(BaseModel):
    name: str
    share_percent: float | None = None


class Supervisor(BaseModel):
    """건설사업관리단(감리단)."""
    raw_text: Sourced[str]
    members: list[SupervisorMember] = Field(default_factory=list)
    responsible_person: Sourced[str] | None = Field(
        default=None, description="책임건설사업관리기술인 (감독자)"
    )


class Scope(BaseModel):
    """공사의 주요 내용."""
    location: Sourced[str] | None = None
    stations: Sourced[str] | None = None
    length_km: Sourced[float] | None = None
    items: list[Sourced[str]] = Field(default_factory=list)


class Money(BaseModel):
    """금액 (한글+숫자 동시 보관 → consistency.py 가 자동 검증)."""
    krw: Sourced[int]
    korean: Sourced[str] | None = None


class ContractRevision(BaseModel):
    """변경계약 1회분."""
    seq: int = Field(description="변경 회차 (0=최초, 1·2·...=변경). YAML에서 'no' 키는 boolean 충돌 우려로 사용 금지.")
    date: Sourced[date]
    contract_no: Sourced[str] | None = Field(default=None, description="제2023-1-146-202301-NN호")
    period_start: Sourced[date]
    period_end: Sourced[date]
    duration_days: Sourced[int]
    duration_diff_days: Sourced[int] = Field(description="직전 회차 대비 증감일")
    amount: Money
    amount_diff_krw: Sourced[int] | None = None
    reason: Sourced[str] | None = None


class ContractTerms(BaseModel):
    """변경계약서에 표기되는 표준 계약조건."""
    price_adjustment_method: Sourced[str] | None = Field(
        default=None, description="물가변동계약금액조정방법 — 지수조정율/품목조정율"
    )
    delay_penalty_rate: Sourced[str] | None = Field(default=None, description="지체상금율 — 예: '0.5/1000'")
    defect_warranty_rate: Sourced[str] | None = None


class ContractTrack(BaseModel):
    """총공사 계약 또는 차수별 계약 트랙."""
    initial: ContractRevision = Field(description="최초 계약 (seq=0)")
    revisions: list[ContractRevision] = Field(
        default_factory=list, description="변경 계약 회차 배열. revisions[i].seq = i+1"
    )


class SubsequentContract(BaseModel):
    """2차수·3차수 등 후속 차수 (참고용 메타)."""
    year_no: int
    initial_date: Sourced[date]
    period_start: Sourced[date]
    period_end: Sourced[date]
    amount_krw: Sourced[int]
    note: Sourced[str] | None = None


class Rates(BaseModel):
    """산출내역서에서 추출한 요율."""
    general_admin_percent: Sourced[float] | None = None
    profit_percent: Sourced[float] | None = None
    industrial_accident_insurance_percent: Sourced[float] | None = None
    employment_insurance_percent: Sourced[float] | None = None


class ExcludedPeriod(BaseModel):
    period_start: date
    period_end: date
    days: int
    reason: str


class CalculationTarget(BaseModel):
    """공기연장 간접비 산정 대상 기간 (보고서 4.1.1.2 표).

    사람의 판단(예: 3차수와 중복되는 4회 변경분 제외)이 반영되는 필드.
    method=manual 또는 computed.
    """
    period_start: Sourced[date]
    period_end: Sourced[date]
    days: Sourced[int]
    excluded_periods: list[ExcludedPeriod] = Field(default_factory=list)


class InputFile(BaseModel):
    role: str = Field(description="문서 역할 (변경계약서_1차_4회, 산출내역서, ...)")
    name: str | None = Field(default=None, description="_source.file 에서 참조하는 단축명 (없으면 path basename)")
    path: str
    sha256: str
    pages: int | None = None
    method: Literal["pdf_text", "ocr", "xlsx", "docx", "hwp"]


class Extraction(BaseModel):
    extractor_version: str
    extracted_at: str = Field(description="ISO8601")
    input_files: list[InputFile]
    warnings: list[str] = Field(default_factory=list)


# ─────────────────────────────────────────────────────────────────
# 최상위
# ─────────────────────────────────────────────────────────────────

class ContractMeta(BaseModel):
    """공기연장 보고서가 참조하는 단일 메타 객체."""
    schema_version: Literal["0.1.0"] = "0.1.0"
    project: Project
    owner: Owner
    contractor: Contractor
    co_contractors: list[Contractor] = Field(
        default_factory=list,
        description="공동이행 시 추가 계약상대자. contractor 가 대표사, co_contractors 가 나머지.",
    )
    supervisor: Supervisor | None = None
    scope: Scope | None = None
    contract_terms: ContractTerms | None = None
    total_contract: ContractTrack
    first_year_contract: ContractTrack
    subsequent_contracts: list[SubsequentContract] = Field(default_factory=list)
    subcontractors: list[Contractor] = Field(default_factory=list)
    rates: Rates | None = None
    calculation_target: CalculationTarget | None = None
    extraction: Extraction

    model_config = ConfigDict(populate_by_name=True)

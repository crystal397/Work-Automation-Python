"""계약서·변경계약서 (텍스트 PDF) 자동 메타 추출.

㈜A건설 A공구(도시철도) 양식 (지방계약법, 발주처(도시철도) 발주) 기준 디폴트 패턴.
다른 양식은 정규식 보강 또는 mapping yaml.

추출 필드
- 공사명, 공사관리번호
- 발주자 (수요기관), 계약법구분 (지방/국가)
- 계약상대자(들) — 공동이행이면 복수
- 계약금액 (총공사·차수), 낙찰율
- 착공일, 준공기한
- 변경계약: 변경계약일자, 변경 사유, 새 준공기한
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pdfplumber


_YYYYMMDD = r"\d{4}/\d{1,2}/\d{1,2}"


@dataclass
class ContractTextMeta:
    file: str
    project_name: str | None = None
    project_management_no: str | None = None
    contract_no: str | None = None
    owner: str | None = None
    contract_law: str | None = None              # 지방계약법 / 국가계약법
    bidding_method: str | None = None            # 총액입찰 / 내역입찰
    contractors: list[dict] = field(default_factory=list)
    total_amount_krw: int | None = None
    current_amount_krw: int | None = None
    bid_rate_percent: float | None = None
    start_date: date | None = None
    end_date_total: date | None = None
    end_date_current: date | None = None
    total_days: int | None = None
    current_days: int | None = None
    revision_date: date | None = None
    revision_reason: str | None = None
    end_date_current_before: date | None = None
    end_date_current_after: date | None = None
    delay_penalty_rate: str | None = None


def extract_meta(pdf_path: str | Path) -> ContractTextMeta:
    pdf_path = Path(pdf_path)
    meta = ContractTextMeta(file=str(pdf_path.name))
    with pdfplumber.open(pdf_path) as pdf:
        text = "".join((p.extract_text() or "") + "\n" for p in pdf.pages)

    lines = text.splitlines()
    # 공백 제거 텍스트 (한국어 PDF는 단어 사이 공백이 사라지는 경우 많음)
    nospace = re.sub(r"\s+", "", text)

    def _next_line_value(label_patterns: list[str]) -> str | None:
        """라벨이 단독으로 한 줄에 있고 값이 다음 줄에 있는 양식."""
        for i, line in enumerate(lines):
            s = re.sub(r"\s+", "", line).strip()
            for lab in label_patterns:
                if s == lab and i + 1 < len(lines):
                    return lines[i + 1].strip()
        return None

    def _same_line_value(label_patterns: list[str]) -> str | None:
        """라벨과 값이 같은 줄에 있는 양식 (예: '공 사 명 고속국도 ...')."""
        for line in lines:
            normalized = re.sub(r"\s+", "", line).strip()
            for lab in label_patterns:
                if normalized.startswith(lab):
                    # 원본 line 에서 라벨 뒤 부분
                    pattern = re.compile(r"\s*".join(lab) + r"\s+(.+)")
                    m = pattern.search(line)
                    if m:
                        return m.group(1).strip()
        return None

    # 공사관리번호 (예: 2302766-00 또는 230276600)
    m = re.search(r"공사관리번호[::\s]*(\d{7,12}(?:-\d+)?)", nospace)
    if m:
        meta.project_management_no = m.group(1)

    # 공사명 (도급계약서 양식: '공사명:A공구(도시철도)2호선...공사라.계약방법:...')
    m = re.search(r"공사명[::\s]*([가-힣\d\s\(\)\-]+?공사)(?:라|라\.|마\.)", nospace)
    if m:
        meta.project_name = m.group(1)
    # 공사명 (변경계약서 양식: '계약건명: ...공사')
    if not meta.project_name:
        m = re.search(r"계약건명[::\s]*([가-힣\d\(\)\-]+?공사)", nospace)
        if m:
            meta.project_name = m.group(1)
    # 공사명 (B공구 표 양식: '공 사 명\n<값>' 또는 '공 사 명 <값>')
    if not meta.project_name:
        v = _next_line_value(["공사명"]) or _same_line_value(["공사명"])
        if v:
            # 끝에 '계 약 금 액' 같은 다음 라벨이 붙어 있으면 제거
            v = re.split(r"\s+(?:계\s*약|착\s*공|준\s*공|총\s*공|공\s*고|발\s*주)", v)[0].strip()
            meta.project_name = v

    # 계약번호
    m = re.search(r"계약번호[::\s]*([\d\-]+)", nospace)
    if m:
        meta.contract_no = m.group(1)

    # 발주자 / 수요기관
    m = re.search(r"수요기관[::\s]*([가-힣\s]+?(?:본부|청|공단|공사|시청|군청|구청))", nospace)
    if m:
        meta.owner = m.group(1)
    # B공구 표 양식: '발 주 처\n발주처(고속도로)' 또는 '발 주 처 발주처(고속도로)'
    if not meta.owner:
        v = (_next_line_value(["발주처", "발주자", "발주기관"]) or
             _same_line_value(["발주처", "발주자", "발주기관"]))
        if v:
            v = re.split(r"\s+(?:ㆍ|계\s*약|공\s*고)", v)[0].strip()
            meta.owner = v

    # 계약법구분
    if "지방계약법" in nospace:
        meta.contract_law = "지방계약법"
    elif "국가계약법" in nospace:
        meta.contract_law = "국가계약법"

    # 입찰방식
    if "총액입찰" in nospace:
        meta.bidding_method = "총액입찰"
    elif "내역입찰" in nospace:
        meta.bidding_method = "내역입찰"

    # 계약금액 — '계약금액:금차8,143,000,000원(총공사166,168,310,000원)'
    m = re.search(r"금차([\d,]+)원[^\d]*총공사([\d,]+)원?", nospace)
    if m:
        meta.current_amount_krw = int(m.group(1).replace(",", ""))
        meta.total_amount_krw = int(m.group(2).replace(",", ""))
    # B공구 양식: '계약금액\n금 ... 원정(￦135,909,000 )' + '총공사부기금액\n금 ... 원정(￦244,...)'
    if meta.current_amount_krw is None:
        m = re.search(r"계약금액[^\\￦]{0,80}[\\￦]([\d,]+)", text)
        if m:
            meta.current_amount_krw = int(m.group(1).replace(",", ""))
    if meta.total_amount_krw is None:
        m = re.search(r"총공사부기금액[^\\￦]{0,80}[\\￦]([\d,]+)", text)
        if m:
            meta.total_amount_krw = int(m.group(1).replace(",", ""))
    # 총공사 기간 (세종): '총공사기간 : 1,350일'
    if meta.total_days is None:
        m = re.search(r"총공사기간[:\s]*([\d,]+)\s*일", text)
        if m:
            meta.total_days = int(m.group(1).replace(",", ""))

    # 낙찰율
    m = re.search(r"낙찰율[::\s]*([\d.]+)%", nospace)
    if m:
        meta.bid_rate_percent = float(m.group(1))

    # 착공일자
    m = re.search(rf"착공일자[::\s]*({_YYYYMMDD})", nospace)
    if m:
        meta.start_date = _parse_date_slash(m.group(1))
    # B공구 양식: '착 공 년 월 일\n2019-12-27'
    if meta.start_date is None:
        v = _next_line_value(["착공년월일", "착공일자"])
        if v:
            m = re.match(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", v)
            if m:
                meta.start_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if meta.end_date_current is None:
        v = _next_line_value(["준공년월일", "준공일자"])
        if v:
            m = re.match(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", v)
            if m:
                meta.end_date_current = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # 준공기한 (금차 형식: '금차2025/01/22(착공일로부터금차365일,총공사1734일)')
    m = re.search(rf"준공기한[::\s]*금차?({_YYYYMMDD})[^\d]*금차(\d+)일[^\d]*총공사(\d+)일", nospace)
    if m:
        meta.end_date_current = _parse_date_slash(m.group(1))
        meta.current_days = int(m.group(2))
        meta.total_days = int(m.group(3))

    # 변경계약 추가 패턴 (A공구(도시철도) 변경계약서 양식)
    # 계약금액 (변경계약서) — '계약금액:일금팔십일억사천삼백만원(\8,143,000,000)'
    if meta.current_amount_krw is None:
        m = re.search(r"계약금액[::\s]*[^\d]{0,80}\(?\\?([\d,]+)\)?", nospace)
        if m:
            try:
                v = int(m.group(1).replace(",", ""))
                if v > 1000000:  # 100만원 이상만
                    meta.current_amount_krw = v
            except ValueError:
                pass
    if meta.total_amount_krw is None:
        m = re.search(r"총부기계약금액[::\s]*[^\d]{0,80}\(?\\?([\d,]+)\)?", nospace)
        if m:
            try:
                v = int(m.group(1).replace(",", ""))
                if v > 1000000:
                    meta.total_amount_krw = v
            except ValueError:
                pass

    # 변경계약서 날짜들
    m = re.search(rf"착공일자[::\s]*({_YYYYMMDD})[^\d]*금차준공일자[::\s]*({_YYYYMMDD})[^\d]*총준공일자[::\s]*({_YYYYMMDD})", nospace)
    if m:
        if meta.start_date is None:
            meta.start_date = _parse_date_slash(m.group(1))
        meta.end_date_current = _parse_date_slash(m.group(2))
        meta.end_date_total = _parse_date_slash(m.group(3))

    # 변경 전/후 금차 준공기한
    m = re.search(rf"변경전금차준공기한[::\s]*({_YYYYMMDD})[^\d]*변경후금차준공기한[::\s]*({_YYYYMMDD})", nospace)
    if m:
        meta.end_date_current_before = _parse_date_slash(m.group(1))
        meta.end_date_current_after = _parse_date_slash(m.group(2))

    # 변경 사유
    m = re.search(r"준공기한변경사유([가-힣\(\)\s]+?)(?:\[|$|총부기)", nospace)
    if m:
        meta.revision_reason = m.group(1).strip()
    elif (m := re.search(r"변경사유([가-힣\(\)\s]+?)(?:변경계약증감액|준공기한)", nospace)):
        meta.revision_reason = m.group(1).strip()

    # 지체상금율
    m = re.search(r"지체상금[율률][::\s]*([\d.]+\s*%|[\d.]+/[\d]+)", nospace)
    if m:
        meta.delay_penalty_rate = m.group(1)

    # 변경계약일자
    m = re.search(rf"변경계약일자[::\s]*({_YYYYMMDD})", nospace)
    if m:
        meta.revision_date = _parse_date_slash(m.group(1))

    # 계약상대자 (들) — 공동이행 인식
    # 패턴: '상호 ㈜A건설 ... 지분율 55%'
    for m in re.finditer(
        r"상호([가-힣A-Za-z\(\)주식회사\s]+?(?:주식회사|건설|건설업|\(주\))).*?사업자등록번호([\d-]+).*?대표자([가-힣]{2,5}).*?지분율(\d{1,3})\s*%",
        nospace,
    ):
        meta.contractors.append({
            "name_legal": m.group(1).strip(),
            "business_no": m.group(2),
            "representative": m.group(3),
            "share_ratio_percent": int(m.group(4)),
        })

    # 단독 계약 (지분율 100%)
    if not meta.contractors:
        m = re.search(r"계약상대자[::\s]*([가-힣A-Za-z주식회사\(\)\s]+?주식회사)\(대표자[::\s]*([가-힣]{2,5})\)", nospace)
        if m:
            meta.contractors.append({
                "name_legal": m.group(1).strip(),
                "representative": m.group(2),
                "share_ratio_percent": 100,
            })

    # B공구 양식: 같은 줄에 'ㆍ상 호 ㈜B건설 주식회사' / 'ㆍ대표자 대표자_A ㆍ전 화 번 호 ...'
    if not meta.contractors:
        contractor_blocks: list[dict] = []
        current: dict = {}
        for line in lines:
            # 'ㆍ상 호 <회사>' 또는 'ㆍ상호 <회사>'
            m = re.search(r"ㆍ\s*상\s*호\s+([가-힣A-Za-z\(\)\s주식회사]+?)(?:\s+ㆍ|$)", line)
            if m:
                name = m.group(1).strip()
                if name and ("주식회사" in name or name.startswith("(주)") or name.endswith("(주)")):
                    if current.get("name_legal"):
                        contractor_blocks.append(current)
                    current = {"name_legal": name}
            m = re.search(r"ㆍ\s*대\s*표\s*자\s+([가-힣]{2,5})(?:\s+ㆍ|$)", line)
            if m and current.get("name_legal") and "representative" not in current:
                current["representative"] = m.group(1)
        if current.get("name_legal"):
            contractor_blocks.append(current)
        # 연대보증인 블록은 제외 (회사명에 '연대보증인' 들어가지 않음 — 단 빈 데이터 행 가능)
        contractor_blocks = [b for b in contractor_blocks if b.get("name_legal") and "연대" not in b["name_legal"]]
        if contractor_blocks:
            n = len(contractor_blocks)
            share = round(100 / n)
            for b in contractor_blocks:
                b["share_ratio_percent"] = share
            meta.contractors = contractor_blocks

    return meta


def _parse_date_slash(s: str) -> date:
    y, m, d = s.split("/")
    return date(int(y), int(m), int(d))

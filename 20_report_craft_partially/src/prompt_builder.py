"""
claude.ai 에 붙여넣을 귀책분석 작성 프롬프트를 생성.

입력:
  output/reference_patterns.md  — learn 단계 결과
  output/correspondence_texts.md — scan 단계 결과 (전문)
  output/scan_result.json        — 공문 목록 (편집된 것)

출력:
  output/prompt_for_claude.md   — claude.ai에 그대로 복사·붙여넣기
  output/귀책분석_schema.json   — generate 단계 JSON 스키마 (claude.ai가 채울 형식)
"""

from __future__ import annotations

import json
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))
import config


SCHEMA_GUIDE = """
## [출력 형식 — JSON]

아래 스키마를 **그대로** 사용하여 귀책분석 데이터를 작성하세요.

---

### ❌ 절대 금지 — 이 필드명은 사용하면 보고서에 아무것도 출력되지 않습니다

| 잘못된 필드명 (사용 금지) | 올바른 필드명 |
|---|---|
| `intro_paragraph` | `background_paragraphs` (리스트) |
| `contract_info` | (사용하지 않음 — 내용은 background_paragraphs에 포함) |
| `chapter_title`, `section_title` | (사용하지 않음) |
| `detail_narratives[].body` | `detail_narratives[].paragraphs` (리스트) |
| `detail_narratives[].section_no` | (사용하지 않음) |
| `items[].summary` | `items[].causal_description` |
| `accountability_diagram[].delay_cause` | `accountability_diagram[].cause` (또는 `delay_cause` 둘 다 허용) |

**위 잘못된 필드명을 사용하면 해당 섹션 전체가 공백으로 생성됩니다. 반드시 올바른 필드명을 사용하세요.**

---

### ✅ 필수 필드 체크리스트 (전부 작성해야 함)

출력 JSON을 저장하기 전에 아래 항목을 모두 확인하세요:

- [ ] `project_name` — 프로젝트명 (문자열)
- [ ] `total_delay_days` — 변경계약서 연장일수 합산 정수 (0이면 안 됨)
- [ ] `background_paragraphs` — 2단락 이상의 리스트 (빈 [] 금지)
- [ ] `items` — 공문 목록 (show_in_table: true/false 포함)
- [ ] `detail_narratives` — 각 블록은 반드시 `title`(또는 `label`)과 `paragraphs`(리스트) 포함
- [ ] `detail_narratives` 각 블록 마지막 단락 — **소결 3단계 포함 여부 확인** (①원인 ②귀책조항 ③"계약상대자의 책임 없는 사유")
- [ ] `pre_diagram_paragraphs` — 2단락 이상의 리스트 (빈 [] 금지)
- [ ] `accountability_diagram` — 귀책사유 도식표 항목 리스트
- [ ] `accountability_diagram` 각 항목 — **`delay_days` 필드(정수) 반드시 포함**, 일수 미확정 시 `0` 사용
- [ ] `accountability_diagram` 항목 `delay_days` 합계 = `total_delay_days` — **합계 행 별도 추가 금지** (자동 합산하므로 합계 행 추가 시 2배로 계산되어 오류 발생)
- [ ] `accountability_diagram` 각 항목 — **`source_docs` 필드 반드시 포함** (근거 공문번호 또는 변경계약 차수 리스트, 예: `["제22-0123호", "3차 변경계약"]`)
- [ ] `conclusion_paragraphs` — **반드시 2단락 이상** (빈 [] 절대 금지)
- [ ] `summary` — 종합 요약 문자열 (빈 "" 금지)

---

### JSON 스키마 (이 구조 그대로 채울 것)

```json
{
  "project_name": "프로젝트명",

  "total_delay_days": 123,

  "background_paragraphs": [
    "단락 1: 사업 배경, 최초 계약일, 계약금액, 발주처·시공사·감리 등 기본 정보",
    "단락 2: 공기연장 경위 개요 — 몇 차례 변경계약, 최종 준공일, 귀책분석 대상 기간"
  ],

  "items": [
    {
      "no": 1,
      "scan_no": 1,
      "source_file": "원본 파일명 (확정 공문 목록 표의 source_file 값을 그대로 복사)",
      "date": "YYYY.MM.DD",
      "doc_number": "공문번호 (발주처 승인 공문에 감리단 통보 공문이 연이어 있으면 '공사부-XXX → 감리단-YYY' 형식으로 병기)",
      "sender": "발신처 (연쇄 공문이면 '발주처 → 감리단' 형식으로 병기)",
      "receiver": "수신처",
      "subject": "공문 제목",
      "event_type": "지연요인|설계변경|추가공사|인허가|기타",
      "causal_party": "발주처|시공사|설계사|감리|불가항력",
      "causal_description": "귀책 사유 상세 서술 (공문번호·일자 반드시 포함)",
      "delay_days": 0,
      "note": "비고",
      "show_in_table": true
    }
  ],

  "detail_narratives": [
    {
      "title": "블록 제목 (예: 1차 공기연장 경위, 변경계약 체결 경위 등)",
      "paragraphs": [
        "서술 단락 1 — 반드시 리스트(배열) 형태로 작성. 문자열 하나가 한 단락.",
        "서술 단락 2 (필요 시 추가)"
      ]
    }
  ],

  "pre_diagram_paragraphs": [
    "단락 1: 계약조건 조항 인용, 귀책 판단 기준 설명",
    "단락 2: 각 지연 사유별 귀책 귀속 분석 — '계약상대자의 책임 없는 사유에 해당합니다' 등"
  ],

  "accountability_diagram": [
    {
      "⚠️❌_RULE": "합계 행 절대 금지 — cause가 '합계'·'소계'·'계'인 항목을 배열에 추가하는 것은 어떤 경우에도 금지. delay_days 합계는 report_generator가 자동 계산하므로 합계 행 추가 시 total_delay_days의 2배로 오류 발생.",
      "cause": "공기지연 사유 (구체적 원인)",
      "basis": "관련 근거 (공문번호, 변경계약 차수, 계약조건 조항 등)",
      "responsible_party": "발주처",
      "delay_days": 0,
      "source_docs": ["근거 공문번호 또는 변경계약 차수 (예: 제22-0123호, 3차 변경계약)"]
    }
  ],

  "conclusion_paragraphs": [
    "결론 단락 1 — 귀책 판단 요약: 어떤 사유가, 어떤 근거(변경계약서·승인공문)로, 발주처 귀책임이 확인되는지 명시",
    "결론 단락 2 — 청구 적법성: 연장 기간 동안의 간접비를 근거 조항에 따라 청구할 수 있음을 서술"
  ],

  "summary": "[종합] 귀책 원인, 발주처 귀책 확정, 간접비 청구 근거를 2~4문장으로 서술"
}
```

---

### 작성 원칙

- **[최우선] 변경계약서 전수 확인**: 공문 전문에서 변경계약서를 모두 찾아 계약번호 suffix(-01/-02 등), 변경계약일자, 변경사유, 준공기한 변경 내용을 먼저 정리. 2회 이상 공기연장이 있는 경우 모든 회차를 items 및 detail_narratives에 포함.
- **[공문번호 병기]** 발주처 승인 공문에 감리단 통보 공문이 연이어 있는 경우, 두 공문번호를 하나의 item에 '공사부-XXX → 감리단-YYY' 형식으로 반드시 병기. 감리단 통보 공문을 별도 item으로 분리하거나 누락하지 말 것.
- causal_description·background_paragraphs·detail_narratives에 반드시 공문번호와 날짜를 인용할 것
- **conclusion_paragraphs는 반드시 2단락 이상 작성. 빈 배열 [] 절대 금지.**
- **total_delay_days**: 모든 변경계약서의 연장일수 합산값을 정수로 기재 (0이면 안 됨)
- 시간 인과관계 순서: 발신일자 기준 시간순
- 문체: 객관적 서술체, "~한바", "~하였으나", "~에 따라" 등 관용 표현 사용
- 추측·가정 금지 — 공문 및 제공 자료에 실제로 적힌 내용만 사용
- **[수치 출처 필수 인용]** 지연일수, 공제일수, 계산 결과 등 구체적 수치를 기재할 때는 반드시 출처 문서(공문번호 또는 파일명)를 괄호 안에 병기. 예: "사전 작업기간 4일 공제(CM제23-420 붙임 검토의견서)". 수치만 기재하고 출처를 생략하는 것은 금지.
- **[귀책 인과관계 명시]** 발주처 귀책으로 판단한 경우, causal_description 및 detail_narratives에 반드시 다음 세 단계를 포함: ① 발주처의 구체적 행위 또는 부작위(예: 규격 설정 부실, 조달 지연, 일시정지 지시), ② 그 근거 공문(공문번호·날짜), ③ 그로 인한 공사 지연의 인과관계. "발주처 귀책에 해당한다"고만 기재하는 것은 불충분.
- **[계약금액 조정신청 절차 확인]** conclusion_paragraphs에 절차적 요건 준수 여부를 반드시 명시: ① 계약금액 조정신청 공문이 scan 자료에 있으면 해당 공문번호·날짜를 인용. ② 별도 공문이 없는 경우 "본 보고서 제출로써 계약금액 조정신청에 갈음하며, 준공대가 수령 전에 신청이 이루어진 것으로 절차적 요건을 충족한다"는 취지로 기재.
- REFERENCE 패턴의 문장 구조·길이·서술 밀도를 최대한 따를 것
- **[날짜 기준 통일]** items의 `date` 필드는 반드시 **시행일(발신일)**을 사용. 공문에 시행일과 접수일이 모두 표기된 경우, `date`에는 시행일을 기재하고 `note`에 "접수: YYYY.MM.DD (접수번호)" 형식으로 병기.
- **[연장 기간 3층 구분 필수]** 공기연장 일수 산정 근거를 detail_narratives에 서술할 때, 아래 세 층위를 반드시 구분하여 각각 명시할 것:
  1. **원인 발생 기간**: 실제 공사 지장이 발생한 날짜 범위 (예: 재착공일~관급자재 계약 전일)
  2. **변경계약 반영 시점**: 해당 연장일수가 포함된 변경계약 차수 및 변경계약일
  3. **간접비 청구 기간**: 변경계약으로 연장된 준공기한 구간 (당초 준공일 다음날~변경 준공일)
  → 원인 발생 기간(①)과 간접비 청구 기간(③)의 연도·날짜가 다를 수 있음. 혼동 금지.
- **[복수 사유 청구 기간 배분 — 계산 필수]** 하나의 변경계약에 복수의 지연사유가 합산된 경우(예: 레미콘 35일 + 회계연도 36일 = 71일), 아래 방법으로 각 사유별 간접비 청구 기간을 도출하고 계산 근거를 detail_narratives에 명시할 것:
  - 배분 순서: 원인 발생 시점이 앞선 사유가 청구 기간 앞 구간을 차지
  - 계산 방법: 변경계약 시작일(당초 준공일 다음날 = Day 1)부터 각 사유 일수를 순차 합산
    - 사유1 청구기간: Day 1 ~ Day (X일수)
    - 사유2 청구기간: Day (X+1) ~ Day (X+Y일수) = 변경 준공일
  - 계산 결과 날짜를 detail_narratives에 명기. 공문 원문에 날짜가 명시된 경우 그 날짜를 우선 사용하고 출처를 병기. 임의 추정 날짜 사용 절대 금지.
- **[공기연장 사유 전수 포함]** 각 차수의 공기연장 사유가 복수인 경우(예: ① 레미콘 조달지연 ② 회계연도 마감 ③ 지질조건 상이) 모든 사유를 빠짐없이 items 및 detail_narratives에 포함. 공기연장 요청 공문(실정보고)에 명기된 사유 목록과 1:1 대조하여 누락 여부를 직접 확인한 뒤 작성.
- **[background_paragraphs 필수 포함 항목]** background_paragraphs에 반드시 아래 항목을 포함할 것: ① 계약일(최초 계약 체결일), ② 착공일, ③ 당초 준공일, ④ 최종 변경 준공일, ⑤ 총 공기연장 일수 및 변경계약 회차. 이 중 하나라도 누락되면 독자가 공기연장 경위를 파악할 수 없으므로 공문 전문에서 반드시 확인하여 기재할 것.
- **[소결 3단계 필수]** detail_narratives 각 블록 및 pre_diagram_paragraphs에서 서술을 마칠 때 반드시 아래 3단계 소결을 포함: ① 원인 분석("본건 공기연장의 원인은 [구체적 원인]으로..."), ② 귀책분류("공사계약일반조건 제[조]제[항]의 [호]에 해당하며..."), ③ 판단("계약상대자의 책임 없는 사유에 해당합니다"). 사건 나열만 하고 소결 없이 끝내는 것은 불충분.
- **[귀책사유 구체성 필수]** causal_description 및 detail_narratives에서 귀책사유를 "공사용지 미확보", "인허가 지연" 등 추상적 표현으로 기재 금지. 반드시 구체적 위치·허가 유형·기관명을 포함. 예: "공사용지 미확보" → "[TRESTLE & JETTY 구간] 선행 항만공사 지연으로 공사용지 인도 지연", "인허가 지연" → "[도로관리심의 분기제 운영으로 인한] 도로점용허가 지연".
- **[공동수급체 구성원 직접 확인]** 시공사(공동수급체) 구성원을 background_paragraphs 또는 detail_narratives에 기재할 때, 반드시 공사계약서·공동수급협정서에 기재된 업체명과 지분율 전체를 직접 확인하여 기재. 기억이나 추측으로 작성 금지. "대표사 외 N사" 축약도 금지 — 전체 업체명을 나열할 것.
- **[법적 근거 조항 전수 인용]** pre_diagram_paragraphs에서 근거 조항 인용 시, 공문·계약서에 실제로 인용된 조항을 모두 나열. 단순화를 위해 일부만 인용하지 말 것. 국가계약법 공사의 경우 제23조(금액조정)·제26조(연장신청)·제31조 또는 제32조·제47조 해당 여부를 공문에서 직접 확인.
- **[결론 — 청구 금액·계약예규 기준 반드시 명시]** conclusion_paragraphs 마지막 단락에 다음 두 가지를 반드시 포함:
  ① 총 청구 간접비 금액(원, VAT 포함 여부 병기)
  ② 계약예규 적용 기준 — 입찰공고일 또는 계약체결일 기준, 해당 예규 버전 시행일 (예: "입찰공고일 2019.05. 기준, 계약예규 [시행 2019.1.1.] 제16장 실비 산정 적용")
"""



def _extract_contract_changes(items: list[dict]) -> list[dict]:
    """
    scan_result.json items의 full_text에서 변경계약서 관련 항목을 파싱.
    반환: [{"no": N, "subject": "...", "date": "...", "차수": "...",
            "연장일수": "...", "준공변경": "...", "변경사유": "..."}]
    """
    import re as _re

    _KW = _re.compile(
        r"(공사변경계약서|변경계약서|계약변경통보|공사계약.{0,3}변경|변경계약.{0,3}체결|"
        r"준공기한.{0,3}변경|공기.{0,3}연장.{0,6}\d+일|계약기간.{0,3}연장)",
        _re.IGNORECASE,
    )
    # 연장일수: '공기연장 N일', 'N일 연장/증가/조정', 'N일간 연장' 등
    _DAYS_RE = _re.compile(
        r"(?:공기연장|연장|계약기간\s*연장)\s*[:\s]*\s*(\d+)\s*일"
        r"|(\d+)\s*일\s*(?:간\s*)?(?:연장|증가|조정|공기연장)"
    )
    _DATE_RE  = _re.compile(r"(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})")
    # 차수: '제N차', '제N회', 'N-M차' 변경계약/공기연장 등
    _ORD_RE = _re.compile(
        r"제?\s*(\d+(?:-\d+)?)\s*(?:차|회)\s*(?:변경계약|공기연장|계약변경|변경)"
    )
    _REASON_RE = _re.compile(
        r"변경\s*사유\s*[:：]?\s*(.{5,80}?)(?:\n|$)|"
        r"사유\s*[:：]\s*(.{5,80}?)(?:\n|$)"
    )
    # 준공변경 명시 패턴: '당초/변경전 날짜 → 변경/변경후 날짜'
    _CHG_RE = _re.compile(
        r"(?:당초|변경\s*전)[^0-9\n]{0,20}(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})"
        r"[^0-9\n]{0,30}(?:\u2192|변경\s*후?|변경\s*준공)[^0-9\n]{0,20}"
        r"(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})"
    )

    results = []
    for item in items:
        full = item.get("full_text", "") or ""
        subj = item.get("subject", "") or ""
        combined = subj + "\n" + full[:3000]

        if not _KW.search(combined):
            continue

        entry: dict = {
            "no": item.get("no", ""),
            "date": item.get("date", ""),
            "subject": subj[:60],
            "source_file": item.get("file_path", ""),
            "차수": "",
            "연장일수": "",
            "준공변경": "",
            "변경사유": "",
        }

        m = _ORD_RE.search(combined)
        if m:
            entry["차수"] = f"제{m.group(1)}차"

        # 연장일수: 두 그룹 중 값 있는 것 수집, 중복 제거
        seen_days: list[str] = []
        for dm in _DAYS_RE.finditer(combined):
            val = dm.group(1) or dm.group(2)
            if val and val not in seen_days:
                seen_days.append(val)
        if seen_days:
            entry["연장일수"] = "+".join(seen_days) + "일"

        # 준공변경: 명시 패턴('당초→변경') 우선, fallback 단순 날짜 쌍
        cm = _CHG_RE.search(combined[:2000])
        if cm:
            entry["준공변경"] = f"{cm.group(1)} \u2192 {cm.group(2)}"
        else:
            dates = _DATE_RE.findall(combined[:1000])
            if len(dates) >= 2:
                entry["준공변경"] = f"{dates[0]} \u2192 {dates[-1]}"

        m = _REASON_RE.search(combined)
        if m:
            entry["변경사유"] = (m.group(1) or m.group(2) or "").strip()[:60]

        results.append(entry)

    return results


def build(output_dir: Path, project_name: str = "") -> Path:
    """
    output 폴더의 파일들을 조합하여 claude.ai용 프롬프트 생성.
    반환: 생성된 프롬프트 파일 경로
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # ── 파일 로드 ──────────────────────────────────────────────────────────────
    corr_texts_path  = output_dir / "correspondence_texts.md"
    scan_result_path = output_dir / "scan_result.json"

    # 스캔 결과 파일은 필수
    missing = [p.name for p in [corr_texts_path, scan_result_path] if not p.exists()]
    if missing:
        raise FileNotFoundError(
            f"필요한 파일이 없습니다: {', '.join(missing)}\n"
            "스캔이 완료되지 않았습니다. 먼저 [1] 스캔을 실행하세요."
        )

    # reference_patterns.md: BASE_DIR/output/ 먼저, 없으면 BUNDLE_DIR/output/ 확인
    # (exe 배포 시 번들 내 _internal/output/에 위치)
    _ref_candidates = [
        output_dir.parent / "reference_patterns.md",          # BASE_DIR/output/ (일반·배포 후 복사)
        config.BUNDLE_DIR / "output" / "reference_patterns.md",  # BUNDLE_DIR/_internal/output/
    ]
    ref_path_found = next((p for p in _ref_candidates if p.exists()), None)
    if ref_path_found:
        ref_text = ref_path_found.read_text(encoding="utf-8")
    else:
        ref_text = ""
        print("  ⚠️  reference_patterns.md 없음 — 참고 패턴 없이 프롬프트를 생성합니다.")

    corr_text = corr_texts_path.read_text(encoding="utf-8")

    # 귀책분석 패턴집: BASE_DIR 먼저, 없으면 BUNDLE_DIR 확인
    _pb_candidates = [
        output_dir.parent.parent / "귀책분석_패턴집.md",   # BASE_DIR/
        config.BUNDLE_DIR / "귀책분석_패턴집.md",           # _internal/
    ]
    pb_path_found = next((p for p in _pb_candidates if p.exists()), None)
    pattern_book_text = pb_path_found.read_text(encoding="utf-8") if pb_path_found else ""

    with open(scan_result_path, encoding="utf-8") as f:
        scan_data = json.load(f)

    # 확정 공문 목록 표
    items = scan_data.get("items", [])
    if not items:
        raise ValueError("scan_result.json 에 공문 항목이 없습니다.")

    # vendor_dirs (리스트) 또는 구버전 vendor_dir (문자열) 모두 처리
    vendor_dirs = scan_data.get("vendor_dirs", [])
    if not vendor_dirs:
        vendor_dirs = [scan_data.get("vendor_dir", "")]
    vendor_dir_display = "\n".join(f"  - {d}" for d in vendor_dirs)
    stats = scan_data.get("stats", {})

    # ── 변경계약서 사전 파싱 ────────────────────────────────────────────
    contract_changes = _extract_contract_changes(items)

    # ── OCR 품질 주의 항목 수집 ──────────────────────────────────────────
    ocr_warn_items = [i for i in items if i.get("ocr_quality") == "WARN"]

    # ── 계약 유형별 조항 참조표 생성 ────────────────────────────────────
    _ctbl = ["## [계약 유형별 귀책 근거 조항 참조표]", "",
             "⚠️ **공문·변경계약서에 실제로 인용된 조항만 사용하세요.** 아래는 참조용입니다.",
             "", "| 조항 | 내용 | 적용 계약 유형 |", "|------|------|--------------|"]
    _seen: set[str] = set()
    for _lt, _rows in config.CONTRACT_CLAUSES.items():
        for _cl, _ds in _rows:
            if _cl not in _seen:
                _ctbl.append(f"| {_cl} | {_ds} | {_lt} |")
                _seen.add(_cl)
    clause_section = "\n".join(_ctbl)

    # ── 프롬프트 조립 ──────────────────────────────────────────────────────────
    pname = project_name or scan_data.get("project_name", "")

    prompt_lines = [
        "# 귀책분석 작성 요청",
        "",
        "당신은 건설/엔지니어링 프로젝트의 '2장 귀책분석' 전문 작성자입니다.",
        "",
        "**[Claude Code 필독] 아래 두 파일을 순서대로 Read 도구로 읽은 뒤 작업하세요:**",
        f"1. 이 파일 (`prompt_for_claude.md`) — 지시사항 + 확정 공문 목록",
        f"2. `{corr_texts_path}` — 공문 전문 전체",
        "",
        "두 파일을 모두 읽은 뒤 아래 지시에 따라 귀책분석_data.json을 작성하세요.",
        "",
        "---",
        "",
        "## [프로젝트 정보]",
        "",
        f"- 프로젝트명: {pname if pname else '(아래 공문에서 확인)'}",
        f"- 공문 수신 경로:\n{vendor_dir_display}",
        f"- 전체 파일 수: {stats.get('total_files', '?')}개",
        f"- 귀책분석 관련 확정 공문: {stats.get('relevant_confirmed', len(items))}개",
        "",
        "---",
        "",
        "## [귀책분석 패턴집 — 수신자료 기반 판단 가이드]",
        "",
        "(아래는 실제 처리된 프로젝트들에서 누적된 귀책 유형, 핵심 공문 패턴, 판단 로직입니다.",
        "수신자료만으로 귀책사유를 판단할 때 이 패턴집을 우선 참고하세요.)",
        "",
        pattern_book_text if pattern_book_text else "(패턴집 없음 — 최초 프로젝트)",
        "",
        "---",
        "",
        "## [REFERENCE 보고서 패턴]",
        "",
        "(아래는 실제 귀책분석 보고서 13개에서 추출한 패턴입니다.",
        "표 구조, 문체, 귀책 판단 방식, 공문 인용 방식을 학습하세요.)",
        "",
        ref_text,
        "",
        "---",
        "",
        clause_section,
        "",
        "---",
        "",
        *((["⚠️ **OCR 품질 주의 파일 목록** — 아래 파일의 공문번호·날짜·발신처를 직접 확인하세요.", ""] +
          [f"- `{i.get('file_path', i.get('subject', '?'))}`" for i in ocr_warn_items] +
          ["", "---", ""])
         if ocr_warn_items else []),
        "## [확정 공문 목록]",
        "",
        "**중요: 아래 표의 `scan_no`와 `source_file` 값을 data.json items에 그대로 복사하세요.**",
        "(`scan_no`와 `source_file`은 원본 파일 추적에 사용됩니다. 절대 수정하지 마세요.)",
        "",
        "⚠️ **OCR 신뢰도 주의**: 공문번호·발신처·수신처는 OCR 자동추출값으로 부정확할 수 있습니다.",
        "- 발신일자는 파일명 날짜(YYMMDD 접두사)로 교정되었습니다.",
        "- 공문번호·발신처가 의심스러운 경우 아래 [공문 전문] 및 source_file 경로(폴더명 포함)를 참조하여 직접 확인하세요.",
        "- 폴더명(예: `감리단 제출`, `발주처 공문`, `2206_간접비 청구`)은 발신/수신 방향과 업무 맥락 파악에 활용하세요.",
        "- **(참고)** 표시 항목은 귀책 키워드 미매칭 자동 포함 항목입니다. 귀책분석와 무관하면 items에서 제외하세요.",
        "",
        "| scan_no | 발신일자 | 공문번호 | 발신처 | 수신처 | 제목 | source_file (폴더\\파일명) |",
        "|---------|---------|---------|--------|--------|------|--------------------------|",
    ]

    for item in items:
        fp = item.get('file_path', '')
        # 마지막 2단계 경로 표시 (상위폴더\파일명) — 발신 방향·맥락 파악에 활용
        parts = fp.replace('\\', '/').split('/')
        source_display = '/'.join(parts[-2:]) if len(parts) >= 2 else fp
        relevance_label = " (참고)" if item.get('relevance') == 'borderline' else ""
        prompt_lines.append(
            f"| {item.get('no', '')} | {item.get('date', '')}{relevance_label} | "
            f"{item.get('doc_number', '')} | {item.get('sender', '')} | "
            f"{item.get('receiver', '')} | {item.get('subject', '')} | "
            f"{source_display} |"
        )

    prompt_lines += [
        "",
        "---",
        "",
        "## [공문 전문]",
        "",
        "**공문 전문은 아래 파일에 저장되어 있습니다. 반드시 파일을 직접 열어 확인하세요:**",
        "",
        f"  `{corr_texts_path}`",
        "",
        "*(Claude Code: 위 경로의 파일을 Read 도구로 읽어 공문 내용을 분석하세요)*",
        "",
        "---",
        "",
        "## [⚡ 변경계약서 자동 감지 결과 — 반드시 원문과 대조하세요]",
        "",
        *((
            ["아래는 스캔 자료에서 자동 감지한 변경계약서 관련 항목입니다.",
             "공문번호·일수·사유는 OCR 품질에 따라 부정확할 수 있습니다. 반드시 원문 확인 후 사용하세요.",
             "",
             "| No | 날짜 | 차수 | 연장일수 | 준공변경 | 감지된 변경사유 | 파일 |",
             "|-----|------|------|---------|---------|--------------|------|"] +
            [f"| {c['no']} | {c['date']} | {c['차수']} | {c['연장일수']} "
             f"| {c['준공변경']} | {c['변경사유'][:40]} "
             f"| {str(c['source_file']).replace(chr(92),'/').split('/')[-1]} |"
             for c in contract_changes]
        ) if contract_changes else
            ["(변경계약서 관련 항목이 자동 감지되지 않았습니다.",
             " 공문 전문에서 직접 검색하여 확인하세요.)"]
        ),
        "",
        "---",
        "",
        "## [data.json 작성 전 필수 — 변경계약서 사전 확인]",
        "",
        "아래 순서대로 확인한 뒤 data.json을 작성하세요.",
        "",
        "**STEP 1. 변경계약서 전수 목록 작성**",
        "공문 전문에서 '공사변경계약서' 또는 '공사계약변경통보서'를 모두 찾아 아래 형식으로 정리하세요:",
        "- 계약번호: 23000505-XX  |  변경계약일: YYYY.MM.DD  |  변경사유: (원문 그대로)  |  준공기한: 변경 전 → 변경 후  |  연장일수: N일",
        "",
        "**STEP 2. 총 공기연장 일수 확정**",
        "각 변경계약서의 연장일수를 합산하여 총 연장일수를 확정하세요.",
        "변경계약이 2회 이상인 경우 모든 회차를 items와 detail_narratives에 포함해야 합니다.",
        "",
        "**STEP 3. 회차별 귀책사유 분류**",
        "각 변경계약서의 변경사유(원문)를 근거로, 발주처 귀책인지 여부를 판단하세요.",
        "변경계약서에 직접 명기된 사유는 발주처가 인정한 것으로 가장 강력한 근거입니다.",
        "",
        "위 STEP을 완료한 후 JSON을 작성하세요.",
        "",
        "---",
        "",
        SCHEMA_GUIDE,
    ]

    prompt_path = output_dir / "prompt_for_claude.md"
    prompt_path.write_text("\n".join(prompt_lines), encoding="utf-8")

    # ── JSON 스키마 파일 저장 (빈 템플릿 — 모든 서술 필드 포함) ─────────────
    schema = {
        "project_name": pname,
        "total_delay_days": 0,
        "background_paragraphs": [],
        "items": [
            {
                "no": item.get("no", i + 1),
                "date": item.get("date", ""),
                "doc_number": item.get("doc_number", ""),
                "sender": item.get("sender", ""),
                "receiver": item.get("receiver", ""),
                "subject": item.get("subject", ""),
                "event_type": "",
                "causal_party": "",
                "causal_description": "",
                "delay_days": 0,
                "note": "",
                "show_in_table": True,
            }
            for i, item in enumerate(items)
        ],
        "detail_narratives": [],
        "pre_diagram_paragraphs": [],
        "accountability_diagram": [],
        "conclusion_paragraphs": [],
        "summary": "",
    }
    schema_path = output_dir / "귀책분석_schema.json"
    schema_path.write_text(
        json.dumps(schema, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"\n저장 완료:")
    print(f"  - {prompt_path}  ← 이 파일을 Claude에 전달")
    print(f"  - {schema_path}  ← JSON 스키마 참고용")

    return prompt_path

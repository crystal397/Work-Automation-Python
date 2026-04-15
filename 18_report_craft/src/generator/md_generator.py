"""
마크다운 보고서 생성기
analysis_result.json + 계산 결과 → 보고서_초안.md
"""

from __future__ import annotations
from datetime import date
from . import laws_db
import config as _cfg


# ── 비목명 합산 맵 ─────────────────────────────────────────────────────────────
# 유사 비목을 하나의 대표명으로 통합하여 보고서에 표시
_EXPENSE_MERGE_MAP: dict[str, str] = {
    "전력비":       "전력수도광열비",
    "전기료":       "전력수도광열비",
    "전기세":       "전력수도광열비",
    "수도광열비":   "전력수도광열비",
    "수도비":       "전력수도광열비",
    "광열비":       "전력수도광열비",
    "여비교통비":   "여비교통통신비",
    "교통비":       "여비교통통신비",
    "여비":         "여비교통통신비",
    "통신비":       "여비교통통신비",
    "통신요금":     "여비교통통신비",
}


def _merge_direct_rows(rows: list[dict]) -> list[dict]:
    """유사 비목명을 합산하여 반환 (전력비+수도광열비→전력수도광열비 등)"""
    merged: dict[str, dict] = {}
    for row in rows:
        raw_name    = (row.get("item") or "").strip()
        merged_name = _EXPENSE_MERGE_MAP.get(raw_name, raw_name)
        if merged_name not in merged:
            merged[merged_name] = {
                "item":      merged_name,
                "actual":    0,
                "estimated": 0,
                "sources":   [],
            }
        merged[merged_name]["actual"]    += row.get("actual", 0) or 0
        merged[merged_name]["estimated"] += row.get("estimated", 0) or 0
        src = row.get("source", "")
        if src and src not in merged[merged_name]["sources"]:
            merged[merged_name]["sources"].append(src)

    result = []
    for d in merged.values():
        result.append({
            "item":      d["item"],
            "actual":    d["actual"],
            "estimated": d["estimated"],
            "total":     d["actual"] + d["estimated"],
            "source":    ", ".join(d["sources"]),
        })
    return result


# ── 유틸 ──────────────────────────────────────────────────────────────────────

def _fmt(v: int | None) -> str:
    """정수를 천 단위 쉼표 포맷으로 변환"""
    if v is None:
        return "확인 필요"
    return f"{v:,}"


def _pct(v: float | None) -> str:
    if v is None:
        return "확인 필요"
    return f"{v:.2%}"


def _or(v, fallback="확인 필요"):
    return v if v else fallback


def _bq(text: str) -> str:
    """법령 원문 멀티라인을 blockquote 형식으로 변환.
    \\n\\n → 빈 blockquote 행(문단 구분), \\n → 동일 blockquote 내 줄바꿈"""
    return text.replace("\n\n", "\n>\n> ").replace("\n", "\n> ")


def _law_box(law: dict) -> str:
    """법령 조문을 레퍼런스 보고서 형식으로 포맷.

    레퍼런스 docx 구조:
      Row 0: 「법률명(약칭 : 단축명)」   ← 굵게, 가운데 정렬
             [시행일] [법률/예규번호]    ← 굵게, 가운데 정렬 (enactment 있을 때)
      Row 1: 제XX조(조문 제목)           ← 굵게, 왼쪽 정렬
             본문 (일부 **<u>강조</u>**) ← 일반 (bold+underline 강조구절 포함)
    """
    # 법률명: 'name' 우선, 없으면 'law' (INSURANCE_LAWS 등)
    name      = law.get("name") or law.get("law", "")
    short     = law.get("short", "")
    article   = law.get("article", "")
    art_title = law.get("article_title", "")
    enactment = law.get("enactment", "")
    content   = law.get("content", "")
    purpose   = law.get("purpose", "")

    # 법률명 헤더: 약칭이 있으면 (약칭 : XXX) 추가
    name_hdr = name
    if short and short not in name:
        name_hdr = f"{name}(약칭 : {short})"

    # 조문 헤더: 제XX조(조문제목) 형식
    art_str = f"{article}({art_title})" if art_title else article

    lines = [f"> **「{name_hdr}」**"]
    if enactment:
        lines.append(f"> **[{enactment}]**")
    lines.append(">")
    lines.append(f"> **{art_str}**")
    if purpose:
        lines.append(f"> *— {purpose}*")
    lines.append(">")
    lines.append(f"> {_bq(content)}")
    return "\n".join(lines) + "\n"


def _fmt_date_kr(iso_date: str) -> str:
    """ISO 날짜 문자열(2024-11-07) → 한국어 날짜(2024. 11. 7.) 변환"""
    try:
        parts = iso_date.split("-")
        if len(parts) == 3:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            return f"{y}. {m}. {d}."
    except Exception:
        pass
    return iso_date


def _sub_period_prose(report_type: str, name: str,
                      calc_s: str, calc_e: str, days: int | str,
                      period_s: str, period_e: str) -> str:
    """하도급사별 공기연장 기간 설명 산문 (유형별 법령 인용)"""
    kr_s = _fmt_date_kr(calc_s)
    kr_e = _fmt_date_kr(calc_e)

    # claim 기간이 원래 계약기간과 동일하면 "계약기간인", 다르면 "원도급사 공기연장 기간 중"
    if calc_s == period_s and calc_e == period_e:
        period_desc = f"계약기간인 {kr_s}에서 {kr_e}까지 {days}일"
    else:
        period_desc = f"원도급사의 공기연장 기간 중 공사기간이 연장된 기간인 {kr_s}에서 {kr_e}까지 {days}일"

    if report_type == "A":
        law_intro = (
            "본건 지방자치단체 입찰 및 계약집행기준 중 공사계약 일반조건 제7절 4. 가. 및 "
            "지방자치단체 입찰 및 계약집행기준 제7절 1. 가. 에 따라 원도급 공사의 계약기간이 "
            "연장된 경우 하도급업체가 지출한 비용을 포함하여 산정하도록 규정하고 있습니다."
        )
    elif report_type == "B":
        law_intro = (
            "본건 공사계약 일반조건 제23조 및 「(계약예규) 정부 입찰·계약 집행기준」 제16장에 따라 "
            "원도급 공사의 계약기간이 연장된 경우 하도급업체가 지출한 비용을 포함하여 "
            "산정하도록 규정하고 있습니다."
        )
    else:
        law_intro = (
            "본건 공사계약 일반조건의 규정에 따라 원도급 공사의 계약기간이 연장된 경우 "
            "하도급업체가 지출한 비용을 포함하여 산정하도록 규정하고 있습니다."
        )

    return (
        f"{law_intro} 따라서 하도급사인 {name}의 간접비 산정 기간은 "
        f"{period_desc}에 대하여 산정하였습니다."
    )


# ── 섹션 빌더 ─────────────────────────────────────────────────────────────────

def _toc(data: dict) -> str:
    """목차 (Table of Contents)"""
    subcontractors = data.get("subcontractors", [])
    lines = [
        "---\n",
        "## 목차\n",
        "- 제1장 개요",
        "  - 제1절 과업의 목적",
        "  - 제2절 공사의 개요",
        "  - 제3절 계약현황",
        "  - 제4절 과업 수행 절차",
        "- 제2장 계약의 성격",
        "  - 제1절 적용 법령 체계",
        "  - 제2절 계약문서 구성 및 주요 용어 정의",
        "  - 제3절 계약당사자 현황",
        "  - 제4절 장기계속공사(계속비공사)의 성격",
        "  - 제5절 내역입찰 공사의 특성",
        "- 제3장 공기연장 귀책 분석 및 계약금액조정 신청 적정성 검토",
        "  - 제1절 공기지연의 귀책사유 분석",
        "    - 1.1 공기지연 관련 공문 이력",
        "    - 1.2 귀책사유 분류",
        "    - 1.3 변경계약 귀책사유",
        "  - 제2절 청구의 근거",
        "    - 2.1 계약에 의한 청구권",
        "    - 2.2 법령에 의한 청구권",
        "    - 2.3 사정변경의 원칙",
        "    - 2.4 청구의 근거 검토 (종합)",
        "- 제4장 공기연장 간접비 산정",
        "  - 제1절 공기연장 기간의 산정",
        "    - 1.1 공기연장 일수의 산정방식",
        "    - 1.2 공기연장 일수 산정",
        "  - 제2절 간접비 산정방식",
        "  - 제3절 원도급사 간접비 산정 결과",
        "    - 3.1 집계표",
        "    - 3.2 간접노무비 상세",
        "    - 3.3 경비 상세",
        "    - 3.4 일반관리비 및 이윤",
    ]
    if subcontractors:
        lines += [
            "  - 제4절 하도급사 간접비 산정 결과",
            "    - 4.1 집계표",
            "    - 4.2 간접노무비",
            "    - 4.3 경비",
            "    - 4.4 일반관리비 및 이윤",
        ]
    lines += [
        "- 결론",
        "- 제5장 첨부자료",
        "",
    ]
    return "\n".join(lines)


def _cover(data: dict, calc: dict) -> str:
    c = data.get("contract", {})
    ext = data.get("extension", {})
    grand = calc.get("grand_total", {})
    labor = calc.get("indirect_labor", {})
    expenses = calc.get("expenses", {})
    subtotal = calc.get("subtotal", {})
    admin = calc.get("general_admin", {})
    profit = calc.get("profit", {})
    vat = calc.get("vat", 0)
    final = calc.get("final_rounded", 0)
    sub_results = calc.get("subcontractor_results", [])
    vat_zero = data.get("vat_zero_rated", False)

    author = getattr(_cfg, "REPORT_AUTHOR", "") or ""
    contractor = _or(c.get("contractor"))
    contract_name = _or(c.get("name"))

    # ── 제출문 ────────────────────────────────────────────────────────────────
    report_type  = data.get("report_type", "A")
    report_month = date.today().strftime("%Y년 %m월")

    if report_type == "A":
        law_ref = "지방계약법령, 계약예규"
    elif report_type == "B":
        law_ref = "국가계약법령, 계약예규"
    else:
        law_ref = "공사계약 일반조건, 관련 법령"

    lines = [
        "# 공기연장 간접비 산정 보고서\n",
        "---\n",
        "## 제 출 문\n",
        f"{contractor} 귀 중\n",
        f"귀사로부터 의뢰받은 「{contract_name}」현장의 공사기간 연장과 관련하여, "
        f"본 공사의 계약조건(공사계약 일반조건, 설계서), {law_ref}, 관련 판례 및 "
        "귀사로부터 제공받은 근거자료를 토대로 간접비를 산정하여 아래와 같이 보고서를 제출합니다.\n",
        f"{report_month}\n",
        f"{author if author else '[작성기관명]'}\n",
        "",
    ]

    # ── 요약문 ────────────────────────────────────────────────────────────────
    vat_note = "영세율(0%)" if vat_zero else "10%"
    jv_ratio       = calc.get("jv_share_ratio", 1.0)
    grand_after_jv = calc.get("grand_after_jv", grand.get("total", 0))

    lines += [
        "---\n",
        "## 요 약 문\n",
        f"본 과업은 「{contract_name}」의 공기연장에 따른 간접비를 산정함으로써 "
        "합리적인 업무수행에 필요한 기초 참고 자료를 제공하는데, 그 목적이 있다.\n",
        f"{contractor}의 「{contract_name}」 현장 및 본사에서 제공한 본계약 관련 계약문서 및 "
        "관련 서류의 내용을 근거로 하여 해당 절차에 따라 산정하였으며 그 요약 결과는 다음과 같다.\n",
        "**＜산 정 결 과＞**\n",
        "*(단위: 원, 부가세 포함)*\n",
        "| 구 분 | | 산정 기간 | 산정 금액 | 비 고 |",
        "|------|------|------|------|------|",
    ]

    # 원도급 행
    ext_s    = _or(ext.get("start_date"))
    ext_e    = _or(ext.get("end_date"))
    ext_days = _or(ext.get("total_days"))
    est_flag = any(r.get("estimated_days", 0) > 0
                   for r in calc.get("indirect_labor", {}).get("rows", []))
    orig_note = "추정 포함" if est_flag else ""
    lines.append(
        f"| 가. 공기연장에 따른 간접비 | (원도급사) "
        f"| {ext_s}~{ext_e}({ext_days}일) | {_fmt(final)} | {orig_note} |"
    )

    # 하도급사 행
    sub_grand_total = 0
    for sr in sub_results:
        agg  = sr.get("aggregate") or {}
        name = sr.get("name", "")
        p_s  = _or(sr.get("period_start"))
        p_e  = _or(sr.get("period_end"))
        # 일수 계산
        from datetime import datetime as _dt
        try:
            days = (_dt.strptime(p_e, "%Y-%m-%d") - _dt.strptime(p_s, "%Y-%m-%d")).days + 1
        except Exception:
            days = "—"
        sub_final = agg.get("final_rounded", 0) or 0
        sub_grand_total += sub_final
        sub_est = any(r.get("estimated_days", 0) > 0
                      for r in (sr.get("labor") or {}).get("rows", []))
        sub_note = "추정 포함" if sub_est else ""
        lines.append(
            f"| 나. 공기연장에 따른 간접비 | {name} "
            f"| {p_s}~{p_e}({days}일) | {_fmt(sub_final)} | {sub_note} |"
        )

    if sub_results:
        lines.append(
            f"| 나. 공기연장에 따른 간접비 | 소 계 | | {_fmt(sub_grand_total)} | |"
        )

    grand_total_all = final + sub_grand_total
    lines += [
        f"| **합 계 (가 + 나)** | | | **{_fmt(grand_total_all)}** | |",
        "",
        f"> **총 청구금액: 금 {_fmt(grand_total_all)}원** (부가가치세 포함, 천원 미만 절사)\n",
    ]

    # ── 요약문 2: 비목별 원도급+하도급 합산 상세 테이블 ──────────────────────
    # 하도급사 비목별 합계 집계
    sub_labor_tot = sum(
        (sr.get("aggregate") or {}).get("indirect_labor", {}).get("total", 0) or 0
        for sr in sub_results
    )
    sub_exp_tot = sum(
        (sr.get("aggregate") or {}).get("expenses", {}).get("total", 0) or 0
        for sr in sub_results
    )
    sub_exp_direct_tot = sum(
        ((sr.get("aggregate") or {}).get("expenses", {}).get("total_direct_actual", 0) or 0) +
        ((sr.get("aggregate") or {}).get("expenses", {}).get("total_direct_estimated", 0) or 0)
        for sr in sub_results
    )
    sub_exp_rate_tot = sum(
        sum((rr.get("total") or 0) for rr in
            ((sr.get("aggregate") or {}).get("expenses", {}).get("rate_rows", []) or []))
        for sr in sub_results
    )
    sub_subtotal_tot = sum(
        (sr.get("aggregate") or {}).get("subtotal", {}).get("total", 0) or 0
        for sr in sub_results
    )
    sub_admin_tot = sum(
        (sr.get("aggregate") or {}).get("general_admin", {}).get("total", 0) or 0
        for sr in sub_results
    )
    sub_profit_tot = sum(
        (sr.get("aggregate") or {}).get("profit", {}).get("total", 0) or 0
        for sr in sub_results
    )
    sub_grand_pretax = sum(
        (sr.get("aggregate") or {}).get("grand_total", {}).get("total", 0) or 0
        for sr in sub_results
    )
    sub_vat_tot = sum(
        (sr.get("aggregate") or {}).get("vat", 0) or 0
        for sr in sub_results
    )

    # 원도급 비목별
    orig_labor  = labor.get("total", 0) or 0
    orig_exp    = expenses.get("total", 0) or 0
    orig_exp_d  = ((expenses.get("total_direct_actual", 0) or 0) +
                   (expenses.get("total_direct_estimated", 0) or 0))
    orig_exp_r  = sum((rr.get("total") or 0) for rr in expenses.get("rate_rows", []))
    orig_sub    = subtotal.get("total", 0) or 0
    orig_admin  = admin.get("total", 0) or 0
    orig_profit = profit.get("total", 0) or 0
    orig_grand  = grand.get("total", 0) or 0
    orig_vat    = vat or 0

    has_sub_data = any(
        (sr.get("aggregate") or {}).get("indirect_labor", {}).get("total", 0)
        for sr in sub_results
    )

    lines += [
        "*(단위: 원, 추정치 포함)*\n",
        "| 구 분 | | 원도급사 금액 | 하도급사 금액 | 산정 금액 | 비 고 |",
        "|------|------|------|------|------|------|",
        f"| 1 | 간접노무비 | {_fmt(orig_labor)} | {_fmt(sub_labor_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_labor + sub_labor_tot if has_sub_data else orig_labor)} | |",
        f"| 2 | 경비 | {_fmt(orig_exp)} | {_fmt(sub_exp_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_exp + sub_exp_tot if has_sub_data else orig_exp)} | |",
        f"|   | 가. 직접계상비목 | {_fmt(orig_exp_d)} | {_fmt(sub_exp_direct_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_exp_d + sub_exp_direct_tot if has_sub_data else orig_exp_d)} | |",
        f"|   | 나. 승률계상비목 | {_fmt(orig_exp_r)} | {_fmt(sub_exp_rate_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_exp_r + sub_exp_rate_tot if has_sub_data else orig_exp_r)} | |",
        f"| 3 | 소계 | {_fmt(orig_sub)} | {_fmt(sub_subtotal_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_sub + sub_subtotal_tot if has_sub_data else orig_sub)} | 1+2 |",
        f"| 4 | 일반관리비 | {_fmt(orig_admin)} | {_fmt(sub_admin_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_admin + sub_admin_tot if has_sub_data else orig_admin)} | (3)×요율 |",
        f"| 5 | 이윤 | {_fmt(orig_profit)} | {_fmt(sub_profit_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_profit + sub_profit_tot if has_sub_data else orig_profit)} | (3+4)×요율 |",
        f"| 6 | 총원가 | {_fmt(orig_grand)} | {_fmt(sub_grand_pretax) if has_sub_data else '—'} "
        f"| {_fmt(orig_grand + sub_grand_pretax if has_sub_data else orig_grand)} | 3+4+5 |",
        f"| 7 | 부가가치세 ({vat_note}) | {_fmt(orig_vat)} | {_fmt(sub_vat_tot) if has_sub_data else '—'} "
        f"| {_fmt(orig_vat + sub_vat_tot if has_sub_data else orig_vat)} | (6)×{'0%' if vat_zero else '10%'} |",
        f"| 8 | **총 액** | **{_fmt(final)}** | **{_fmt(sub_grand_total) if has_sub_data else '—'}** "
        f"| **{_fmt(grand_total_all)}** | 천원 미만 절사 |",
        "",
    ]

    return "\n".join(lines)


def _section1_overview(data: dict) -> str:
    c = data.get("contract", {})
    changes = data.get("changes", [])
    ext = data.get("extension", {})
    subcontractors = data.get("subcontractors", [])

    # 제1절 과업의 목적
    lines = [
        "---\n",
        "# 제1장 개요\n",
        "## 제1절 과업의 목적\n",
        f"본 과업은 「{_or(c.get('name'))}」의 공기연장에 따른 간접비를 산정함으로써, "
        "계약상대자의 정당한 계약금액 조정 청구를 위한 기초 참고 자료를 제공하는 것을 목적으로 한다.\n",
        "공기연장 간접비는 발주자의 책임있는 사유 또는 불가항력 등 계약상대자의 책임없는 사유로 인하여 "
        "공사기간이 연장된 경우, 관련 법령 및 공사계약 일반조건에 근거하여 청구할 수 있다.\n",
    ]

    # 제2절 공사의 개요 표
    supervisor = _or(c.get("supervisor"))
    lines += [
        "## 제2절 공사의 개요\n",
        "| 구 분 | 내 용 | 출처 |",
        "|------|------|------|",
        f"| 계약명 | {_or(c.get('name'))} (이하 '본건 공사') | {_or(c.get('name_source'))} |",
        f"| 발주자 | {_or(c.get('client'))} | {_or(c.get('client_source'))} |",
        f"| 계약상대자 | {_or(c.get('contractor'))} | {_or(c.get('contractor_source'))} |",
        f"| 건설사업관리단 | {supervisor} | — |",
        f"| 계약체결일 | {_or(c.get('initial_date'))} | {_or(c.get('initial_date_source'))} |",
        f"| 계약금액 (1차수 최초) | ₩{_fmt(c.get('initial_amount'))} | {_or(c.get('initial_amount_source'))} |",
        f"| 공사기간 (1차수) | {_or(c.get('initial_start'))} ~ {_or(c.get('initial_end'))} → "
        f"{_or(ext.get('end_date'))} (연장 {_or(ext.get('total_days'))}일) | {_or(ext.get('source'))} |",
        f"| 계약형태 | {_or(c.get('contract_type'))} | {_or(c.get('contract_type_source'))} |",
        f"| 공사의 주요내용 | {_or(c.get('scope_description'))} | — |",
        "",
        "> ※ 위치도 (평면도, 종단면도 등) — [첨부 예정]\n",
    ]

    # 제3절 계약현황
    total_c = data.get("total_contract", {})
    lines += [
        "## 제3절 계약현황\n",
    ]
    # 총공사 계약 현황 (장기계속계약의 경우)
    if total_c:
        lines += [
            "### 총공사 계약 현황\n",
            "| 구분 | 계약일 | 총계약금액(원) | 총공사기간 | 비 고 |",
            "|------|------|------|------|------|",
            f"| 최초 | {_or(total_c.get('initial_date'))} | {_fmt(total_c.get('initial_amount'))} "
            f"| {_or(total_c.get('initial_start'))} ~ {_or(total_c.get('initial_end'))} | 총공사 최초 계약 |",
        ]
        for ch in total_c.get("changes", []):
            lines.append(
                f"| {ch.get('seq','')}차 변경 | {_or(ch.get('date'))} | {_fmt(ch.get('amount'))} "
                f"| — ~ {_or(ch.get('new_end_date'))} | {_or(ch.get('reason'))} |"
            )
        lines += ["", f"> ※ 총공사 계약금액: {_fmt(total_c.get('final_amount') or total_c.get('initial_amount'))}원 / 총공사 기간: {_or(total_c.get('total_days'))}일\n"]
    lines += [
        "### 1차수 공사 계약 현황\n",
        "| 구분 | 계약일 | 계약금액(원) | 공사기간 | 연장일수 | 사유 |",
        "|------|------|------|------|------|------|",
        f"| 최초 | {_or(c.get('initial_date'))} | {_fmt(c.get('initial_amount'))} "
        f"| {_or(c.get('initial_start'))} ~ {_or(c.get('initial_end'))} | — | — |",
    ]
    for ch in changes:
        lines.append(
            f"| {ch.get('seq', '')}차 변경 | {_or(ch.get('date'))} | {_fmt(ch.get('amount'))} "
            f"| — ~ {_or(ch.get('new_end_date'))} | 증 {_or(ch.get('extension_days'))}일 "
            f"| {_or(ch.get('reason'))} |"
        )
    lines += [
        "",
        f"> **총 연장일수**: {_or(ext.get('total_days'))}일 "
        f"({_or(ext.get('start_date'))} ~ {_or(ext.get('end_date'))})\n",
    ]

    # 공기연장 타임라인
    if changes:
        lines += [
            "**공사기간 변경 현황 (타임라인)**\n",
            "| 시점 | 준공기한 | 비 고 |",
            "|------|------|------|",
            f"| {_or(c.get('initial_date'))} (최초 계약) | {_or(c.get('initial_end'))} | 1차수 최초 준공기한 |",
        ]
        for ch in changes:
            ext_days = ch.get('extension_days')
            note = f"+{ext_days}일" if ext_days and ext_days > 0 else "공기 변동 없음 (설계변경)"
            lines.append(
                f"| {_or(ch.get('date'))} ({ch.get('seq','')}차 변경) "
                f"| {_or(ch.get('new_end_date'))} | {note} |"
            )
        lines.append("")

    # 하도급 계약 현황
    if subcontractors:
        lines += [
            "### 하도급 계약 현황\n",
        ]
        for sub in subcontractors:
            name           = sub.get("name", "")
            sub_changes    = sub.get("changes", [])
            initial_amount = sub.get("initial_amount")
            initial_date   = sub.get("contract_date") or sub.get("initial_date")
            orig_start     = _or(sub.get("period_start"))
            orig_end       = _or(sub.get("period_end"))

            lines += [
                f"#### {name}\n",
                f"본건 공사의 하도급사인 {name}의 계약현황은 아래와 같습니다.\n",
                "| 구 분 | 계약일 | 공사기간 | 증감일 | 계약금액(원) | 증감액(원) |",
                "|------|------|------|------|------|------|",
                f"| 최초 | {_or(initial_date)} | {orig_start} ~ {orig_end} "
                f"| — | {_fmt(initial_amount) if initial_amount else '[확인 필요]'} | — |",
            ]
            for ch in sub_changes:
                ext_days = ch.get("extension_days")
                ext_str  = f"+{ext_days}일" if ext_days else "—"
                lines.append(
                    f"| {ch.get('seq', '')}차 변경 | {_or(ch.get('date'))} "
                    f"| — ~ {_or(ch.get('new_end_date'))} "
                    f"| {ext_str} "
                    f"| {_fmt(ch.get('amount'))} "
                    f"| {_fmt(ch.get('delta_amount'))} |"
                )
            if not sub_changes:
                lines.append(
                    "> ※ 변경계약 이력 — [확인 필요] (하도급 변경계약서 자료 미입력)\n"
                )
            lines.append("")

    # 제4절 과업 수행 절차
    lines += [
        "## 제4절 과업 수행 절차\n",
        "본 과업은 아래의 절차에 따라 수행되었다.\n",
        "| 단계 | 내용 |",
        "|------|------|",
        "| 1단계 | 기초자료 조사 및 검토 — 계약서, 변경계약서, 공사원가계산서, 공기연장 검토서, "
        "급여명세서, 경비 영수증·세금계산서 등 수집 및 검토 |",
        "| 2단계 | 발생 사유 및 청구 계약 근거 검토 — 공기지연 귀책사유 분석, 관련 법령 및 "
        "공사계약 일반조건 검토 |",
        "| 3단계 | 사실관계 분석 및 적정 추가공사비 산정 — 간접노무비·경비·일반관리비·이윤 항목별 실비 산정 |",
        "| 4단계 | 보고서 작성 — 산정 결과 정리 및 증빙자료 첨부 |",
        "",
    ]
    return "\n".join(lines)


def _section2_reason(data: dict) -> str:
    report_type    = data.get("report_type", "A")
    laws_from_data = data.get("laws", [])
    correspondence = data.get("correspondence", [])
    changes        = data.get("changes", [])

    law_header    = laws_db.get_law_header(report_type)
    contract_laws = laws_db.get_contract_laws(report_type)
    tpl = laws_db.CONTRACT_NATURE.get(report_type, laws_db.CONTRACT_NATURE["C"])

    lines = [
        "---\n",
        "# 제3장 공기연장 귀책 분석 및 계약금액조정 신청 적정성 검토\n",
    ]

    # ── 제1절 공기지연의 귀책사유 분석 (사실관계 먼저) ───────────────────────
    lines += [
        "## 제1절 공기지연의 귀책사유 분석\n",
        "### 1.1 공기지연 관련 공문 이력\n",
        "| 일자 | 발신 → 수신 | 내용 | 문서번호 | 출처 |",
        "|------|------|------|------|------|",
    ]
    if correspondence:
        for co in correspondence:
            lines.append(
                f"| {_or(co.get('date'))} | {_or(co.get('direction'))} "
                f"| {_or(co.get('title'))} | {_or(co.get('doc_number', '—'))} "
                f"| {_or(co.get('source'))} |"
            )
    else:
        lines.append("| — | — | (분석 결과에 공문 이력 없음) | — | — |")
    lines.append("")

    lines += [
        "### 1.2 귀책사유 분류\n",
        "| 공기지연 사유 | 관련 근거 | 비용 부담자 |",
        "|------|------|------|",
    ]
    for row in laws_db.get_attribution_table(report_type):
        lines.append(
            f"| {row['cause']} | {row['basis']} | {row['burden']} |"
        )
    lines.append("")

    if changes:
        lines += [
            "### 1.3 변경계약 귀책사유\n",
            "| 차수 | 사유 | 출처 |",
            "|------|------|------|",
        ]
        for ch in changes:
            lines.append(
                f"| {ch.get('seq', '')}차 변경 | {_or(ch.get('reason'))} | {_or(ch.get('reason_source'))} |"
            )
        lines.append("")

    lines.append(f"\n> **[결론]** {laws_db.get_claim_conclusion(report_type)}\n")

    # ── 제2절 청구의 근거 (법적 근거) ────────────────────────────────────────
    lines += [
        "## 제2절 청구의 근거\n",
        "### 2.1 계약에 의한 청구권\n",
        law_header + "\n",
        "**[공사계약 일반조건 근거 조문]**\n",
    ]
    for cl in tpl.get("claim_basis_contract", []):
        lines.append(_law_box(cl))

    lines += [
        "### 2.2 법령에 의한 청구권\n",
        "**[적용 법령 및 근거 조문]**\n",
    ]
    for law in contract_laws:
        lines.append(_law_box(law))

    ret_law = laws_db.INDIRECT_LABOR_BASIS["retirement"]
    lines.append(_law_box(ret_law))
    for ins in laws_db.INSURANCE_LAWS:
        ins_with_name = dict(ins, name=ins["law"])
        lines.append(_law_box(ins_with_name))
    if data.get("vat_zero_rated"):
        vat_source = data.get("vat_zero_rated_source", "관련 법령")
        lines += [
            f"> **■ 영세율(0%) 적용** *(출처: {vat_source})*\n>\n"
            "> 본 공사는 영세율 적용 대상으로 부가가치세를 면제한다.\n",
        ]
    else:
        lines.append(_law_box(laws_db.VAT_BASIS))

    # 현장 추가 법령
    all_db_names = (
        {l["name"] for l in contract_laws}
        | {ins["law"] for ins in laws_db.INSURANCE_LAWS}
        | {ret_law["law"], laws_db.VAT_BASIS["law"]}
    )
    extra_laws = [l for l in laws_from_data if l.get("name") and l["name"] not in all_db_names]
    if extra_laws:
        lines.append("**[현장 문서 추가 인용 법령]**\n")
        for i, law in enumerate(extra_laws, 1):
            lines.append(
                f"{i}. 「{_or(law.get('name'))}」{_or(law.get('article'))} "
                f"— {_or(law.get('purpose'))} *(출처: {_or(law.get('source'))})*"
            )
        lines.append("")

    lines += [
        "### 2.3 사정변경의 원칙\n",
        "계약 체결 이후 예측하지 못한 사정의 변경으로 인하여 당사자 일방에게 현저한 불이익이 "
        "발생한 경우, 민법상 사정변경의 원칙(clausula rebus sic stantibus)에 의하여 "
        "계약 내용의 변경 또는 이에 상응하는 추가 비용의 청구가 가능합니다.\n",
        "본건 공사는 착공 이후 발주자의 계획 변경, 설계 변경, 인허가 지연 등 "
        "계약 성립 당시 예측하지 못한 사정의 변경이 발생하여 공사기간이 연장되었으므로, "
        "이에 따른 간접비용을 청구하는 것은 사정변경의 원칙에도 부합합니다.\n",
    ]

    precedent = tpl.get("precedent", "")
    lines += [
        "### 2.4 청구의 근거 검토 (종합)\n",
        "이상과 같이 본건 공기연장 간접비 청구는 계약에 의한 청구권(공사계약 일반조건), "
        "법령에 의한 청구권(관련 법령), 사정변경의 원칙 등 다양한 근거에 의하여 정당화된다.\n",
    ]
    if precedent:
        lines += [
            "**[관련 판례]**\n",
            f"> {precedent}\n",
        ]

    return "\n".join(lines)


def _section_contract_nature(data: dict) -> str:
    """제2장 계약의 성격 — laws_db.CONTRACT_NATURE 템플릿 기반 생성"""
    report_type = data.get("report_type", "A")
    c = data.get("contract", {})
    subcontractors = data.get("subcontractors", [])
    tpl = laws_db.CONTRACT_NATURE.get(report_type, laws_db.CONTRACT_NATURE["C"])

    law_full  = tpl["law_system_full"]
    law_short = tpl["law_system_short"]
    long_term_law = tpl["long_term_law"]
    long_term_art = tpl["long_term_article"]
    contract_type = _or(c.get("contract_type"))

    # ── 제1절 적용 법령 체계 ───────────────────────────────────────────────────
    law_table_rows = "\n".join(
        f"| {l['article']} | {l.get('purpose', '')} |"
        for l in laws_db.get_contract_laws(report_type)
    )
    # 집행기준 행 추가
    gc_ref = tpl["general_conditions_ref"]
    gc_art = tpl["general_conditions_claim_articles"]

    # ── 제2절 계약문서 구성 및 주요 용어 ───────────────────────────────────────
    term_rows = "\n".join(
        f"| {term} | {defn} |"
        for term, defn in tpl["key_terms"].items()
    )

    # ── 제3절 계약당사자 현황 ───────────────────────────────────────────────────
    sub_names = ", ".join(s.get("name", "") for s in subcontractors) if subcontractors else "없음"
    supervisor = _or(c.get("supervisor"))          # analysis_result에 없으면 확인 필요

    # ── 제4절 장기계속공사 / 계속비공사 ────────────────────────────────────────
    long_desc = tpl["long_term_desc"]

    # ── 제5절 내역입찰 공사의 특성 ─────────────────────────────────────────────
    rates = data.get("rates", {})
    ia_r   = _pct(rates.get("industrial_accident"))
    emp_r  = _pct(rates.get("employment"))
    ga_r   = _pct(rates.get("general_admin"))
    prf_r  = _pct(rates.get("profit"))
    ia_src  = _or(rates.get("industrial_accident_source"))
    emp_src = _or(rates.get("employment_source"))
    ga_src  = _or(rates.get("general_admin_source"))
    prf_src = _or(rates.get("profit_source"))

    lines = [
        "---\n",
        "# 제2장 계약의 성격\n",
        "## 제1절 적용 법령 체계\n",
        f"본건 공사의 발주자는 {tpl['client_type']}으로서, 본건 계약은 "
        f"「{law_full}」(이하 '{law_short}') 체계를 적용한다.\n",
        "| 법령 | 적용 조항 | 내용 |",
        "|------|------|------|",
    ]
    for l in laws_db.get_contract_laws(report_type):
        lines.append(
            f"| {l['name']} | {l['article']} | {l.get('purpose', '')} |"
        )
    lines += [
        f"| {gc_ref} | {gc_art} | 공기연장 간접비 산정 기준 |",
        "",
        "## 제2절 계약문서 구성 및 주요 용어 정의\n",
        f"### 계약문서 구성 ({tpl['contract_doc_article']})\n",
        "본건 공사의 계약문서는 계약서, 설계서(도면·시방서·현장설명서 등), "
        "산출내역서, 공사계약 일반조건, 특수조건 등으로 구성된다.\n",
        "### 주요 용어 정의\n",
        "| 용어 | 정의 |",
        "|------|------|",
        term_rows,
        "",
        "## 제3절 계약당사자 현황\n",
        "| 구 분 | 내 용 |",
        "|------|------|",
        f"| 발주자 | {_or(c.get('client'))} |",
        f"| 계약상대자 | {_or(c.get('contractor'))} |",
        f"| 건설사업관리단 | {supervisor} |",
        f"| 하도급사 | {sub_names} |",
        "",
        f"## 제4절 {contract_type}의 성격\n",
        f"본건 공사는 「{long_term_law}」 {long_term_art}에 따른 **{contract_type}**으로 체결되었다.\n",
        long_desc + "\n",
        "## 제5절 내역입찰 공사의 특성\n",
        "본건 공사는 **내역입찰** 방식으로 계약이 체결되었다. "
        "「(계약예규) 정부 입찰·계약 집행기준」 제18조에 의해 입찰서에 산출내역서를 첨부하여 "
        "입찰 후 낙찰자로 선정되어 최초 계약을 체결한 공사도급계약에 해당한다. "
        "이에 따라 산출내역서(공사원가계산서)에 기재된 요율이 계약 내용의 일부를 구성하므로, "
        "공기연장 간접비 산정 시 해당 요율을 기준으로 적용한다.\n",
        "| 비목 | 적용 요율 | 출처 |",
        "|------|------|------|",
        f"| 산재보험료율 | {ia_r} | {ia_src} |",
        f"| 고용보험료율 | {emp_r} | {emp_src} |",
        f"| 일반관리비율 | {ga_r} | {ga_src} |",
        f"| 이윤율 | {prf_r} | {prf_src} |",
        "",
    ]
    return "\n".join(lines)


def _section3_method(data: dict, calc: dict) -> str:
    report_type = data.get("report_type", "A")
    ext = data.get("extension", {})
    rates = data.get("rates", {})
    admin_r = calc.get("general_admin", {})
    profit_r = calc.get("profit", {})
    subcontractors = data.get("subcontractors", [])
    sub_results = calc.get("subcontractor_results", [])

    if report_type == "A":
        law_ref = ("「지방자치단체를 당사자로 하는 계약에 관한 법률」제22조, 동법 시행령 제75조, "
                   "「지방자치단체 입찰 및 계약집행기준」제1장 제7절")
    elif report_type == "B":
        law_ref = ("「국가를 당사자로 하는 계약에 관한 법률」, 동법 시행령, "
                   "「정부 입찰·계약 집행기준」제13장 제73조")
    else:
        law_ref = "「공사계약일반조건」제12조, 「정부 입찰·계약 집행기준」제73조"

    # Section 3.1 narrative
    n_changes = len(data.get("changes", []))
    initial_end = data.get("contract", {}).get("initial_end", "확인 필요")
    final_end   = ext.get("end_date", "확인 필요")
    total_days  = ext.get("total_days", "확인 필요")

    if n_changes > 0:
        narrative31 = (
            f"본건 공사는 총 {n_changes}회의 변경계약이 있었고, "
            f"준공일을 당초 {initial_end}에서 {final_end}로 {total_days}일 연장하였습니다. "
            f"따라서 공기연장 일수는 변경계약으로 합의한 {total_days}일로 보는 것이 타당합니다."
        )
    else:
        narrative31 = (
            f"본건 공사의 공기연장 기간은 **{_or(ext.get('start_date'))} ~ {final_end}** "
            f"({total_days}일)입니다. *(출처: {_or(ext.get('source'))})*"
        )

    lines = [
        "---\n",
        "# 제4장 공기연장 간접비 산정\n",
        "## 제1절 공기연장 기간의 산정\n",
        "### 1.1 공기연장 일수의 산정방식\n",
        narrative31 + "\n",
        "### 1.2 공기연장 일수 산정\n",
        "| 구분 | 계약일 | 계약금액(원) | 공사기간 | 연장일수 | 사유 |",
        "|------|------|------|------|------|------|",
        f"| 최초 | {_or(data.get('contract', {}).get('initial_date'))} "
        f"| {_fmt(data.get('contract', {}).get('initial_amount'))} "
        f"| {_or(data.get('contract', {}).get('initial_start'))} ~ "
        f"{_or(data.get('contract', {}).get('initial_end'))} | — | — |",
    ]
    for ch in data.get("changes", []):
        lines.append(
            f"| {ch.get('seq', '')}차 변경 | {_or(ch.get('date'))} | {_fmt(ch.get('amount'))} "
            f"| — ~ {_or(ch.get('new_end_date'))} | 증 {_or(ch.get('extension_days'))}일 "
            f"| {_or(ch.get('reason'))} |"
        )
    lines += [
        f"\n> **총 연장일수**: {total_days}일 "
        f"({_or(ext.get('start_date'))} ~ {final_end})\n",
    ]

    # 하도급사별 연장 기간 — 하도급사별 개별 항목 + 산문 설명 (레퍼런스 형식)
    if subcontractors:
        sr_map = {sr.get("name", ""): sr for sr in sub_results} if sub_results else {}
        for s in subcontractors:
            name   = _or(s.get("name"))
            orig_s = s.get("period_start", "")
            orig_e = s.get("period_end", "")
            sr     = sr_map.get(s.get("name", ""), {})
            calc_s = sr.get("period_start", orig_s) if sr else orig_s
            calc_e = sr.get("period_end", orig_e)   if sr else orig_e
            # 일수 계산
            from datetime import datetime as _dt
            try:
                days = (_dt.strptime(calc_e, "%Y-%m-%d") - _dt.strptime(calc_s, "%Y-%m-%d")).days + 1
            except Exception:
                days = "확인 필요"
            prose = _sub_period_prose(report_type, name, calc_s, calc_e, days, orig_s, orig_e)
            lines += [
                f"#### {name}\n",
                prose + "\n",
            ]

    # 보험료 법령 DB에서 가져오기
    ins_map = {ins["item"]: ins for ins in laws_db.INSURANCE_LAWS}
    ret_law = laws_db.INDIRECT_LABOR_BASIS["retirement"]
    profit_law = laws_db.PROFIT_BASIS

    # 일반관리비 법령: B타입은 계약금액 기준으로 B_large/B_small 분기
    if report_type == "B":
        _amount = data.get("contract", {}).get("initial_amount", 0) or 0
        _ga_key = "B_large" if _amount >= _cfg.LARGE_CONTRACT_THRESHOLD else "B_small"
        admin_law = laws_db.GENERAL_ADMIN_BASIS.get(_ga_key, laws_db.GENERAL_ADMIN_BASIS["C"])
    else:
        admin_law = laws_db.GENERAL_ADMIN_BASIS.get(report_type, laws_db.GENERAL_ADMIN_BASIS["C"])

    section32_intro = laws_db.get_section32_intro(report_type)
    lines += [
        "## 제2절 간접비 산정방식\n",
        section32_intro + "\n",

        "**① 간접노무비**\n",
        "간접노무비란 공사 현장에서 직접 공사에 종사하지 아니하고 현장의 관리·감독·행정 업무에 "
        "종사하는 간접 인원에게 지급하는 임금을 말한다. 공기연장 기간 중 현장에 상주하여 공사를 "
        "관리·감독한 인원의 실제 지급 임금을 기준으로 산정하며, 현장소장·공무·안전·품질·측량 "
        "등 직접공사비에 계상되지 않는 관리 인원이 대상이 된다.\n",
        "| 항목 | 산정 기준 |",
        "|------|------|",
        "| 기본급 | 월 급여 ÷ 30일 × 연장일수 |",
        f"| 퇴직급여충당금 | 기본급 × {ret_law['rate']} "
        f"(「{ret_law['law']}」 {ret_law['article']}) |",
        "| 합계 | 기본급 + 퇴직급여충당금 |",
        "",
        "> ※ 준공일 이전 기간은 실비 산정, 준공일 이후 잔여 기간은 실비 기준 추정 산정\n",
    ]
    _sp = laws_db.STANDARD_PRICE_ADDITIONAL
    _sp_labor = _sp.get("제2장제3절제18조", {})
    if _sp_labor:
        lines.append(_law_box(_sp_labor))
    _sp_exp = _sp.get("제2장제3절제19조", {})
    lines += [
        "**② 경비**\n",
        "경비는 공사 현장의 운영·유지에 직접 소요되는 비용으로, 직접계상비목과 승률계상비목으로 "
        "구분하여 산정한다.\n",
    ]
    if _sp_exp:
        lines.append(_law_box(_sp_exp))
    lines += [
        "- **직접계상비목** (영수증·세금계산서 기준 실지급액)\n",
        "  - 전력비·수도광열비: 현장 가설사무실, 숙소, 창고 등에서 발생한 전력요금, 수도요금, 난방비\n",
        "  - 여비·교통비·통신비: 시공현장에서 직접 소요되는 여비, 차량유지비, 전신전화사용료, 우편료\n",
        "  - 지급임차료: 계약목적물 시공에 직접 사용되는 토지, 건물, 기계기구(건설기계 제외)의 사용료\n",
        "  - 지급수수료: 법률로 규정되거나 의무 지워진 수수료 (다른 비목에 미계상분)\n",
        "  - 도서인쇄비: 참고서적구입비, 인쇄비, 사진제작비, 공사시공기록책자 제작비 등\n",
        "  - 세금과공과: 당해 공사와 직접 관련된 재산세, 차량세 등 세금 및 공공단체 납부 공과금\n",
        "  - 복리후생비: 노무자·현장사무소직원 등의 의료위생약품대, 지급피복비, 급식비 등 작업조건 유지 비용\n",
        "  - 소모품비: 작업현장에서 발생하는 문방구, 장부대 등 소모용품\n",
        f"  - 국민연금 (사용자 부담분): 간접노무비 × {_pct(rates.get('national_pension', 0.045))} "
        f"(「{ins_map['국민연금']['law']}」 {ins_map['국민연금']['article']})\n",
        f"  - 건강보험 (사용자 부담분): 간접노무비 × {_pct(rates.get('health_insurance', 0.03545))} "
        f"(「{ins_map['건강보험']['law']}」 {ins_map['건강보험']['article']})\n",
        f"  - 노인장기요양보험: 건강보험료 × {_pct(rates.get('long_term_care', 0.1295))} "
        f"(「{ins_map['노인장기요양보험']['law']}」 {ins_map['노인장기요양보험']['article']})\n",
        "- **승률계상비목** (간접노무비 대비 요율 적용)\n",
        f"  - 산재보험료: 간접노무비 × {_pct(rates.get('industrial_accident'))} "
        f"(「{ins_map['산재보험료']['law']}」 {ins_map['산재보험료']['article']}, "
        f"출처: {_or(rates.get('industrial_accident_source'))})\n",
        f"  - 고용보험료: 간접노무비 × {_pct(rates.get('employment'))} "
        f"(「{ins_map['고용보험료']['law']}」 {ins_map['고용보험료']['article']}, "
        f"출처: {_or(rates.get('employment_source'))})\n",
        "> ※ **직접계상비목 관련 판례**: 서울중앙지법 2014. 12. 10. 선고 2012가합80465 판결 — "
        "간접공사비 산정 방식의 합의가 없는 경우, 실제 지출한 공사비용 중 공사기간 연장과 "
        "객관적 관련성이 있고 상당한 범위 안에서 간접공사비를 산정한다.\n",
        "**[보증수수료 산정방식]**\n",
        "보증수수료는 「(계약예규) 정부 입찰·계약 집행기준」 제73조 제4항에 따라 "
        "계약상대자로부터 제출받은 보증수수료의 영수증 등 객관적인 자료에 의하여 "
        "확인된 금액을 기준으로 산출한다.\n",
    ]

    # 하도급사 경비 요율 설명 (계산 결과의 실제 적용 요율 사용)
    if subcontractors and sub_results:
        lines += ["**[하도급사 경비 적용 요율]**\n"]
        for sr in sub_results:
            ia  = sr.get("ia_rate",  0.0)
            emp = sr.get("emp_rate", 0.0)
            lines.append(
                f"- {sr.get('name', '')} — "
                f"산재보험료: 간접노무비 × {_pct(ia)}, "
                f"고용보험료: 간접노무비 × {_pct(emp)}\n"
            )
        lines.append(
            "> ※ 하도급사의 경우 산재보험료는 원도급사가 일괄 계상하는 특성으로 인하여 "
            "별도 계상하지 않으므로 산정 제외.\n"
        )

    _sp = laws_db.STANDARD_PRICE_ADDITIONAL
    lines += [
        "**③ 일반관리비**\n",
        "일반관리비란 기업의 유지를 위한 관리활동 부문에서 발생하는 제비용으로, 임원급여·사무직원 "
        "급여·사무용품비·통신비 등 현장 공사원가에 포함되지 않는 본사 운영비용을 의미한다. "
        "공기연장 간접비에 있어서는 (간접노무비 + 경비) 소계에 아래 요율을 적용하여 산정한다.\n",
        f"- **적용 법령**: 「{admin_law['law']}」 {admin_law['article']}\n",
        f"- **요율 한도**: {admin_law['limit']}\n",
        f"- **적용 요율**: {_pct(admin_r.get('rate'))} (한도 이내 적용)\n",
        f"- **산정 방식**: (간접노무비 + 경비) × {_pct(admin_r.get('rate'))}\n",
    ]
    lines.append(_law_box(admin_law))
    if admin_r.get("note"):
        lines.append(f"- ⚠️ {admin_r['note']}\n")

    lines += [
        "**④ 이윤**\n",
        "이윤은 공사 수행에 따른 정당한 수익을 의미하며, 공기연장 간접비 산정 시 "
        "(간접노무비 + 경비 + 일반관리비) 합계에 요율을 적용하여 산정한다. "
        "다만 본건에서는 계약 내역서상 이윤율을 기준으로 적용하되, 법령 한도를 초과하지 않는 범위에서 산정한다.\n",
        f"- **적용 법령**: 「{profit_law['law']}」 {profit_law['article']}\n",
        f"- **요율 한도**: {profit_law['limit']}\n",
        f"- **적용 요율**: {_pct(profit_r.get('rate'))}\n",
        f"- **산정 방식**: (간접노무비 + 경비 + 일반관리비) × {_pct(profit_r.get('rate'))}\n",
    ]
    lines.append(_law_box(profit_law))
    if profit_r.get("note"):
        lines.append(f"- ⚠️ {profit_r['note']}\n")

    _vat18 = _sp.get("제18조_vat", {})
    if data.get("vat_zero_rated"):
        lines += [
            "**⑤ 부가가치세 (영세율)**\n",
            "본 공사는 영세율 적용 대상으로 부가가치세 세율 0%를 적용한다.\n",
            f"- **적용 근거**: {_or(data.get('vat_zero_rated_source', '부가가치세법 제11조 등 영세율 규정'))}\n",
            "- **세율**: 영세율(0%)\n",
            "- **산정 방식**: 총원가 × 0% = 0원\n",
        ]
        if _vat18:
            lines.append(_law_box(_vat18))
    else:
        lines += [
            "**⑤ 부가가치세**\n",
            "부가가치세는 공급가액에 대하여 10%의 세율을 적용한다.\n",
            f"- **적용 법령**: 「{laws_db.VAT_BASIS['law']}」 {laws_db.VAT_BASIS['article']}\n",
            f"- **세율**: {laws_db.VAT_BASIS['rate']}\n",
        ]
        lines.append(_law_box(laws_db.VAT_BASIS))

    return "\n".join(lines)


def _section4_result(data: dict, calc: dict) -> str:
    labor = calc.get("indirect_labor", {})
    expenses = calc.get("expenses", {})
    subtotal = calc.get("subtotal", {})
    admin = calc.get("general_admin", {})
    profit = calc.get("profit", {})
    grand = calc.get("grand_total", {})
    vat = calc.get("vat", 0)
    final = calc.get("final_rounded", 0)

    # 직접계상 합계 (영수증 항목 + 3대보험)
    total_direct_act = (expenses.get("total_direct_actual") or 0)
    total_direct_est = (expenses.get("total_direct_estimated") or 0)
    total_direct_tot = total_direct_act + total_direct_est

    ext = data.get("extension", {})
    kr_ext_s = _fmt_date_kr(_or(ext.get("start_date")))
    kr_ext_e = _fmt_date_kr(_or(ext.get("end_date")))
    ext_days = _or(ext.get("total_days"))

    lines = [
        "---\n",
        "## 제3절 원도급사 간접비 산정 결과\n",
        "### 3.1 집계표\n",
        f"본건 공사의 공기연장 기간인 {kr_ext_s}부터 {kr_ext_e}까지 총 {ext_days}일에 "
        "발생한 간접비에 대하여 산정하였습니다.\n",
        "*(단위: 원)*\n",
        "| 구 분 | A. 실비 금액 | B. 추정 금액 | C. 합계(A+B) | 비 고 |",
        "|------|------|------|------|------|",
        f"| 1. 간접노무비 | {_fmt(labor.get('actual'))} | {_fmt(labor.get('estimated'))} "
        f"| {_fmt(labor.get('total'))} | |",
        f"| 2. 경비 소계 | {_fmt(expenses.get('actual'))} | {_fmt(expenses.get('estimated'))} "
        f"| {_fmt(expenses.get('total'))} | |",
        f"|  가. 직접계상비목 | {_fmt(total_direct_act)} | {_fmt(total_direct_est)} "
        f"| {_fmt(total_direct_tot)} | |",
    ]

    # 직접계상 세부: 영수증 항목 (비목명 합산 적용)
    merged_direct_rows = _merge_direct_rows(expenses.get("direct_rows", []))
    for row in merged_direct_rows:
        lines.append(
            f"|   ① {row['item']} | {_fmt(row.get('actual'))} | {_fmt(row.get('estimated'))} "
            f"| {_fmt(row.get('total'))} | |"
        )

    # 직접계상 세부: 3대보험
    for ins in expenses.get("insurance_direct_rows", []):
        rb = ins.get("rate_base", "간접노무비")
        lines.append(
            f"|   ② {ins['item']} | {_fmt(ins.get('actual'))} | {_fmt(ins.get('estimated'))} "
            f"| {_fmt(ins.get('total'))} | ×{_pct(ins.get('rate'))}({rb}) |"
        )

    # 승률계상 항목
    for rr in expenses.get("rate_rows", []):
        lines.append(
            f"|  나. {rr['item']} | {_fmt(rr.get('actual'))} | {_fmt(rr.get('estimated'))} "
            f"| {_fmt(rr.get('total'))} | ×{_pct(rr.get('rate'))} |"
        )

    lines += [
        f"| 3. 소계 (1+2) | {_fmt(subtotal.get('actual'))} | {_fmt(subtotal.get('estimated'))} "
        f"| {_fmt(subtotal.get('total'))} | |",
        f"| 4. 일반관리비 | {_fmt(admin.get('actual'))} | {_fmt(admin.get('estimated'))} "
        f"| {_fmt(admin.get('total'))} | ×{_pct(admin.get('rate'))} |",
        f"| 5. 이윤 | {_fmt(profit.get('actual'))} | {_fmt(profit.get('estimated'))} "
        f"| {_fmt(profit.get('total'))} | ×{_pct(profit.get('rate'))} |",
        f"| 6. 총원가 (3+4+5) | {_fmt(grand.get('actual'))} | {_fmt(grand.get('estimated'))} "
        f"| {_fmt(grand.get('total'))} | |",
    ]
    jv_ratio       = calc.get("jv_share_ratio", 1.0)
    grand_after_jv = calc.get("grand_after_jv", grand.get("total", 0))
    if jv_ratio < 1.0:
        lines.append(
            f"| 6-1. JV 지분율 적용 ({jv_ratio:.0%}) | | | {_fmt(grand_after_jv)} | ×{_pct(jv_ratio)} |"
        )
    lines += [
        f"| 7. 부가가치세 ({'영세율(0%)' if data.get('vat_zero_rated') else '10%'}) | | | {_fmt(vat)} | |",
        f"| **합 계 (천원 미만 절사)** | | | **{_fmt(final)}** | |",
        "",
    ]

    # 4.2 간접노무비 상세
    has_estimated = any(r.get("estimated_days", 0) > 0 for r in labor.get("rows", []))
    labor_target_note = (
        "계약상대자가 제출한 경력증명서 및 현장조직도 등의 자료를 검토하여 "
        "간접노무인원의 간접비 산정 대상 기간을 특정하였으며, 세부 검토 내용은 아래와 같습니다."
    )
    labor_salary_note = "간접노무비를 산정하기 위해 계약상대자가 증빙자료로 제출한 급여명세서 등의 자료를 기준으로 산정하였습니다."
    if has_estimated:
        labor_salary_note += (
            " 변경된 준공일이 현재 도래하지 않았으므로, 일부 기간에 대하여 "
            "실비 산정 금액을 기준으로 추정산출을 하였습니다."
        )
    lines += [
        "### 3.2 간접노무비 상세\n",
        "#### ① 대상인원 현황\n",
        f"> {labor_target_note}\n",
        "| No. | 소속 | 이름 | 직무 | 산정기간 | 실비일수 | 추정일수 | 출처 |",
        "|------|------|------|------|------|------|------|------|",
    ]
    for i, row in enumerate(labor.get("rows", []), 1):
        total_days_row = row.get("actual_days", 0) + row.get("estimated_days", 0)
        lines.append(
            f"| {i} | {_or(row.get('org'))} | {_or(row.get('name'))} | {_or(row.get('role'))} "
            f"| {_or(row.get('period'))} | {row.get('actual_days', 0)} | {row.get('estimated_days', 0)} "
            f"| {_or(row.get('source'))} |"
        )
    lines += [
        "",
        "#### ② 급여 산정내역\n",
        f"> {labor_salary_note}\n",
        "| No. | 소속 | 이름 | 직무 | 월급여(원) | A. 급여(원) | B. 퇴직충당금(원) | C. 소계(A+B) | 실비(원) | 추정(원) | 합계(원) |",
        "|------|------|------|------|------|------|------|------|------|------|------|",
    ]
    for i, row in enumerate(labor.get("rows", []), 1):
        salary_total  = row.get("act_salary", 0) + row.get("est_salary", 0)
        retire_total  = row.get("act_retire", 0) + row.get("est_retire", 0)
        lines.append(
            f"| {i} | {_or(row.get('org'))} | {_or(row.get('name'))} | {_or(row.get('role'))} "
            f"| {_fmt(row.get('monthly_salary'))} | {_fmt(salary_total)} "
            f"| {_fmt(retire_total)} | {_fmt(row.get('total'))} "
            f"| {_fmt(row.get('actual'))} | {_fmt(row.get('estimated'))} | {_fmt(row.get('total'))} |"
        )
    total_salary = sum(r.get("act_salary", 0) + r.get("est_salary", 0) for r in labor.get("rows", []))
    total_retire = sum(r.get("act_retire", 0) + r.get("est_retire", 0) for r in labor.get("rows", []))
    lines += [
        f"| | | | **소계** | | {_fmt(total_salary)} | {_fmt(total_retire)} "
        f"| {_fmt(labor.get('total'))} | {_fmt(labor.get('actual'))} "
        f"| {_fmt(labor.get('estimated'))} | {_fmt(labor.get('total'))} |",
        "",
        "> ※ 실비: 보고서 작성 기준일 이전 지급 완료 금액 / 추정: 이후 지급 예정 금액\n",
    ]

    # 4.3 경비 상세
    merged_direct = _merge_direct_rows(expenses.get("direct_rows", []))
    lines += [
        "### 3.3 경비 상세\n",
        "#### 가. 직접계상비목\n",
        "| 항목 | 실비(원) | 추정(원) | 합계(원) | 산정기준 | 출처 |",
        "|------|------|------|------|------|------|",
    ]
    for row in merged_direct:
        lines.append(
            f"| {_or(row.get('item'))} | {_fmt(row.get('actual'))} | {_fmt(row.get('estimated'))} "
            f"| {_fmt(row.get('total'))} | 영수증·세금계산서 | {_or(row.get('source'))} |"
        )
    # 3대보험도 직접계상비목 표에 추가
    for ins in expenses.get("insurance_direct_rows", []):
        rb = ins.get("rate_base", "간접노무비")
        lines.append(
            f"| {ins['item']} | {_fmt(ins.get('actual'))} | {_fmt(ins.get('estimated'))} "
            f"| {_fmt(ins.get('total'))} | ×{_pct(ins.get('rate'))}({rb}) | {_or(ins.get('source'))} |"
        )
    # 승률계상비목 intro (type-aware)
    report_type_for_rate = data.get("report_type", "A")
    ia_r_str  = _pct(data.get("rates", {}).get("industrial_accident"))
    emp_r_str = _pct(data.get("rates", {}).get("employment"))
    if report_type_for_rate == "A":
        rate_intro = (
            "「지방자치단체 입찰 및 계약집행기준」제7절 2항 공사이행기간의 변경에 따른 실비 산정기준을 "
            f"따르도록 명시되어 있습니다. 이에 당초 산출내역서상 해당 비목의 요율로써 "
            f"산재보험료는 {ia_r_str}를, 고용보험료는 {emp_r_str}를 반영하였습니다."
        )
    elif report_type_for_rate == "B":
        rate_intro = (
            "「(계약예규) 정부 입찰·계약 집행기준」 제16장(실비의 산정)에 따른 산정기준을 "
            f"따르도록 명시되어 있습니다. 이에 당초 산출내역서상 해당 비목의 요율로써 "
            f"산재보험료는 {ia_r_str}를, 고용보험료는 {emp_r_str}를 반영하였습니다."
        )
    else:
        rate_intro = (
            f"당초 산출내역서상 해당 비목의 요율로써 "
            f"산재보험료는 {ia_r_str}를, 고용보험료는 {emp_r_str}를 반영하였습니다."
        )
    lines += [
        "",
        "#### 나. 승률계상비목\n",
        f"> {rate_intro}\n",
        "| 항목 | 요율 | 실비(원) | 추정(원) | 합계(원) | 출처 |",
        "|------|------|------|------|------|------|",
    ]
    for rr in expenses.get("rate_rows", []):
        lines.append(
            f"| {rr['item']} | {_pct(rr.get('rate'))} | {_fmt(rr.get('actual'))} "
            f"| {_fmt(rr.get('estimated'))} | {_fmt(rr.get('total'))} | {_or(rr.get('source'))} |"
        )
    lines.append("")

    # 4.4 일반관리비·이윤
    lines += [
        "### 3.4 일반관리비 및 이윤\n",
        f"- **일반관리비** = 소계 × {_pct(admin.get('rate'))} "
        f"= {_fmt(subtotal.get('actual'))} × {_pct(admin.get('rate'))} + "
        f"{_fmt(subtotal.get('estimated'))} × {_pct(admin.get('rate'))} "
        f"= **{_fmt(admin.get('total'))}원**\n",
        f"- **이윤** = (소계 + 일반관리비) × {_pct(profit.get('rate'))} "
        f"= **{_fmt(profit.get('total'))}원**\n",
        (f"- **부가가치세** = 총원가 × 영세율(0%) = **0원**\n"
         if data.get("vat_zero_rated") else
         f"- **부가가치세** = 총원가 × 10% = {_fmt(grand.get('total'))} × 10% "
         f"= **{_fmt(vat)}원**\n"),
    ]

    return "\n".join(lines)


def _section5_subcontractor(data: dict, calc: dict) -> str:
    """5장: 하도급사 간접비 산정 결과"""
    sub_results = calc.get("subcontractor_results", [])
    subcontractors = data.get("subcontractors", [])
    report_type = data.get("report_type", "A")

    lines = [
        "---\n",
        "## 제4절 하도급사 간접비 산정 결과\n",
    ]

    if not subcontractors:
        lines.append("> 하도급사가 없습니다.\n")
        return "\n".join(lines)

    # 5.1 집계표 (하도급사별 합계표)
    lines += [
        "### 4.1 집계표\n",
        "본건 공사의 공기연장 기간 동안 하도급사별로 발생한 간접비에 대하여 산정하였으며, "
        "산정한 결과는 아래의 표와 같습니다.\n",
        "*(단위: 원)*\n",
        "| 하도급사 | 간접노무비 | 경비 | 소계 | 일반관리비 | 이윤 | 합계 | 비고 |",
        "|------|------|------|------|------|------|------|------|",
    ]
    grand_sub_total = 0
    for sr in sub_results:
        agg = sr.get("aggregate")
        name = sr.get("name", "")
        if agg:
            labor_tot   = agg.get("indirect_labor", {}).get("total", 0)
            exp_tot     = agg.get("expenses", {}).get("total", 0)
            sub_tot     = agg.get("subtotal", {}).get("total", 0)
            admin_tot   = agg.get("general_admin", {}).get("total", 0)
            profit_tot  = agg.get("profit", {}).get("total", 0)
            grand_tot   = agg.get("grand_total", {}).get("total", 0)
            grand_sub_total += grand_tot
            lines.append(
                f"| {name} | {_fmt(labor_tot)} | {_fmt(exp_tot)} | {_fmt(sub_tot)} "
                f"| {_fmt(admin_tot)} | {_fmt(profit_tot)} | {_fmt(grand_tot)} | |"
            )
        else:
            lines.append(f"| {name} | [확인 필요] | [확인 필요] | — | — | — | — | 데이터 미입력 |")
    lines += [
        f"| **합 계** | | | | | | **{_fmt(grand_sub_total)}** | |",
        "",
    ]

    # 5.2 간접노무비 (하도급사별)
    lines += [
        "### 4.2 간접노무비\n",
        "#### 대상인원\n",
        "계약상대자가 제출한 개인이력카드 및 현장조직도 등의 자료를 검토하여 "
        "간접노무인원의 간접비 산정 대상 기간을 특정하였으며, 세부 검토 내용은 아래와 같습니다.\n",
    ]
    for sr in sub_results:
        name    = sr.get("name", "")
        labor   = sr.get("labor")
        p_start = sr.get("period_start", "")
        p_end   = sr.get("period_end", "")
        from datetime import datetime as _dt2
        try:
            days_sub = (_dt2.strptime(p_end, "%Y-%m-%d") - _dt2.strptime(p_start, "%Y-%m-%d")).days + 1
        except Exception:
            days_sub = "확인 필요"
        kr_ps = _fmt_date_kr(p_start)
        kr_pe = _fmt_date_kr(p_end)
        lines += [
            f"#### {name}\n",
            f"{name}은(는) {kr_ps}부터 {kr_pe}까지 {days_sub}일의 공기연장 간접비를 산정하였습니다.\n",
        ]
        if labor is None:
            lines.append("> [확인 필요] — 하도급사 직원 명단 및 급여 자료가 입력되지 않았습니다.\n")
            continue
        rows = labor.get("rows", [])
        if not rows:
            lines.append("> [확인 필요] — 간접노무비 인원 정보 없음\n")
            continue
        # 대상인원 현황
        lines += [
            "**① 대상인원 현황**\n",
            "| No. | 소속 | 이름 | 직무 | 산정기간 | 실비일수 | 추정일수 | 출처 |",
            "|------|------|------|------|------|------|------|------|",
        ]
        for i, row in enumerate(rows, 1):
            lines.append(
                f"| {i} | {_or(row.get('org'))} | {_or(row.get('name'))} | {_or(row.get('role'))} "
                f"| {_or(row.get('period'))} | {row.get('actual_days', 0)} | {row.get('estimated_days', 0)} "
                f"| {_or(row.get('source'))} |"
            )
        # 급여 산정내역
        lines += [
            "",
            "**② 급여 산정내역**\n",
            "> 간접노무비를 산정하기 위해 계약상대자가 증빙자료로 제출한 급여명세서 등의 자료를 기준으로 산정하였습니다.\n",
            "| No. | 소속 | 이름 | 직무 | 월급여(원) | A. 급여(원) | B. 퇴직충당금(원) | C. 소계(A+B) | 실비(원) | 추정(원) | 합계(원) |",
            "|------|------|------|------|------|------|------|------|------|------|------|",
        ]
        for i, row in enumerate(rows, 1):
            salary_total = row.get("act_salary", 0) + row.get("est_salary", 0)
            retire_total = row.get("act_retire", 0) + row.get("est_retire", 0)
            lines.append(
                f"| {i} | {_or(row.get('org'))} | {_or(row.get('name'))} | {_or(row.get('role'))} "
                f"| {_fmt(row.get('monthly_salary'))} | {_fmt(salary_total)} "
                f"| {_fmt(retire_total)} | {_fmt(row.get('total'))} "
                f"| {_fmt(row.get('actual'))} | {_fmt(row.get('estimated'))} | {_fmt(row.get('total'))} |"
            )
        sub_salary_total = sum(r.get("act_salary", 0) + r.get("est_salary", 0) for r in rows)
        sub_retire_total = sum(r.get("act_retire", 0) + r.get("est_retire", 0) for r in rows)
        lines += [
            f"| | | | **소계** | | {_fmt(sub_salary_total)} | {_fmt(sub_retire_total)} "
            f"| {_fmt(labor.get('total'))} | {_fmt(labor.get('actual'))} "
            f"| {_fmt(labor.get('estimated'))} | {_fmt(labor.get('total'))} |",
            "",
            "> 퇴직급여충당금 포함 금액\n",
        ]

    # 5.3 경비 (하도급사별)
    if report_type == "A":
        sub_rate_intro = (
            "「지방자치단체 입찰 및 계약집행기준」제7절 2항 공사이행기간의 변경에 따른 실비 산정기준을 "
            "따르도록 명시되어 있습니다. 하도급사의 승률계상비목 경비도 원도급사와 동일한 기준으로 "
            "산출내역서상 요율을 곱하여 계상하였으나 요율이 상이하여 원도급사와 다른 요율을 곱하여 계상하였습니다."
        )
    elif report_type == "B":
        sub_rate_intro = (
            "「(계약예규) 정부 입찰·계약 집행기준」 제16장(실비의 산정)에 따른 산정기준을 따르도록 "
            "명시되어 있습니다. 하도급사의 승률계상비목도 원도급사와 동일한 기준으로 산출내역서상 요율을 "
            "곱하여 계상하였으나 요율이 상이하여 원도급사와 다른 요율을 적용하였습니다."
        )
    else:
        sub_rate_intro = (
            "하도급사의 승률계상비목 경비는 산출내역서상 요율을 곱하여 계상하였습니다."
        )

    lines += [
        "### 4.3 경비\n",
        "#### 직접계상비목\n",
    ]
    for sr in sub_results:
        name     = sr.get("name", "")
        expenses = sr.get("expenses")
        ia_rate  = sr.get("ia_rate", 0.0)
        emp_rate = sr.get("emp_rate", 0.0062)
        lines.append(f"#### {name}\n")
        if expenses is None:
            lines.append("> [확인 필요] — 경비 자료가 입력되지 않았습니다.\n")
            continue
        # 직접계상
        direct_rows = expenses.get("direct_rows", [])
        if direct_rows:
            lines += [
                "| 항목 | 실비(원) | 추정(원) | 합계(원) | 출처 |",
                "|------|------|------|------|------|",
            ]
            for row in direct_rows:
                lines.append(
                    f"| {_or(row.get('item'))} | {_fmt(row.get('actual'))} | {_fmt(row.get('estimated'))} "
                    f"| {_fmt(row.get('total'))} | {_or(row.get('source'))} |"
                )
            lines.append("")
        else:
            lines.append("> [확인 필요] — 직접계상 경비 자료가 입력되지 않았습니다.\n")

    # 승률계상비목 — 공통 intro + 하도급사별
    lines += [
        "#### 승률계상비목\n",
        f"> {sub_rate_intro}\n",
    ]
    for sr in sub_results:
        name     = sr.get("name", "")
        expenses = sr.get("expenses")
        ia_rate  = sr.get("ia_rate", 0.0)
        emp_rate = sr.get("emp_rate", 0.0062)
        lines.append(f"##### {name}\n")
        if expenses is None:
            lines.append("> [확인 필요]\n")
            continue
        rate_rows = expenses.get("rate_rows", [])
        if rate_rows:
            lines += [
                f"산재보험료: {_pct(ia_rate)}, 고용보험료: {_pct(emp_rate)}\n",
                "| 항목 | 요율 | 실비(원) | 추정(원) | 합계(원) |",
                "|------|------|------|------|------|",
            ]
            for rr in rate_rows:
                lines.append(
                    f"| {rr['item']} | {_pct(rr.get('rate'))} | {_fmt(rr.get('actual'))} "
                    f"| {_fmt(rr.get('estimated'))} | {_fmt(rr.get('total'))} |"
                )
            lines.append("")
        else:
            lines.append(f"> 산재보험료: {_pct(ia_rate)}, 고용보험료: {_pct(emp_rate)} — 자료 미입력\n")

    # 5.4 일반관리비 및 이윤 (하도급사별)
    lines += [
        "### 4.4 일반관리비 및 이윤\n",
        "| 하도급사 | 소계 | 일반관리비율 | 일반관리비 | 이윤율 | 이윤 | 합계 |",
        "|------|------|------|------|------|------|------|",
    ]
    for sr in sub_results:
        name    = sr.get("name", "")
        agg     = sr.get("aggregate")
        admin   = sr.get("general_admin")
        profit  = sr.get("profit")
        ga_rate = sr.get("ga_rate", 0)
        p_rate  = sr.get("profit_rate", 0)
        if agg and admin and profit:
            sub_tot    = agg.get("subtotal", {}).get("total", 0)
            admin_tot  = admin.get("total", 0)
            profit_tot = profit.get("total", 0)
            grand_tot  = agg.get("grand_total", {}).get("total", 0)
            lines.append(
                f"| {name} | {_fmt(sub_tot)} | {_pct(ga_rate)} | {_fmt(admin_tot)} "
                f"| {_pct(p_rate)} | {_fmt(profit_tot)} | {_fmt(grand_tot)} |"
            )
        else:
            lines.append(f"| {name} | [확인 필요] | — | — | — | — | — |")
    lines.append("")

    return "\n".join(lines)


def _section6_conclusion(data: dict, calc: dict) -> str:
    c = data.get("contract", {})
    ext = data.get("extension", {})
    grand_tot = calc.get("grand_total", {}).get("total", 0)
    vat = calc.get("vat", 0)
    final = calc.get("final_rounded", 0)

    lines = [
        "---\n",
        "# 결론\n",
        f"본건 공사 「{_or(c.get('name'))}」의 공기연장 기간은 "
        f"**{_or(ext.get('start_date'))} ~ {_or(ext.get('end_date'))} ({_or(ext.get('total_days'))}일)**이며, "
        "상기 산정 기준에 따라 공기연장 간접비를 산정한 결과 아래와 같습니다.\n",
        "*(단위: 원)*\n",
        "| 구 분 | 금 액 |",
        "|------|------|",
        f"| 공사비 (부가세 전) | {_fmt(grand_tot)} |",
        f"| 부가가치세 ({'영세율(0%)' if data.get('vat_zero_rated') else '10%'}) | {_fmt(vat)} |",
        f"| **청구금액 합계** | **{_fmt(final)}** |",
        "",
        f"> **총 청구금액: 금 {_fmt(final)}원 (₩{_fmt(final)})** (부가가치세 포함, 천원 미만 절사)\n",
        "상기 금액은 관련 법령 및 계약문서에 근거한 실비 산정 결과로서, 계약상대자의 정당한 청구권에 해당합니다.\n",
    ]
    return "\n".join(lines)


def _section7_attachment(data: dict) -> str:
    unresolved   = data.get("unresolved", [])
    subcontractors = data.get("subcontractors", [])
    changes      = data.get("changes", [])
    correspondence = data.get("correspondence", [])
    expenses     = data.get("expenses_direct", [])
    labor        = [p for p in data.get("indirect_labor", []) if p.get("name")]

    attachments = []
    idx = 1

    # 항상 포함
    attachments.append((idx, "간접비 산정 보고서 (본 문서)"))
    idx += 1

    # 계약 관련
    attach_contract = "계약서 원본"
    if changes:
        attach_contract += f", 변경계약서 {len(changes)}부"
    attachments.append((idx, f"{attach_contract}, 산출내역서(공사원가계산서)"))
    idx += 1

    # 간접노무비 증빙
    if labor:
        attachments.append((idx, f"급여명세서·임금대장 (간접노무 인원 {len(labor)}명)"))
        idx += 1

    # 경비 증빙
    if expenses:
        items = ", ".join(set(e.get("item", "") for e in expenses if e.get("item")))
        attachments.append((idx, f"직접계상 경비 영수증·세금계산서 ({items})"))
        idx += 1

    # 보험료 산출 근거
    attachments.append((idx, "보험료 요율 확인서 (산재·고용보험료 요율 적용 근거)"))
    idx += 1

    # 공문 이력
    if correspondence:
        attachments.append((idx, f"수발신 공문 (공기연장 관련 공문 {len(correspondence)}건)"))
        idx += 1
    else:
        attachments.append((idx, "수발신 공문 (공기연장 관련 공문 전체)"))
        idx += 1

    # 하도급사
    if subcontractors:
        sub_names = ", ".join(s.get("name", "") for s in subcontractors[:3])
        if len(subcontractors) > 3:
            sub_names += f" 외 {len(subcontractors)-3}개사"
        attachments.append((idx, f"하도급계약서 및 하도급사 간접비 증빙자료 ({sub_names})"))
        idx += 1

    lines = [
        "---\n",
        "# 제5장 첨부자료\n",
        "| 번호 | 항목 |",
        "|------|------|",
    ]
    for no, item in attachments:
        lines.append(f"| {no} | {item} |")
    lines.append("")

    # 확인 필요 항목
    if unresolved:
        lines += [
            "---\n",
            "## ⚠️ 확인 필요 항목\n",
            "> 아래 항목은 수신자료에서 확인되지 않아 보고서에 [확인 필요]로 표시되었습니다.\n",
            "| 항목 | 확인 불가 사유 | 관련 출처 |",
            "|------|------|------|",
        ]
        for u in unresolved:
            lines.append(
                f"| {_or(u.get('item'))} | {_or(u.get('reason'))} | {_or(u.get('source', '—'))} |"
            )
        lines.append("")

    return "\n".join(lines)


# ── 메인 ──────────────────────────────────────────────────────────────────────

def generate_md(data: dict, calc: dict) -> str:
    """
    analysis_result(data) + calculate() 결과(calc) → 마크다운 문자열 반환
    """
    sections = [
        _cover(data, calc),
        _toc(data),
        _section1_overview(data),
        _section_contract_nature(data),
        _section2_reason(data),
        _section3_method(data, calc),
        _section4_result(data, calc),
        _section5_subcontractor(data, calc),
        _section6_conclusion(data, calc),
        _section7_attachment(data),
    ]
    return "\n\n".join(sections) + "\n\n끝."

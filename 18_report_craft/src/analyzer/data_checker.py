"""
analysis_result.json 데이터 충족도 검사
보고서 생성 전에 미흡한 정보를 항목별로 안내
"""

from __future__ import annotations


# ── 항목별 검사 규칙 ──────────────────────────────────────────────────────────

def _is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and (not v.strip() or v.strip() in ("확인 필요", "null", "")):
        return True
    if isinstance(v, (int, float)) and v == 0:
        return True
    return False


def check(data: dict) -> list[dict]:
    """
    데이터 충족도 검사.
    반환: [{"level": "필수"|"권장", "section": "섹션명", "field": "항목", "message": "설명", "action": "조치방법"}]
    """
    issues = []

    def warn(level, section, field, message, action):
        issues.append({"level": level, "section": section,
                        "field": field, "message": message, "action": action})

    c   = data.get("contract", {})
    ext = data.get("extension", {})
    rates = data.get("rates", {})
    labor = data.get("indirect_labor", [])
    expenses = data.get("expenses_direct", [])
    changes = data.get("changes", [])

    # ── 계약 기본 정보 ──────────────────────────────────────────────────────────
    if _is_empty(c.get("name")):
        warn("필수", "계약 정보", "공사명", "공사명이 없습니다.",
             "계약서에서 공사명을 확인하여 analysis_result.json의 contract.name에 입력하세요.")

    if _is_empty(c.get("client")):
        warn("필수", "계약 정보", "발주처", "발주처명이 없습니다.",
             "계약서 표지에서 발주처를 확인하세요.")

    if _is_empty(c.get("contractor")):
        warn("필수", "계약 정보", "계약상대자", "시공사명이 없습니다.",
             "계약서에서 계약상대자(시공사)를 확인하세요. JV인 경우 지분율도 포함해야 합니다.")

    if _is_empty(c.get("initial_amount")):
        warn("필수", "계약 정보", "최초 계약금액", "최초 계약금액이 0 또는 없습니다.",
             "계약서에서 최초 계약금액(원 단위)을 확인하세요.")

    if _is_empty(c.get("initial_date")):
        warn("권장", "계약 정보", "계약체결일", "최초 계약체결일이 없습니다.",
             "계약서에서 계약체결일(YYYY-MM-DD)을 확인하세요.")

    if _is_empty(c.get("initial_start")) or _is_empty(c.get("initial_end")):
        warn("필수", "계약 정보", "최초 공사기간", "최초 착공일 또는 준공일이 없습니다.",
             "계약서에서 최초 공사기간(착공일~준공일)을 확인하세요.")

    if _is_empty(c.get("contract_type")):
        warn("권장", "계약 정보", "계약형태", "계약형태(장기계속/총액)가 없습니다.",
             "계약서에서 계약형태를 확인하세요.")

    # ── 공기연장 정보 ───────────────────────────────────────────────────────────
    if _is_empty(ext.get("start_date")) or _is_empty(ext.get("end_date")):
        warn("필수", "공기연장", "연장 기간", "공기연장 시작일 또는 종료일(변경 준공일)이 없습니다.",
             "변경계약서에서 최초 준공일과 최종 변경 준공일을 확인하세요.")

    if _is_empty(ext.get("total_days")):
        warn("필수", "공기연장", "총 연장일수", "총 연장일수가 없습니다.",
             "변경계약서 이력에서 차수별 연장일수를 합산하여 입력하세요.")

    if not changes:
        warn("권장", "공기연장", "변경계약 이력", "변경계약 이력이 없습니다.",
             "변경계약서에서 차수별 계약일, 금액, 연장일수, 변경 사유를 입력하세요.")
    else:
        for i, ch in enumerate(changes):
            if _is_empty(ch.get("reason")):
                warn("권장", "공기연장", f"{ch.get('seq', i+1)}차 변경 사유",
                     f"{ch.get('seq', i+1)}차 변경계약의 공기지연 사유가 없습니다.",
                     "변경계약서 또는 관련 공문에서 지연 사유를 확인하세요.")

    # ── 간접노무비 ──────────────────────────────────────────────────────────────
    # _note 전용 항목(직원 아닌 메모) 제외
    persons = [p for p in labor if p.get("name")]

    if not persons:
        warn("필수", "간접노무비", "현장 인원 명단",
             "간접노무 인원 정보가 없습니다. 간접노무비를 산정할 수 없습니다.",
             "급여명세서·현장조직도에서 연장 기간 중 현장 상주 인원(소속, 이름, 직무, 월급여)을 확인하세요.")
    else:
        no_salary = [p.get("name") for p in persons if _is_empty(p.get("monthly_salary"))]
        if no_salary:
            warn("입력 대기", "간접노무비", "월 급여",
                 f"월 급여 미입력 인원 {len(no_salary)}명: {', '.join(no_salary)}",
                 "외부 확인 후 analysis_result.json의 monthly_salary에 입력하세요. "
                 "(python main.py amounts 로 입력 템플릿 출력 가능)")

        no_period = [p.get("name") for p in persons
                     if _is_empty(p.get("period_start")) or _is_empty(p.get("period_end"))]
        if no_period:
            warn("권장", "간접노무비", "개인별 산정 기간",
                 f"산정 기간이 없는 인원: {', '.join(no_period)} (전체 연장 기간 적용됨)",
                 "경력증명서·현장조직도에서 해당 인원의 현장 상주 기간을 확인하세요.")

    # ── 경비 요율 ─────────────────────────────────────────────────────────────
    # 요율은 0%가 정상값일 수 있으므로 None 여부만 확인 (0 ≠ 미입력)
    if rates.get("industrial_accident") is None:
        warn("필수", "경비", "산재보험료율",
             "산재보험료율이 없어 산재보험료를 계산할 수 없습니다.",
             "산출내역서(공사원가계산서)에서 산재보험료율을 확인하세요.")

    if rates.get("employment") is None:
        warn("필수", "경비", "고용보험료율",
             "고용보험료율이 없어 고용보험료를 계산할 수 없습니다.",
             "산출내역서(공사원가계산서)에서 고용보험료율을 확인하세요.")

    # ── 일반관리비·이윤 요율 ───────────────────────────────────────────────────
    if rates.get("general_admin") is None:
        warn("필수", "일반관리비", "일반관리비율",
             "일반관리비율이 없어 일반관리비를 계산할 수 없습니다.",
             "산출내역서(공사원가계산서) 또는 입찰 시 배포된 설계내역서에서 일반관리비율을 확인하세요.")

    if rates.get("profit") is None:
        warn("필수", "이윤", "이윤율",
             "이윤율이 없어 이윤을 계산할 수 없습니다.",
             "산출내역서(공사원가계산서) 또는 입찰 시 배포된 설계내역서에서 이윤율을 확인하세요.")

    # ── 법령 근거 ──────────────────────────────────────────────────────────────
    if not data.get("laws"):
        warn("권장", "청구 근거", "적용 법령",
             "적용 법령 정보가 없습니다.",
             "계약서에서 준거법(지방계약법/국가계약법/공사계약일반조건)을 확인하고 "
             "report_type(A/B/C)이 올바른지 검토하세요.")

    # ── 공문 이력 ──────────────────────────────────────────────────────────────
    if not data.get("correspondence"):
        warn("권장", "귀책사유", "수발신 공문 이력",
             "공기지연 관련 공문 이력이 없습니다.",
             "수발신 공문에서 공기연장 요청·승인 관련 공문 목록(일자, 발신처→수신처, 제목)을 확인하세요.")

    # ── 직접계상 경비 ──────────────────────────────────────────────────────────
    if not expenses:
        warn("입력 대기", "경비", "직접계상비목",
             "직접계상 경비 항목이 없습니다.",
             "외부 확인 후 항목명(item)과 금액(amount_actual)을 입력하세요. "
             "(python main.py amounts 로 입력 템플릿 출력 가능)")
    else:
        no_amount = [e.get("item", "?") for e in expenses
                     if _is_empty(e.get("amount_actual")) and _is_empty(e.get("amount_estimated"))]
        if no_amount:
            warn("입력 대기", "경비", "직접계상 금액",
                 f"금액 미입력 경비 항목: {', '.join(no_amount)}",
                 "외부 확인 후 amount_actual에 금액을 입력하세요. "
                 "(python main.py amounts 로 입력 템플릿 출력 가능)")

    # ── 범위 이상 검증 ─────────────────────────────────────────────────────────
    _check_ranges(data, issues, warn)

    # ── 소스 문서 교차검증 ──────────────────────────────────────────────────────
    _check_source_totals(data, issues, warn)

    return issues


def _check_source_totals(data: dict, issues: list, warn) -> None:
    """소스 문서 합계와 계산 예상값을 비교하여 추출 오류 탐지"""
    st = data.get("source_totals")
    if not st:
        return

    # 간접노무비 교차검증
    src_labor = st.get("indirect_labor_total")
    if src_labor and src_labor > 0:
        labor_persons = [p for p in data.get("indirect_labor", []) if p.get("name")]
        # 월급여 합계 × 연장일수 / 30 로 대략적 예상값 계산
        ext = data.get("extension", {})
        total_days = ext.get("total_days") or 0
        estimated_labor = sum(
            (p.get("monthly_salary") or 0) / 30 * total_days
            for p in labor_persons
        )
        if estimated_labor > 0:
            ratio = abs(estimated_labor - src_labor) / src_labor
            if ratio > 0.15:
                warn("권장", "간접노무비", "소스 문서 합계 불일치",
                     f"소스 문서 합계({src_labor:,.0f}원)와 "
                     f"입력 데이터 기반 추정값({estimated_labor:,.0f}원)이 "
                     f"{ratio:.0%} 차이납니다.",
                     f"급여 누락 인원이 있거나 급여 금액 오류일 수 있습니다. "
                     f"급여대장과 indirect_labor 항목을 재확인하세요. "
                     f"*(출처: {st.get('indirect_labor_total_source', '확인 필요')})*")


def _check_ranges(data: dict, issues: list, warn) -> None:
    """비현실적 수치 탐지 — 입력 오류 조기 발견"""

    persons   = [p for p in data.get("indirect_labor", []) if p.get("name")]
    c         = data.get("contract", {})
    ext       = data.get("extension", {})
    rates     = data.get("rates", {})
    contract_amount = c.get("initial_amount") or 0

    # 급여 null 비율 > 20% → 경고
    if persons:
        null_salary = [p for p in persons
                       if p.get("monthly_salary") is None or p.get("monthly_salary") == 0]
        null_ratio  = len(null_salary) / len(persons)
        if null_ratio > 0.2:
            names = ", ".join(p.get("name", "?") for p in null_salary[:5])
            suffix = f" 외 {len(null_salary)-5}명" if len(null_salary) > 5 else ""
            warn("필수", "간접노무비", "급여 미확인 인원 비율",
                 f"전체 {len(persons)}명 중 {len(null_salary)}명({null_ratio:.0%})의 급여가 0 또는 null입니다: "
                 f"{names}{suffix}",
                 "급여명세서(임금대장)에서 해당 인원의 실제 월 급여를 확인하세요. "
                 "급여가 없는 인원은 간접노무비가 0으로 계산됩니다.")

    # 월 급여가 비현실적으로 큰 값 탐지 (단위 오류: 원 단위여야 하는데 천원 단위로 입력)
    suspicious_salary = [
        p for p in persons
        if p.get("monthly_salary") and 0 < p["monthly_salary"] < 100_000
    ]
    if suspicious_salary:
        names = ", ".join(p.get("name", "?") for p in suspicious_salary[:3])
        warn("필수", "간접노무비", "급여 단위 오류 의심",
             f"월 급여가 100,000원 미만인 인원이 있습니다: {names}. "
             "천원(KRW 000) 단위로 입력한 것은 아닌지 확인하세요.",
             "analysis_result.json의 monthly_salary를 원(KRW) 단위 정수로 수정하세요. "
             "예: 5,200천원 → 5,200,000")

    # 연장 일수 범위 검증 (0일 또는 10년 초과는 오류 가능성)
    total_days = ext.get("total_days") or 0
    if total_days > 0 and total_days > 3650:
        warn("권장", "공기연장", "연장일수 범위 초과",
             f"총 연장일수가 {total_days}일 ({total_days//365}년 이상)입니다. 입력 오류인지 확인하세요.",
             "변경계약서 차수별 연장일수의 합산이 맞는지 재확인하세요.")

    # 요율 범위 검증
    ia = rates.get("industrial_accident")
    if ia is not None and ia > 0.20:
        warn("권장", "경비", "산재보험료율 범위 초과",
             f"산재보험료율이 {ia:.2%}로 통상 범위(0%~20%)를 초과합니다.",
             "산출내역서에서 산재보험료율을 재확인하세요. 소수 표기 여부 확인 (예: 3.7% → 0.037)")

    emp = rates.get("employment")
    if emp is not None and emp > 0.05:
        warn("권장", "경비", "고용보험료율 범위 초과",
             f"고용보험료율이 {emp:.2%}로 통상 범위(0%~5%)를 초과합니다.",
             "산출내역서에서 고용보험료율을 재확인하세요. 소수 표기 여부 확인")

    ga = rates.get("general_admin")
    if ga is not None and ga > 0.10:
        warn("권장", "일반관리비", "일반관리비율 범위 초과",
             f"일반관리비율이 {ga:.2%}로 법정 한도(6%)를 크게 초과합니다.",
             "산출내역서에서 일반관리비율을 재확인하세요. 시스템이 자동으로 한도를 적용하지만 "
             "입력값이 잘못된 경우 보고서의 '적용 요율 주석'이 부정확해질 수 있습니다.")

    pr = rates.get("profit")
    if pr is not None and pr > 0.15:
        warn("권장", "이윤", "이윤율 범위 초과",
             f"이윤율이 {pr:.2%}로 법정 한도(15%)를 초과합니다.",
             "산출내역서에서 이윤율을 재확인하세요. 시스템이 15%로 자동 절사합니다.")


def print_check_report(issues: list[dict]) -> bool:
    """
    검사 결과를 터미널에 출력.
    반환: True = 필수 항목 미흡 있음 (계속 진행 여부는 호출자가 결정)
    """
    if not issues:
        print("  ✅ 데이터 충족도 검사 통과 — 모든 필수 항목 확인됨\n")
        return False

    required  = [i for i in issues if i["level"] == "필수"]
    pending   = [i for i in issues if i["level"] == "입력 대기"]
    optional  = [i for i in issues if i["level"] == "권장"]

    if required:
        print(f"\n  ❌ [필수 항목 미흡] {len(required)}건 — 보고서에 '확인 필요'로 표시됩니다.\n")
        for iss in required:
            print(f"  ❌ [{iss['section']}] {iss['field']}")
            print(f"       문제: {iss['message']}")
            print(f"       조치: {iss['action']}\n")

    if pending:
        print(f"  📋 [금액 입력 대기] {len(pending)}건 — 외부 확인 후 입력하면 보고서에 반영됩니다.\n")
        for iss in pending:
            print(f"  📋 [{iss['section']}] {iss['field']}")
            print(f"       {iss['message']}")
            print(f"       조치: {iss['action']}\n")

    if optional:
        print(f"  ⚠️  [권장 항목 미흡] {len(optional)}건 — 없어도 보고서 생성은 가능합니다.\n")
        for iss in optional:
            print(f"  ⚠️  [{iss['section']}] {iss['field']}")
            print(f"       {iss['message']}")
            print(f"       조치: {iss['action']}\n")

    return len(required) > 0

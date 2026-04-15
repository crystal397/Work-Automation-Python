"""
간접비 계산 엔진
analysis_result.json → 항목별 금액 계산
실비(actual) / 추정(estimated) 분리 산정
"""

import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
import config


# ── 날짜 유틸 ─────────────────────────────────────────────────────────────────────

def _to_date(s: str | None) -> date | None:
    if not s or not isinstance(s, str):
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _days_between(start: str, end: str) -> int:
    s, e = _to_date(start), _to_date(end)
    if s and e:
        return max((e - s).days + 1, 0)
    return 0


def _split_actual_estimated(start: str, end: str) -> tuple[int, int]:
    """
    준공일 기준으로 실비 일수 / 추정 일수 분리.
    오늘 이전 = 실비, 오늘 이후 = 추정
    """
    today = date.today()
    s = _to_date(start)
    e = _to_date(end)
    if not s or not e:
        return 0, 0

    actual_end    = min(e, today)
    estimated_start = max(s, today + __import__('datetime').timedelta(days=1))

    actual_days    = max((actual_end - s).days + 1, 0) if actual_end >= s else 0
    estimated_days = max((e - estimated_start).days + 1, 0) if estimated_start <= e else 0
    return actual_days, estimated_days


# ── 간접노무비 ────────────────────────────────────────────────────────────────────

def calc_indirect_labor(persons: list[dict],
                        ext_start: str, ext_end: str) -> dict:
    """
    인원별 급여 × 일수 + 퇴직급여충당금 산정.
    반환: {rows, actual, estimated, total}
    """
    rows = []
    total_actual = total_estimated = 0

    ext_s = _to_date(ext_start)
    ext_e = _to_date(ext_end)

    for p in persons:
        # _note 전용 항목(직원 아닌 메모) 건너뜀
        if not p.get("name"):
            continue

        monthly = p.get("monthly_salary") or 0
        ret_rate = p.get("retirement_rate") or config.RETIREMENT_DEFAULT_RATE

        # 개인별 산정 기간을 공기연장 기간(ext_start~ext_end)으로 클리핑
        raw_start = _to_date(p.get("period_start")) or ext_s
        raw_end   = _to_date(p.get("period_end"))   or ext_e

        clipped_start = max(raw_start, ext_s) if (raw_start and ext_s) else raw_start or ext_s
        clipped_end   = min(raw_end,   ext_e) if (raw_end   and ext_e) else raw_end   or ext_e

        if not clipped_start or not clipped_end or clipped_start > clipped_end:
            rows.append({
                "name": p.get("name", ""), "org": p.get("org", ""),
                "role": p.get("role", ""), "period": "기간 없음",
                "actual_days": 0, "estimated_days": 0,
                "monthly_salary": monthly, "actual": 0, "estimated": 0, "total": 0,
                "source": p.get("source", ""),
            })
            continue

        p_start = clipped_start.strftime("%Y-%m-%d")
        p_end   = clipped_end.strftime("%Y-%m-%d")

        act_days, est_days = _split_actual_estimated(p_start, p_end)
        daily = monthly / 30

        act_salary  = round(daily * act_days)
        act_retire  = round(act_salary * ret_rate)
        act_total   = act_salary + act_retire

        est_salary  = round(daily * est_days)
        est_retire  = round(est_salary * ret_rate)
        est_total   = est_salary + est_retire

        rows.append({
            "name":           p.get("name", ""),
            "org":            p.get("org", ""),
            "role":           p.get("role", ""),
            "period":         f"{p_start} ~ {p_end}",
            "actual_days":    act_days,
            "estimated_days": est_days,
            "monthly_salary": monthly,
            "ret_rate":       ret_rate,
            "act_salary":     act_salary,
            "act_retire":     act_retire,
            "est_salary":     est_salary,
            "est_retire":     est_retire,
            "actual":         act_total,
            "estimated":      est_total,
            "total":          act_total + est_total,
            "source":         p.get("source", ""),
        })
        total_actual    += act_total
        total_estimated += est_total

    return {
        "rows":      rows,
        "actual":    total_actual,
        "estimated": total_estimated,
        "total":     total_actual + total_estimated,
    }


# ── 경비 ──────────────────────────────────────────────────────────────────────────

def calc_expenses(expenses_direct: list[dict],
                  rates: dict,
                  indirect_labor_actual: int,
                  indirect_labor_estimated: int) -> dict:
    """
    직접계상비목 합계 + 승률계상비목(산재·고용보험료) 계산
    """
    # 직접계상
    direct_actual = direct_estimated = 0
    direct_rows = []
    for e in expenses_direct:
        a = e.get("amount_actual", 0) or 0
        b = e.get("amount_estimated", 0) or 0
        direct_rows.append({
            "item":      e.get("item", ""),
            "actual":    a,
            "estimated": b,
            "total":     a + b,
            "source":    e.get("source", ""),
        })
        direct_actual    += a
        direct_estimated += b

    # 승률계상 — 산재보험료
    ia_rate = rates.get("industrial_accident", 0) or 0
    ia_act  = round(indirect_labor_actual    * ia_rate)
    ia_est  = round(indirect_labor_estimated * ia_rate)

    # 승률계상 — 고용보험료
    emp_rate = rates.get("employment", 0) or 0
    emp_act  = round(indirect_labor_actual    * emp_rate)
    emp_est  = round(indirect_labor_estimated * emp_rate)

    # 직접계상 — 국민연금 (사용자 부담 4.5%)
    np_rate = rates.get("national_pension", config.NATIONAL_PENSION_RATE)
    np_act  = round(indirect_labor_actual    * np_rate)
    np_est  = round(indirect_labor_estimated * np_rate)

    # 직접계상 — 건강보험 (사용자 부담 3.545%)
    hi_rate = rates.get("health_insurance", config.HEALTH_INSURANCE_RATE)
    hi_act  = round(indirect_labor_actual    * hi_rate)
    hi_est  = round(indirect_labor_estimated * hi_rate)

    # 직접계상 — 노인장기요양보험 (건강보험료 × 12.95%)
    ltc_rate_on_hi = rates.get("long_term_care", config.LONG_TERM_CARE_RATE)
    ltc_act = round(hi_act * ltc_rate_on_hi)
    ltc_est = round(hi_est * ltc_rate_on_hi)

    # 직접계상비목에 4대보험 중 직접계상 항목 추가
    insurance_direct_rows = [
        {
            "item":      "국민연금",
            "rate":      np_rate,
            "rate_base": "간접노무비",
            "actual":    np_act,
            "estimated": np_est,
            "total":     np_act + np_est,
            "source":    rates.get("national_pension_source", "국민연금법 제88조"),
        },
        {
            "item":      "건강보험",
            "rate":      hi_rate,
            "rate_base": "간접노무비",
            "actual":    hi_act,
            "estimated": hi_est,
            "total":     hi_act + hi_est,
            "source":    rates.get("health_insurance_source", "국민건강보험법 제69조"),
        },
        {
            "item":      "노인장기요양보험",
            "rate":      ltc_rate_on_hi,
            "rate_base": "건강보험료",
            "actual":    ltc_act,
            "estimated": ltc_est,
            "total":     ltc_act + ltc_est,
            "source":    rates.get("long_term_care_source", "노인장기요양보험법 제8조"),
        },
    ]

    # 직접계상비목 합계(직접계상 경비 + 3대보험)
    ins_direct_act = np_act + hi_act + ltc_act
    ins_direct_est = np_est + hi_est + ltc_est
    total_direct_actual    = direct_actual    + ins_direct_act
    total_direct_estimated = direct_estimated + ins_direct_est

    rate_rows = [
        {
            "item":      "산재보험료",
            "rate":      ia_rate,
            "actual":    ia_act,
            "estimated": ia_est,
            "total":     ia_act + ia_est,
            "source":    rates.get("industrial_accident_source", ""),
        },
        {
            "item":      "고용보험료",
            "rate":      emp_rate,
            "actual":    emp_act,
            "estimated": emp_est,
            "total":     emp_act + emp_est,
            "source":    rates.get("employment_source", ""),
        },
    ]

    exp_actual    = direct_actual    + ia_act + emp_act + np_act + hi_act + ltc_act
    exp_estimated = direct_estimated + ia_est + emp_est + np_est + hi_est + ltc_est

    return {
        "direct_rows":            direct_rows,
        "direct_actual":          direct_actual,
        "direct_estimated":       direct_estimated,
        "insurance_direct_rows":  insurance_direct_rows,
        "ins_direct_actual":      ins_direct_act,
        "ins_direct_estimated":   ins_direct_est,
        "total_direct_actual":    total_direct_actual,
        "total_direct_estimated": total_direct_estimated,
        "rate_rows":              rate_rows,
        "actual":                 exp_actual,
        "estimated":              exp_estimated,
        "total":                  exp_actual + exp_estimated,
    }


# ── 일반관리비 ────────────────────────────────────────────────────────────────────

def calc_general_admin(subtotal_actual: int, subtotal_estimated: int,
                       rate: float, report_type: str,
                       contract_amount: int = 0) -> dict:
    """한도 검증 후 일반관리비 계산"""
    # 한도 결정
    if report_type == "A":
        limit = config.GENERAL_ADMIN_RATE_LIMIT["A"]
    elif report_type == "B":
        limit = (config.GENERAL_ADMIN_RATE_LIMIT["B_large"]
                 if contract_amount >= config.LARGE_CONTRACT_THRESHOLD
                 else config.GENERAL_ADMIN_RATE_LIMIT["B_small"])
    else:
        limit = config.GENERAL_ADMIN_RATE_LIMIT["C"]

    applied_rate = min(rate, limit)
    if rate > limit:
        note = f"입력 요율 {rate:.2%} > 한도 {limit:.2%} → {applied_rate:.2%} 적용"
    else:
        note = ""

    act = round(subtotal_actual    * applied_rate)
    est = round(subtotal_estimated * applied_rate)

    return {
        "rate":        applied_rate,
        "rate_limit":  limit,
        "note":        note,
        "actual":      act,
        "estimated":   est,
        "total":       act + est,
    }


# ── 이윤 ──────────────────────────────────────────────────────────────────────────

def calc_profit(subtotal_actual: int, subtotal_estimated: int,
                admin_actual: int, admin_estimated: int,
                rate: float) -> dict:
    """이윤 = (소계 + 일반관리비) × 이윤율 (한도 15%)"""
    applied_rate = min(rate, config.PROFIT_RATE_LIMIT)
    note = (f"입력 요율 {rate:.2%} > 한도 15% → 15% 적용"
            if rate > config.PROFIT_RATE_LIMIT else "")

    base_act = subtotal_actual    + admin_actual
    base_est = subtotal_estimated + admin_estimated

    act = round(base_act * applied_rate)
    est = round(base_est * applied_rate)

    return {
        "rate":       applied_rate,
        "note":       note,
        "actual":     act,
        "estimated":  est,
        "total":      act + est,
    }


# ── 집계 ──────────────────────────────────────────────────────────────────────────

def _round_down_to_1000(v: int) -> int:
    """천원 단위 절사"""
    return (v // 1000) * 1000


VAT_RATE = 0.10  # 부가가치세 10%


def aggregate(labor: dict, expenses: dict,
              admin: dict, profit: dict,
              jv_share_ratio: float = 1.0,
              vat_zero_rated: bool = False) -> dict:
    """최종 집계표 생성 (부가가치세 + JV 지분율 적용)"""
    sub_act = labor["actual"]    + expenses["actual"]
    sub_est = labor["estimated"] + expenses["estimated"]

    grand_act = sub_act    + admin["actual"]    + profit["actual"]
    grand_est = sub_est    + admin["estimated"] + profit["estimated"]
    grand_tot = grand_act  + grand_est

    # JV 지분율 적용 (기본 1.0 = 적용 안 함)
    ratio        = max(0.0, min(1.0, jv_share_ratio))
    grand_jv     = round(grand_tot * ratio)
    vat          = 0 if vat_zero_rated else round(grand_jv * VAT_RATE)
    total_with_vat = grand_jv + vat

    return {
        "indirect_labor": labor,
        "expenses":        expenses,
        "subtotal": {
            "actual":    sub_act,
            "estimated": sub_est,
            "total":     sub_act + sub_est,
        },
        "general_admin": admin,
        "profit":        profit,
        "grand_total": {
            "actual":    grand_act,
            "estimated": grand_est,
            "total":     grand_tot,
        },
        "jv_share_ratio":     ratio,
        "grand_after_jv":     grand_jv,
        "vat":                vat,
        "total_with_vat":     total_with_vat,
        "final_rounded":      _round_down_to_1000(total_with_vat),
    }


# ── 메인 ──────────────────────────────────────────────────────────────────────────

def calculate(data: dict) -> dict:
    """
    analysis_result.json 딕셔너리 → 계산 결과 딕셔너리
    """
    report_type = data.get("report_type", "A")
    rates       = data.get("rates", {})
    extension   = data.get("extension", {})
    ext_start   = extension.get("start_date", "")
    ext_end     = extension.get("end_date", "")
    contract    = data.get("contract", {})
    amount      = contract.get("initial_amount", 0) or 0

    print(f"\n[계산] 유형: {report_type} | 연장 기간: {ext_start} ~ {ext_end}")

    # 1. 간접노무비
    labor = calc_indirect_labor(
        data.get("indirect_labor", []),
        ext_start, ext_end
    )
    print(f"  간접노무비: {labor['total']:,}원")

    # 2. 경비
    expenses = calc_expenses(
        data.get("expenses_direct", []),
        rates,
        labor["actual"], labor["estimated"]
    )
    print(f"  경비:       {expenses['total']:,}원")

    # 3. 일반관리비
    sub_act = labor["actual"]    + expenses["actual"]
    sub_est = labor["estimated"] + expenses["estimated"]
    admin = calc_general_admin(
        sub_act, sub_est,
        rates.get("general_admin", 0) or 0,
        report_type, amount
    )
    if admin["note"]:
        print(f"  ⚠️  일반관리비: {admin['note']}")
    print(f"  일반관리비: {admin['total']:,}원  ({admin['rate']:.2%})")

    # 4. 이윤
    profit = calc_profit(
        sub_act, sub_est,
        admin["actual"], admin["estimated"],
        rates.get("profit", 0) or 0
    )
    if profit["note"]:
        print(f"  ⚠️  이윤: {profit['note']}")
    print(f"  이윤:       {profit['total']:,}원  ({profit['rate']:.2%})")

    # 5. 집계 (JV 지분율 + 부가가치세 영세율 적용)
    jv_ratio = contract.get("jv_share_ratio", 1.0) or 1.0
    vat_zero_rated = bool(data.get("vat_zero_rated", False))
    if jv_ratio < 1.0:
        print(f"  JV 지분율: {jv_ratio:.0%} 적용")
    if vat_zero_rated:
        print("  부가가치세: 영세율(0%) 적용")
    result = aggregate(labor, expenses, admin, profit, jv_share_ratio=jv_ratio,
                       vat_zero_rated=vat_zero_rated)
    print(f"\n  ✅ 총 청구금액: {result['final_rounded']:,}원")

    # 5. 하도급사 간접비 계산
    sub_results = calc_subcontractor_indirect(data)
    result["subcontractor_results"] = sub_results
    if sub_results:
        print(f"  하도급사 간접비: {len(sub_results)}개사 산정 완료")

    return result


def calc_subcontractor_indirect(data: dict) -> list[dict]:
    """
    하도급사별 간접비 계산.
    각 subcontractor의 indirect_labor / expenses_direct / rates가 있는 경우 계산.
    없는 경우 해당 항목은 None으로 표시.
    반환: 하도급사별 결과 리스트
    """
    subcontractors = data.get("subcontractors", [])
    main_rates = data.get("rates", {})
    report_type = data.get("report_type", "A")
    contract = data.get("contract", {})
    amount = contract.get("initial_amount", 0) or 0
    main_ext = data.get("extension", {})
    main_ext_start = main_ext.get("start_date", "")
    main_ext_end   = main_ext.get("end_date", "")
    results = []

    for sub in subcontractors:
        name = sub.get("name", "")

        # claim_start/claim_end: 업무 판단으로 직접 입력한 청구 기산일 (최우선)
        # 없으면 period_start/period_end ∩ 공기연장 기간 교집합으로 자동 산정
        claim_s = _to_date(sub.get("claim_start"))
        claim_e = _to_date(sub.get("claim_end"))
        ext_s_date = _to_date(main_ext_start)
        ext_e_date = _to_date(main_ext_end)

        if claim_s or claim_e:
            ext_start = (claim_s or ext_s_date).strftime("%Y-%m-%d")
            ext_end   = (claim_e or ext_e_date).strftime("%Y-%m-%d")
        else:
            sub_own_start = _to_date(sub.get("period_start"))
            sub_own_end   = _to_date(sub.get("period_end"))
            if sub_own_start and sub_own_end and ext_s_date and ext_e_date:
                overlap_s = max(sub_own_start, ext_s_date)
                overlap_e = min(sub_own_end,   ext_e_date)
                if overlap_s <= overlap_e:
                    ext_start = overlap_s.strftime("%Y-%m-%d")
                    ext_end   = overlap_e.strftime("%Y-%m-%d")
                else:
                    ext_start = main_ext_start
                    ext_end   = main_ext_end
            else:
                ext_start = main_ext_start
                ext_end   = main_ext_end

        # 하도급사별 요율: 지정된 값 우선, 없으면 원도급 요율 사용
        ia_rate  = sub.get("industrial_accident_rate") or main_rates.get("industrial_accident", 0.0)
        emp_rate = sub.get("employment_rate", 0.0062)

        # 하도급사별 일반관리비/이윤율: 지정값 우선, 없으면 main_rates 사용
        ga_rate     = sub.get("general_admin_rate") or main_rates.get("general_admin", 0) or 0
        profit_rate = sub.get("profit_rate")        or main_rates.get("profit", 0)        or 0

        # 간접노무비 계산
        sub_persons = sub.get("indirect_labor", [])
        if sub_persons:
            sub_rates_for_calc = {
                "industrial_accident": ia_rate,
                "employment": emp_rate,
            }
            labor = calc_indirect_labor(sub_persons, ext_start, ext_end)
        else:
            labor = None  # 데이터 없음

        # 경비 계산
        sub_expenses_direct = sub.get("expenses_direct", [])
        if labor is not None:
            sub_rates_for_expense = {
                "industrial_accident": ia_rate,
                "industrial_accident_source": f"하도급사 기본값 {ia_rate:.2%}",
                "employment": emp_rate,
                "employment_source": f"하도급사 기본값 {emp_rate:.2%}",
            }
            expenses = calc_expenses(
                sub_expenses_direct,
                sub_rates_for_expense,
                labor["actual"], labor["estimated"]
            )
        elif sub_expenses_direct:
            sub_rates_for_expense = {
                "industrial_accident": ia_rate,
                "industrial_accident_source": f"하도급사 기본값 {ia_rate:.2%}",
                "employment": emp_rate,
                "employment_source": f"하도급사 기본값 {emp_rate:.2%}",
            }
            expenses = calc_expenses(sub_expenses_direct, sub_rates_for_expense, 0, 0)
        else:
            expenses = None

        # 일반관리비·이윤 계산
        if labor is not None and expenses is not None:
            sub_act = labor["actual"]    + expenses["actual"]
            sub_est = labor["estimated"] + expenses["estimated"]
            admin  = calc_general_admin(sub_act, sub_est, ga_rate, report_type, amount)
            profit = calc_profit(sub_act, sub_est, admin["actual"], admin["estimated"], profit_rate)
            agg    = aggregate(labor, expenses, admin, profit)
        else:
            sub_act = sub_est = 0
            admin  = None
            profit = None
            agg    = None

        results.append({
            "name":        name,
            "work":        sub.get("work", ""),
            "period_start": ext_start,
            "period_end":   ext_end,
            "ia_rate":     ia_rate,
            "emp_rate":    emp_rate,
            "ga_rate":     ga_rate,
            "profit_rate": profit_rate,
            "labor":       labor,
            "expenses":    expenses,
            "general_admin": admin,
            "profit":      profit,
            "aggregate":   agg,
            "has_data":    labor is not None,
        })

    return results

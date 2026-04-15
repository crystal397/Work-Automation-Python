"""
공기연장 간접비 보고서 자동 생성 — 진입점

사용법:
    python main.py extract     [입력경로] [출력경로]               # Step 1: 수신자료 추출
    python main.py review      [출력경로]                         # Step 2.5: 분석 결과 재검토
    python main.py subextract  "하도급사명" [자료경로] [출력경로]  # Step 1b: 하도급사 자료 추출
    python main.py generate    [출력경로] [--type A|B|C]           # Step 3: 보고서 생성

경로 인수는 선택사항입니다.
  지정하면 → 해당 경로 사용
  생략하면 → .env 파일의 환경변수 → 기본값(프로젝트/input, output/) 순으로 적용
"""

import sys
import io
import json
import shutil
from pathlib import Path

# ── Windows 콘솔 인코딩 강제 UTF-8 (cp949에서 이모지 출력 시 크래시 방지) ──
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).parent))

import config
from src.extractor.file_extractor import extract_folder
from src.extractor.file_classifier import classify_folder, print_classify_report, copy_includes
from src.extractor.quality_checker import check_and_report
from src.analyzer.analyzer import prepare, load
from src.calculator.calculator import calculate
from src.generator.md_generator import generate_md
from src.generator.docx_generator import generate_docx_from_template
from src.analyzer.data_checker import check, print_check_report


# ── 경로 결정 헬퍼 ─────────────────────────────────────────────────────────────

def _resolve_input(arg: str | None) -> Path:
    """입력경로: 인수 → .env/환경변수 → 기본값"""
    return Path(arg) if arg else config.INPUT_DIR

def _resolve_output(arg: str | None) -> Path:
    """출력경로: 인수 → .env/환경변수 → 기본값"""
    return Path(arg) if arg else config.OUTPUT_DIR

def _resolve_filtered_input(output_arg: str | None) -> Path:
    """분류된 입력경로: REPORT_FILTERED_DIR → output/input_filtered/ 순으로 적용"""
    if config.FILTERED_DIR:
        return config.FILTERED_DIR
    return _resolve_output(output_arg) / "input_filtered"


# ── 명령: extract ──────────────────────────────────────────────────────────────

def cmd_extract(input_path: str | None = None, output_path: str | None = None,
                use_filtered: bool = False):
    """Step 1: 수신자료 폴더 전체 추출 → output 저장

    경로 우선순위:
      1) 커맨드라인 인수
      2) 환경변수 REPORT_INPUT_DIR / REPORT_OUTPUT_DIR (.env 또는 시스템)
      3) 기본값: 프로젝트/input/,  프로젝트/output/

    --filtered 옵션:
      classify --copy 결과 폴더(input_filtered/)를 입력으로 사용합니다.
      REPORT_FILTERED_DIR(.env) → output/input_filtered/ 순으로 적용
    """
    if use_filtered:
        input_dir = _resolve_filtered_input(output_path)
    else:
        input_dir = _resolve_input(input_path)
    output_dir = _resolve_output(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_dir.exists():
        if use_filtered:
            print(f"[오류] 분류된 자료 폴더가 없습니다: {input_dir}")
            print("  먼저 classify --copy 를 실행하여 INCLUDE 파일을 복사하세요:")
            print(f"    python main.py classify --copy")
        else:
            print(f"[오류] 수신자료 폴더가 없습니다: {input_dir}")
            print("  방법 1) python main.py extract \"입력경로\" \"출력경로\"")
            print("  방법 2) .env 파일에 REPORT_INPUT_DIR=경로  설정")
            print("  방법 3) 프로젝트 폴더 안에 input/ 폴더 생성 후 자료 복사")
        sys.exit(1)

    print(f"\n[추출 시작]")
    print(f"  입력: {input_dir}")
    print(f"  출력: {output_dir}")
    results = extract_folder(str(input_dir))

    if not results:
        print("[경고] 처리할 파일이 없습니다. 입력 폴더를 확인하세요.")
        sys.exit(0)

    # 품질 검사 + 보고서 작성
    check_and_report(results, str(output_dir))

    # Claude Code 분석용 파일 저장
    analysis_file = prepare(results, str(output_dir))

    print(f"\n[완료] 추출 완료")
    print(f"  - 추출 품질 보고서: {output_dir / 'extraction_report.md'}")
    print(f"  - 분석 대상 파일:   {analysis_file}")
    print()
    print("다음 단계:")
    print("  Claude Code에서 아래 문장을 입력하세요:")
    print(f'  "{analysis_file.name} 파일을 읽고,')
    print(f"   분석지시서.md의 지침에 따라")
    print(f'   analysis_result.json 으로 저장해줘"')


# ── 명령: subextract ──────────────────────────────────────────────────────────

def cmd_subextract(sub_name: str,
                   input_path: str | None = None,
                   output_path: str | None = None):
    """Step 1b: 하도급사 자료 추출 → analysis_result.json의 해당 하도급사에 병합

    하도급사별 급여명세서·경비 자료를 별도 투입하여 분析 후
    기존 analysis_result.json의 subcontractors 항목에 추가합니다.

    사용법:
      python main.py subextract "수호토건(주)" "D:\\하도급자료\\수호토건" "C:\\...\\output(송도)"
    """
    if not sub_name:
        print("[오류] 하도급사명을 첫 번째 인수로 지정하세요.")
        print('  예: python main.py subextract "수호토건(주)" [자료경로] [출력경로]')
        sys.exit(1)

    input_dir  = _resolve_input(input_path)
    output_dir = _resolve_output(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_dir.exists():
        print(f"[오류] 하도급사 자료 폴더가 없습니다: {input_dir}")
        sys.exit(1)

    # 하도급사 전용 추출 파일 경로
    safe_name  = sub_name.replace("/", "_").replace("\\", "_").replace(" ", "_")
    sub_md     = output_dir / f"sub_{safe_name}_extracted.md"
    json_path  = output_dir / "analysis_result.json"

    print(f"\n[하도급사 자료 추출] {sub_name}")
    print(f"  입력: {input_dir}")
    print(f"  출력: {sub_md}")

    results = extract_folder(str(input_dir))
    if not results:
        print("[경고] 처리할 파일이 없습니다.")
        sys.exit(0)

    check_and_report(results, str(output_dir))
    analysis_file = prepare(results, str(output_dir))

    # 하도급사 전용 추출 파일로 복사 (extracted_for_analysis.md 덮어쓰기 방지)
    import shutil as _sh
    _sh.copy2(analysis_file, sub_md)
    # 원래 파일명 복원 (prepare가 항상 extracted_for_analysis.md로 저장하므로)
    analysis_file.rename(output_dir / "extracted_for_analysis.md")

    # 기존 JSON에 하도급사가 있는지 확인
    sub_exists = False
    if json_path.exists():
        with open(json_path, encoding="utf-8") as f:
            existing = json.load(f)
        for s in existing.get("subcontractors", []):
            if s.get("name") == sub_name:
                sub_exists = True
                break

    print(f"\n[완료] 하도급사 추출 파일: {sub_md}")
    print("\n" + "=" * 60)
    print("다음 단계 — Claude Code에 아래 문장을 입력하세요:\n")
    print(f'  "{sub_md.name} 을 읽고,')
    print(f"   analysis_result.json 의 subcontractors 중")
    print(f"   name이 '{sub_name}'인 항목에")
    print(f"   indirect_labor(간접노무비 인원·급여)와")
    print(f"   expenses_direct(직접계상 경비)를 추가해줘.")
    if not sub_exists:
        print(f"   해당 항목이 없으면 새로 추가해줘.")
    print(f'"')
    print("=" * 60 + "\n")


# ── 명령: review ──────────────────────────────────────────────────────────────

def cmd_review(output_path: str | None = None):
    """Step 2.5: analysis_result.json 재검토 프롬프트 생성

    Claude Code에 붙여넣을 수 있는 재검토 요청문을 출력합니다.
    분석지시서의 체크리스트를 기준으로 JSON을 재검토하도록 유도합니다.
    """
    output_dir = _resolve_output(output_path)
    json_path  = output_dir / "analysis_result.json"
    md_path    = output_dir / "extracted_for_analysis.md"

    if not json_path.exists():
        print(f"[오류] {json_path} 파일이 없습니다. Step 2를 먼저 완료하세요.")
        sys.exit(1)

    # 재검토 전 백업 (덮어쓰기 방지)
    backup_path = json_path.with_suffix(".json.bak")
    shutil.copy2(json_path, backup_path)
    print(f"\n[백업] {backup_path.name} 저장 완료")

    # 현재 JSON 간략 요약 출력
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    labor   = [p for p in data.get("indirect_labor", []) if p.get("name")]
    subs    = data.get("subcontractors", [])
    exps    = data.get("expenses_direct", [])
    unresolved = data.get("unresolved", [])
    null_salary = [p.get("name") for p in labor
                   if p.get("monthly_salary") is None or p.get("monthly_salary") == 0]

    print("\n[현재 analysis_result.json 요약]")
    print(f"  간접노무 인원:     {len(labor)}명  (급여 null/0: {len(null_salary)}명)")
    print(f"  하도급사:          {len(subs)}개사")
    print(f"  직접계상 경비:     {len(exps)}항목")
    print(f"  unresolved:       {len(unresolved)}건")
    if null_salary:
        print(f"  급여 없는 인원: {', '.join(null_salary[:5])}" +
              (f" 외 {len(null_salary)-5}명" if len(null_salary) > 5 else ""))

    print("\n" + "=" * 60)
    print("아래 문장을 Claude Code에 그대로 붙여넣으세요:\n")
    print(f'  "{json_path.name} 과 {md_path.name} 을 읽고,')
    print(f"   분석지시서.md의 재검토 체크리스트 6개 항목을 하나씩 확인해줘.")
    print(f"   문제가 있으면 {json_path.name} 을 직접 수정해서 저장해줘.")
    if null_salary:
        names = ", ".join(null_salary[:3])
        suffix = f" 외 {len(null_salary)-3}명" if len(null_salary) > 3 else ""
        print(f"   특히 {names}{suffix}의 급여가 null/0인데,")
        print(f"   {md_path.name} 에서 해당 인원의 급여를 찾아서 보완해줘.")
    if subs:
        print(f"   하도급사 {len(subs)}개사가 하도급계약서 기준으로 맞는지도 확인해줘.")
    print(f'"')
    print("=" * 60 + "\n")


# ── 명령: generate ─────────────────────────────────────────────────────────────

def cmd_generate(output_path: str | None = None, report_type_override: str | None = None):
    """Step 3: analysis_result.json → 보고서 생성

    경로 우선순위:
      1) 커맨드라인 인수
      2) 환경변수 REPORT_OUTPUT_DIR (.env 또는 시스템)
      3) 기본값: 프로젝트/output/

    report_type_override:
      --type A  지방계약법 기반 (지자체 발주)
      --type B  국가계약법 기반 (국가기관·공기업 발주)
      --type C  민간·사감정 기반
    """
    output_dir = _resolve_output(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    json_path = output_dir / "analysis_result.json"
    if not json_path.exists():
        print(f"[오류] 분석 결과 파일이 없습니다: {json_path}")
        print("  Step 2를 먼저 완료해주세요.")
        sys.exit(1)

    print(f"\n[보고서 생성 시작]")
    print(f"  출력 폴더: {output_dir}")

    # 분析 결과 로드
    data = load(str(output_dir))

    # --type 플래그로 유형 덮어쓰기
    if report_type_override:
        report_type_override = report_type_override.upper()
        if report_type_override not in ("A", "B", "C"):
            print(f"[오류] --type 값은 A / B / C 중 하나여야 합니다. 입력값: {report_type_override}")
            sys.exit(1)
        original_type = data.get("report_type", "?")
        data["report_type"] = report_type_override
        if original_type != report_type_override:
            print(f"  유형 덮어쓰기: {original_type} → {report_type_override} (--type 인수)")
        else:
            print(f"  유형: {report_type_override} (--type 인수, 기존과 동일)")
    else:
        print(f"  분析 결과 로드 완료: 유형 {data.get('report_type', '?')}")

    # 데이터 충족도 검사
    print("\n[데이터 충족도 검사]")
    issues = check(data)
    has_required_missing = print_check_report(issues)

    if has_required_missing:
        print("─" * 60)
        print("필수 항목이 미흡합니다. 아래 중 하나를 선택하세요:\n")
        print("  1) analysis_result.json을 직접 수정하고 다시 실행")
        print(f"     → {json_path} 열어서 해당 항목 입력")
        print()
        print("  2) Claude Code에 추가 분석 요청")
        print("     → 아래 문장을 그대로 입력하세요:")
        print()
        print(f"     {json_path.name} 을 읽고, 위에서 미흡하다고 한 항목들을")
        print(f"     {output_dir / 'extracted_for_analysis.md'} 에서 다시 찾아서 보완해줘.")
        print()
        print("  3) 미흡한 채로 보고서 초안 생성 (해당 항목은 '확인 필요'로 표시됨)")
        try:
            ans = input("  선택 (1/2/3, 기본=3): ").strip() or "3"
        except EOFError:
            ans = "3"
            print("3")
        if ans == "1":
            print("\nanalysis_result.json 수정 후 다시 실행하세요.")
            sys.exit(0)
        elif ans == "2":
            print("\nClaude Code에서 위 문장을 입력한 뒤, 다시 실행하세요.")
            sys.exit(0)
        print("\n  미흡한 항목은 보고서에 [확인 필요]로 표시하고 계속 진행합니다.\n")

    # 기존 보고서 버전 백업 (타임스탬프 포함)
    from datetime import datetime as _dt
    ts = _dt.now().strftime("%Y%m%d_%H%M%S")
    for fname in ("보고서_초안.md", "보고서_초안.docx"):
        old_file = output_dir / fname
        if old_file.exists():
            stem, suffix = fname.rsplit(".", 1)
            backup = output_dir / f"{stem}_{ts}.{suffix}"
            shutil.copy2(old_file, backup)
            print(f"  [백업] {backup.name}")

    # 계산
    calc = calculate(data)

    # 마크다운 생성
    md_text = generate_md(data, calc)
    md_path = output_dir / "보고서_초안.md"
    md_path.write_text(md_text, encoding="utf-8")
    print(f"  마크다운 저장: {md_path}")

    # Word 문서 생성 (템플릿 기반, 템플릿 없으면 마크다운 직접 변환)
    docx_path   = output_dir / "보고서_초안.docx"
    report_type = data.get("report_type", "A")
    generate_docx_from_template(data, calc, report_type, docx_path, md_text)

    print(f"\n[완료]")
    print(f"  - 마크다운: {md_path}")
    print(f"  - Word:     {docx_path}")
    print()
    total = calc.get("final_rounded", 0)
    vat   = calc.get("vat", 0)
    print(f"  총 청구금액: {total:,}원 (부가가치세 {vat:,}원 포함, 천원 미만 절사)")

    # 생성 후 최종 체크리스트 출력
    _print_final_checklist(data, calc, md_path)


# ── 명령: amounts ──────────────────────────────────────────────────────────────

def cmd_amounts(output_path: str | None = None):
    """금액 입력 템플릿 출력 — 외부 확인 후 JSON에 입력할 항목 목록

    외부 사이트에서 비목별 금액을 확인한 뒤,
    이 출력 내용을 보면서 analysis_result.json을 수정하세요.
    수정 후 python main.py generate 를 실행하면 보고서에 반영됩니다.
    """
    output_dir = _resolve_output(output_path)
    json_path  = output_dir / "analysis_result.json"

    if not json_path.exists():
        print(f"[오류] {json_path} 파일이 없습니다. Step 2를 먼저 완료하세요.")
        sys.exit(1)

    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    labor    = [p for p in data.get("indirect_labor", []) if p.get("name")]
    expenses = data.get("expenses_direct", [])
    ext      = data.get("extension", {})

    print("\n" + "=" * 70)
    print("  금액 입력 가이드 — 외부 확인 후 analysis_result.json에 입력")
    print("=" * 70)
    print(f"  파일 위치: {json_path}")
    print(f"  산정 기간: {ext.get('start_date','?')} ~ {ext.get('end_date','?')}"
          f" ({ext.get('total_days','?')}일)")
    print()

    # ── 간접노무비 ──
    print("  ┌─ 1. 간접노무비 (indirect_labor > monthly_salary)")
    print("  │   외부 사이트에서 확인한 월 급여를 아래 인원별로 입력하세요.")
    print("  │   단위: 원 (예: 5,200,000)")
    print("  │")
    print(f"  │   {'No.':<4} {'소속':<12} {'이름':<10} {'직무':<10} {'현재값':<14} {'기간'}")
    print("  │   " + "─" * 60)
    for i, p in enumerate(labor, 1):
        salary = p.get("monthly_salary")
        cur    = f"{salary:,}" if salary else "미입력"
        period = f"{p.get('period_start','?')} ~ {p.get('period_end','?')}"
        print(f"  │   {i:<4} {(p.get('org') or ''):<12} {p.get('name',''):<10} "
              f"{(p.get('role') or ''):<10} {cur:<14} {period}")
    print("  │")
    null_count = sum(1 for p in labor if not p.get("monthly_salary"))
    print(f"  └─ 합계 {len(labor)}명 중 미입력 {null_count}명\n")

    # ── 직접계상 경비 ──
    print("  ┌─ 2. 직접계상 경비 (expenses_direct > amount_actual)")
    print("  │   외부 사이트에서 확인한 비목별 금액을 입력하세요.")
    print("  │   단위: 원")
    print("  │")
    if expenses:
        print(f"  │   {'No.':<4} {'항목명':<20} {'실비(현재)':<16} {'추정(현재)'}")
        print("  │   " + "─" * 55)
        for i, e in enumerate(expenses, 1):
            act = e.get("amount_actual")
            est = e.get("amount_estimated")
            act_s = f"{act:,}" if act else "미입력"
            est_s = f"{est:,}" if est else "미입력"
            print(f"  │   {i:<4} {e.get('item','?'):<20} {act_s:<16} {est_s}")
    else:
        print("  │   경비 항목 없음 — 항목명이라도 추가 후 금액 입력 필요")
        print("  │   추가할 비목 예시: 지급임차료, 전력수도광열비, 여비교통통신비,")
        print("  │                     지급수수료, 도서인쇄비, 세금과공과, 복리후생비, 소모품비")
    null_exp = sum(1 for e in expenses
                   if not e.get("amount_actual") and not e.get("amount_estimated"))
    print("  │")
    print(f"  └─ 합계 {len(expenses)}항목 중 미입력 {null_exp}항목\n")

    print("  수정 방법:")
    print(f"    1) {json_path} 파일을 텍스트 에디터로 열기")
    print("    2) indirect_labor 각 인원의 monthly_salary 에 확인된 금액 입력")
    print("    3) expenses_direct 각 항목의 amount_actual 에 금액 입력")
    print("    4) python main.py generate 실행")
    print()
    print("  또는 Claude Code에 다음을 요청:")
    print(f'    "analysis_result.json 을 열고, 아래 금액을 입력해줘:')
    print(f"     - 홍길동 월급여: 5,200,000원")
    print(f"     - 지급임차료 실비: 12,000,000원")
    print(f'     ..."')
    print("=" * 70 + "\n")


# ── 생성 후 최종 체크리스트 ────────────────────────────────────────────────────

def _print_final_checklist(data: dict, calc: dict, md_path: "Path"):
    """보고서 생성 후 담당자가 확인해야 할 항목을 구조적으로 출력"""
    unresolved   = data.get("unresolved", [])
    subs         = data.get("subcontractors", [])
    sub_results  = calc.get("subcontractor_results", [])
    labor        = [p for p in data.get("indirect_labor", []) if p.get("name")]
    null_salary  = [p.get("name") for p in labor
                    if p.get("monthly_salary") is None or p.get("monthly_salary") == 0]
    contract     = data.get("contract", {})
    jv_ratio     = contract.get("jv_share_ratio")

    checklist = []

    # [확인 필요] 항목
    if unresolved:
        checklist.append(("필수", f"[확인 필요] 항목 {len(unresolved)}건 — 보고서 하단에 목록 있음"))
        for u in unresolved[:5]:
            checklist.append(("  →", f"{u.get('item')}: {u.get('reason', '')}"))
        if len(unresolved) > 5:
            checklist.append(("  →", f"외 {len(unresolved)-5}건 (보고서 하단 참조)"))

    # 급여 미확인 인원
    if null_salary:
        checklist.append(("필수", f"급여 미확인 인원 {len(null_salary)}명 → 간접노무비 0원으로 계산됨"))
        checklist.append(("  →", f"{', '.join(null_salary[:3])}" +
                          (f" 외 {len(null_salary)-3}명" if len(null_salary) > 3 else "")))

    # 하도급사 데이터 미입력
    subs_no_data = [sr.get("name") for sr in sub_results if not sr.get("has_data")]
    if subs_no_data:
        checklist.append(("권장", f"하도급사 간접비 미산정: {', '.join(subs_no_data[:3])}" +
                          (f" 외 {len(subs_no_data)-3}개사" if len(subs_no_data) > 3 else "")))
        checklist.append(("  →", "python main.py subextract \"하도급사명\" [자료경로]  실행 후 재생성"))

    # JV 지분율 미적용
    if jv_ratio is not None and jv_ratio < 1.0:
        checklist.append(("확인", f"JV 지분율 {jv_ratio:.0%} 적용 — 청구금액이 지분 비율로 조정됨"))
    elif "JV" in (contract.get("contractor") or "") or "공동수급" in (contract.get("contractor") or ""):
        checklist.append(("권장", "계약상대자가 JV/공동수급체로 보입니다 — jv_share_ratio 확인 필요"))
        checklist.append(("  →", "analysis_result.json의 contract.jv_share_ratio에 지분율 입력 (예: 0.6)"))

    # 법령 근거 없음
    if not data.get("laws"):
        checklist.append(("권장", "적용 법령 정보 없음 — 보고서 2장에서 수동 확인 필요"))

    # 변경계약 이력 없음
    if not data.get("changes"):
        checklist.append(("권장", "변경계약 이력 없음 — 1.2 계약현황 표가 최초 계약만 표시됨"))

    # 첨부자료
    checklist.append(("확인", "첨부자료(7장) 목록이 실제 준비된 서류와 일치하는지 확인"))

    # 출력
    print("\n" + "=" * 60)
    print("  최종 체크리스트 — 보고서 납품 전 확인사항")
    print("=" * 60)
    if not checklist:
        print("  모든 항목 이상 없음. 보고서를 검토 후 납품하세요.")
    else:
        for level, msg in checklist:
            icon = {"필수": "❌", "권장": "⚠️ ", "확인": "☑️ ", "  →": "   "}.get(level, "   ")
            print(f"  {icon} {msg}")
    print(f"\n  보고서 위치: {md_path.parent}")
    print("=" * 60 + "\n")


# ── 명령: classify ────────────────────────────────────────────────────────────

def cmd_classify(input_path: str | None = None, output_path: str | None = None,
                 auto_copy: bool = False):
    """Step 0: 수신자료 자동 분류 — INCLUDE/SKIP/UNKNOWN으로 분류

    파일명 패턴과 내용 키워드를 분석하여 각 파일을 자동 분류합니다.
    - INCLUDE: 계약서, 공문, 조직도, 산출내역서, 하도급계약서 → extract에 포함
    - SKIP   : 급여대장, 경비 영수증 → Step 2.7에서 수동 입력
    - UNKNOWN: 판단 불가 → 수동 확인 필요

    --copy 옵션 사용 시 INCLUDE 파일만 자동으로 출력폴더/input_filtered/ 에 복사합니다.
    """
    input_dir  = _resolve_input(input_path)
    output_dir = _resolve_output(output_path)

    if not input_dir.exists():
        print(f"[오류] 수신자료 폴더가 없습니다: {input_dir}")
        sys.exit(1)

    print(f"\n[파일 분류 시작]")
    print(f"  대상 폴더: {input_dir}")

    results = classify_folder(input_dir)
    if not results:
        print("[경고] 처리할 파일이 없습니다.")
        sys.exit(0)

    print_classify_report(results, str(input_dir))

    if auto_copy:
        filtered_dir = config.FILTERED_DIR or (output_dir / "input_filtered")
        copied = copy_includes(results, input_dir, filtered_dir)
        print(f"\n[자동 복사 완료] {len(copied)}개 파일 → {filtered_dir}")
        print("  다음 단계 — 분류된 파일만 추출:")
        print(f"    python main.py extract --filtered")
    else:
        print("\n  INCLUDE 파일만 자동 복사하려면 --copy 옵션을 추가하세요:")
        print(f"    python main.py classify \"{input_dir}\" --copy")


# ── 진입점 ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    cmd = sys.argv[1].lower()
    if cmd == "classify":
        args       = [a for a in sys.argv[2:] if not a.startswith("--")]
        auto_copy  = "--copy" in sys.argv
        input_path  = args[0] if len(args) > 0 else None
        output_path = args[1] if len(args) > 1 else None
        cmd_classify(input_path, output_path, auto_copy)
    elif cmd == "extract":
        use_filtered = "--filtered" in sys.argv
        args        = [a for a in sys.argv[2:] if not a.startswith("--")]
        input_path  = args[0] if len(args) > 0 and not use_filtered else None
        output_path = args[1] if len(args) > 1 else (args[0] if len(args) > 0 and use_filtered else None)
        cmd_extract(input_path, output_path, use_filtered)
    elif cmd == "amounts":
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        cmd_amounts(output_path)
    elif cmd == "subextract":
        sub_name    = sys.argv[2] if len(sys.argv) > 2 else ""
        input_path  = sys.argv[3] if len(sys.argv) > 3 else None
        output_path = sys.argv[4] if len(sys.argv) > 4 else None
        cmd_subextract(sub_name, input_path, output_path)
    elif cmd == "review":
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        cmd_review(output_path)
    elif cmd == "generate":
        # --type A|B|C 파싱
        type_override = None
        for i, arg in enumerate(sys.argv[2:], 2):
            if arg == "--type" and i + 1 < len(sys.argv):
                type_override = sys.argv[i + 1]
            elif arg.startswith("--type="):
                type_override = arg.split("=", 1)[1]
        # 경로 인수: --type 과 그 값을 제외한 첫 번째 비-플래그 인수
        skip_next = False
        path_args = []
        for arg in sys.argv[2:]:
            if skip_next:
                skip_next = False
                continue
            if arg == "--type":
                skip_next = True
                continue
            if arg.startswith("--type="):
                continue
            path_args.append(arg)
        output_path = path_args[0] if path_args else None
        cmd_generate(output_path, type_override)
    else:
        print(f"[오류] 알 수 없는 명령: {cmd}")
        print("  사용법: python main.py classify   [입력경로] [--copy]")
        print("          python main.py extract    [입력경로] [출력경로]")
        print("          python main.py extract    --filtered [출력경로]")
        print('          python main.py subextract "하도급사명" [자료경로] [출력경로]')
        print("          python main.py review     [출력경로]")
        print("          python main.py generate   [출력경로] [--type A|B|C]")
        sys.exit(1)


if __name__ == "__main__":
    main()

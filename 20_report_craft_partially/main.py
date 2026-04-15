"""
귀책분석 자동화 시스템 — 진입점

사용법:
    python main.py learn                                    # Step 1: reference 보고서 학습
    python main.py scan   <공문폴더경로> [--project <이름>]  # Step 2: 업체 공문 스캔
    python main.py prepare [프로젝트명]                      # Step 3: claude.ai용 프롬프트 조립
    python main.py validate [프로젝트명]                     # Step 3.5: JSON 사전 검증 (generate 전 오류 확인)
    python main.py generate [프로젝트명]                     # Step 4: JSON → docx 생성

콤보 명령 (단계 합산):
    python main.py scanprepare <공문폴더경로> [--project <이름>]  # scan + prepare 한 번에
    python main.py finish [프로젝트명]                            # validate 후 오류 없으면 generate 자동 실행

워크플로우:
  1. python main.py learn
     → output/reference_patterns.md 생성 (reference 보고서 귀책분석 패턴)

  2. python main.py scan "C:\\Users\\...\\수신자료"
     → output/<프로젝트명>/scan_summary.md
     → output/<프로젝트명>/scan_result.json
     → output/<프로젝트명>/correspondence_texts.md
     ※ 프로젝트명은 스캔 경로의 폴더명에서 자동 생성.
        직접 지정하려면 --project <이름> 옵션 사용.
     ※ scan_result.json 을 열어 관련 없는 항목 삭제/추가

  3. python main.py prepare [프로젝트명]
     → output/<프로젝트명>/prompt_for_claude.md
     → output/<프로젝트명>/귀책분석_schema.json
     ※ 프로젝트명 생략 시 마지막 scan 의 프로젝트 자동 사용

  3.5 (Claude가 JSON 저장 후) python main.py validate [프로젝트명]
     → 귀책분석_data.json 의 스키마 오류 + 소결 누락 + delay_days 누락 등을 사전 점검
     → 오류 발견 시 "Claude에게 전달할 수정 요청" 블록을 자동 출력
     ※ generate 실행 전에 validate 로 오류를 확인하면 재작업 횟수를 줄일 수 있습니다.

  4. python main.py generate [프로젝트명]
     → output/<프로젝트명>/02_귀책분석_[프로젝트명]_[날짜].docx
     ※ 프로젝트명 생략 시 마지막 scan 의 프로젝트 자동 사용

--project 옵션:
    scan 단계에서만 사용하며, 자동 생성된 폴더명 대신 원하는 이름을 지정한다.
    예) python main.py scan "C:\\...\\수신자료" --project 인덕원5공구
"""

import io
import sys
from pathlib import Path

# Windows 콘솔 UTF-8
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent))

import json

import config
from src.reference_learner import learn_all
from src.correspondence_scanner import scan
from src.prompt_builder import build as build_prompt
from src.report_generator import generate, _validate_data


def _print_header():
    print("=" * 60)
    print("  귀책분석 자동화 시스템  v1.0")
    print("=" * 60)


def _resolve_project_dir(project_name: str | None) -> tuple[str, Path]:
    """
    프로젝트명 → (project_name, project_dir) 반환.
    project_name 이 None 이면 .current_project 에서 읽는다.
    둘 다 없으면 오류 출력 후 종료.
    """
    if project_name:
        pdir = config.get_project_dir(project_name)
    else:
        project_name = config.load_current_project()
        if not project_name:
            print(
                "[오류] 프로젝트명을 알 수 없습니다.\n"
                "  먼저 scan 을 실행하거나 명령에 프로젝트명을 직접 지정하세요.\n"
                "  예) python main.py prepare 인덕원5공구"
            )
            sys.exit(1)
        pdir = config.get_project_dir(project_name)

    if not pdir.exists():
        print(
            f"[오류] 프로젝트 폴더가 없습니다: {pdir}\n"
            f"  먼저 python main.py scan 을 실행하세요."
        )
        sys.exit(1)

    return project_name, pdir


# ── Step 1: learn ────────────────────────────────────────────────────────────

def _token_overlap_score(a: str, b: str) -> float:
    """
    두 문자열의 단어 겹침 비율 (0.0 ~ 1.0).
    공백·기호 분리 + 한국어 숫자 경계 분리까지 적용하여 공통 토큰 수 / min(len_a, len_b)
    """
    import re as _re
    def _tokens(s: str) -> set[str]:
        # 1) 특수문자를 공백으로 치환
        norm = _re.sub(r"[\s\-~_\[\]()\.·×,·!?·]", " ", s)
        # 2) 공백 분리 + 2자 이상 토큰
        base = {p for p in norm.split() if len(p) >= 2}
        # 3) 한국어+숫자 경계에서 추가 분리 (예: "부산도시철도양산선1공구현장" → "양산선1", "1공구")
        extra: set[str] = set()
        for seg in _re.findall(r"[가-힣]+[0-9]+[가-힣]*|[0-9]+[가-힣]+", norm.replace(" ", "")):
            if len(seg) >= 2:
                extra.add(seg)
        # 4) 숫자 포함 한국어 키워드 단독 (예: "11공구", "5공구", "1공구")
        for seg in _re.findall(r"[0-9]+공구", norm):
            extra.add(seg)
        return base | extra

    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    common = ta & tb
    return len(common) / min(len(ta), len(tb))


def _print_reference_output_mapping():
    """
    reference 파일명 ↔ output 폴더명 퍼지 매칭 결과를 출력.
    학습 완료 후 진단용으로 호출한다.
    """
    ref_dir = config.REFERENCE_DIR
    out_dir = config.OUTPUT_DIR

    ref_files = (
        sorted(ref_dir.glob("*.docx")) + sorted(ref_dir.glob("*.DOCX"))
        + sorted(ref_dir.glob("*.pdf")) + sorted(ref_dir.glob("*.PDF"))
    )
    # 템플릿 파일 제외 (보고서_템플릿* / 비용산정기준*)
    ref_files = [
        f for f in ref_files
        if not any(kw in f.name for kw in ["템플릿", "비용산정기준", "template", "Template"])
    ]

    if not out_dir.exists():
        return

    out_folders = [d for d in sorted(out_dir.iterdir()) if d.is_dir()]
    if not out_folders:
        return

    print("\n[참고] reference ↔ output 폴더 매칭 진단")
    print("─" * 60)

    THRESHOLD = 0.35  # 이 점수 이상이면 매칭으로 간주

    matched_refs: set[str] = set()
    for folder in out_folders:
        best_score = 0.0
        best_ref: Path | None = None
        for rf in ref_files:
            score = _token_overlap_score(folder.name, rf.stem)
            if score > best_score:
                best_score = score
                best_ref = rf
        if best_ref and best_score >= THRESHOLD:
            tag = "✅"
            matched_refs.add(best_ref.name)
        else:
            tag = "❓"
            best_ref = None

        ref_label = best_ref.name if best_ref else "(매칭 없음)"
        print(f"  {tag} output/{folder.name}/")
        print(f"       ← reference/{ref_label}")

    # 매칭 안 된 reference 파일
    unmatched = [rf for rf in ref_files if rf.name not in matched_refs]
    if unmatched:
        print()
        print("  ⚠️  output 폴더 없이 reference만 존재하는 파일:")
        for rf in unmatched:
            print(f"     · {rf.name}")

    print("─" * 60)


def cmd_learn():
    _print_header()
    print("\n[1단계] REFERENCE 보고서 학습")
    print("─" * 60)

    if not config.REFERENCE_DIR.exists():
        print(f"[오류] reference 폴더가 없습니다: {config.REFERENCE_DIR}")
        sys.exit(1)

    # learn 결과는 프로젝트 무관 공통 → OUTPUT_DIR 루트에 저장
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = learn_all(config.REFERENCE_DIR, config.OUTPUT_DIR)

    print("\n" + "─" * 60)
    print("REFERENCE 학습 완료.")
    print(f"\n→ {out_path}")

    _print_reference_output_mapping()

    print("\n다음 단계: python main.py scan <공문폴더경로>")


# ── Step 2: scan ─────────────────────────────────────────────────────────────

def _parse_scan_args(args: list[str]) -> tuple[list[str], str | None]:
    """
    scan 인수를 파싱한다.
    --project <이름> 옵션을 분리하고 나머지는 경로 목록으로 반환.
    반환: (경로 목록, 프로젝트명 또는 None)
    """
    paths: list[str] = []
    project: str | None = None
    i = 0
    while i < len(args):
        if args[i] == "--project" and i + 1 < len(args):
            project = args[i + 1]
            i += 2
        else:
            paths.append(args[i])
            i += 1
    return paths, project


def _extract_project_name(path: Path) -> str:
    """
    경로의 상위 폴더를 탐색하여 프로젝트명을 추출한다.

    지원 패턴 (순서대로 시도):
      1. YYMMDD_업체명_프로젝트명 공기연장 간접비  (언더스코어 두 개)
         예) 251120_경남기업_평택기지-오산 천연가스 공급설비 건설공사 2공구 공기연장 간접비
             → 평택기지-오산 천연가스 공급설비 건설공사 2공구
      2. YYMMDD 프로젝트명 공기연장 간접비  (공백 구분, 업체명 없음)
         예) 240226 울산-포항 전철2공구 공기연장 간접비
             → 울산-포항 전철2공구
      3. YYMMDD_업체명 프로젝트명 공기연장 간접비  (언더스코어 한 개, 업체+프로젝트 공백 구분)
         예) 230920_디엘이앤씨 울릉공항건설공사 공기연장 간접비
             → 디엘이앤씨 울릉공항건설공사

    패턴 모두 실패하면 상위 폴더 중 '간접비'/'공기연장'이 포함된 가장 가까운 폴더명을 반환.
    그것도 없으면 입력 경로의 상위 폴더명을 반환한다.
    """
    import re
    patterns = [
        re.compile(r"^\d{6}_[^_]+_(.+?)\s*(?:공기연장\s*)?간접비.*$"),   # 언더스코어 두 개
        re.compile(r"^\d{6}\s+(.+?)\s*(?:공기연장\s*)?간접비.*$"),        # 공백 구분
        re.compile(r"^\d{6}_(.+?)\s*(?:공기연장\s*)?간접비.*$"),          # 언더스코어 한 개
    ]
    candidates = [path, *path.parents]
    for parent in candidates:
        for pat in patterns:
            m = pat.match(parent.name)
            if m:
                return m.group(1).strip()
    # 패턴 미매칭 — '간접비' 또는 '공기연장'이 포함된 가장 가까운 상위 폴더명 사용
    for parent in candidates:
        if any(kw in parent.name for kw in ("간접비", "공기연장", "클레임")):
            return parent.name
    # 최후 폴백: 입력 경로 자체의 폴더명
    return path.name


def cmd_scan(raw_args: list[str]):
    _print_header()
    print("\n[2단계] 업체 공문 스캔 및 필터링")
    print("─" * 60)

    import os

    vendor_path_strs, project_name_arg = _parse_scan_args(raw_args)

    # 경로 결정: 인수 → 환경변수 → 오류
    if vendor_path_strs:
        vendor_dirs = [Path(p) for p in vendor_path_strs]
    else:
        env_path = os.environ.get("CORRESPONDENCE_DIR", "")
        if env_path:
            vendor_dirs = [Path(p.strip()) for p in env_path.split(";") if p.strip()]
        else:
            print("[오류] 공문 폴더 경로를 지정하세요.")
            print("  python main.py scan <경로1> [경로2] ... [--project <이름>]")
            print("  또는 .env 파일에 CORRESPONDENCE_DIR=<경로> 설정")
            sys.exit(1)

    for d in vendor_dirs:
        if not d.exists():
            print(f"[오류] 폴더가 존재하지 않습니다: {d}")
            sys.exit(1)

    # 프로젝트명 결정: --project > 상위 폴더 패턴 추출 > 스캔 폴더명
    if project_name_arg:
        project_name = project_name_arg
    else:
        project_name = _extract_project_name(vendor_dirs[0])

    project_dir = config.get_project_dir(project_name)
    project_dir.mkdir(parents=True, exist_ok=True)

    # 현재 프로젝트 저장
    config.save_current_project(project_name)

    print(f"\n프로젝트 폴더: output/{project_dir.name}/")
    result_path = scan(vendor_dirs, project_dir)

    print("\n" + "─" * 60)
    print("스캔 완료.")
    print(f"\n→ output/{project_dir.name}/scan_summary.md  를 열어 공문 목록을 확인하세요.")
    print(f"→ output/{project_dir.name}/scan_result.json 을 편집하여 항목을 추가/제외한 후")
    print(f"  python main.py prepare 를 실행하세요.")


# ── Step 3: prepare ──────────────────────────────────────────────────────────

def cmd_prepare(project_name: str | None):
    _print_header()
    print("\n[3단계] claude.ai 프롬프트 조립")
    print("─" * 60)

    project_name, project_dir = _resolve_project_dir(project_name)
    print(f"\n프로젝트 폴더: output/{project_dir.name}/")

    try:
        prompt_path = build_prompt(project_dir, project_name)
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"[오류] {e}")
        sys.exit(1)

    print("\n" + "─" * 60)
    print("프롬프트 생성 완료.")
    print(f"\n→ {prompt_path}")
    print("\n다음 단계 (둘 중 하나 선택):")
    print()
    print("  ★ Claude Code 사용 (터미널에서 바로):")
    print(f'    "output/{project_dir.name}/prompt_for_claude.md 읽고')
    print(f'     귀책분석_data.json 생성해줘"')
    print(f"    → Claude Code가 파일을 읽고 분석하여 JSON을 직접 저장")
    print()
    print("  ★ claude.ai 사용 (웹 브라우저):")
    print("    1. prompt_for_claude.md 전체 내용을 복사")
    print("    2. claude.ai 새 대화창에 붙여넣기")
    print("    3. claude.ai가 작성한 JSON 전체를 복사")
    print(f"    4. output/{project_dir.name}/귀책분석_data.json 파일로 저장")
    print()
    print("  → 완료 후: python main.py generate")


# ── Step 4: generate ─────────────────────────────────────────────────────────

def cmd_generate(project_name: str | None):
    _print_header()
    print("\n[4단계] docx 생성")
    print("─" * 60)

    project_name, project_dir = _resolve_project_dir(project_name)
    # 폴더명에서 날짜 접두사·불필요 접미사를 제거한 순수 프로젝트명 추출
    extracted = _extract_project_name(project_dir)
    display_name = extracted if extracted != project_dir.name else project_name
    print(f"\n프로젝트 폴더: output/{project_dir.name}/")

    try:
        out_path = generate(project_dir, display_name)
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"[오류] {e}")
        sys.exit(1)

    print("\n" + "─" * 60)
    print("docx 생성 완료.")
    print(f"\n→ {out_path}")


# ── Step 3.5: validate ───────────────────────────────────────────────────────

def cmd_validate(project_name: str | None, _stop_on_error: bool = True) -> bool:
    """검증 실행. 반환: True=hard_error 없음, False=hard_error 있음."""
    _print_header()
    print("\n[3.5단계] 귀책분석_data.json 사전 검증")
    print("─" * 60)

    project_name, project_dir = _resolve_project_dir(project_name)
    print(f"\n프로젝트 폴더: output/{project_dir.name}/")

    # ── JSON 파일 탐색 ────────────────────────────────────────────────
    import glob as _glob
    import io as _io

    candidates = _glob.glob(str(project_dir / "*data.json"))
    if not candidates:
        print(f"\n[오류] 귀책분석_data.json 파일이 없습니다: {project_dir}")
        print("  먼저 Claude에게 prompt_for_claude.md를 전달하여 JSON을 생성하세요.")
        sys.exit(1)
    data_path = Path(candidates[0])

    with open(data_path, encoding="utf-8") as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError as e:
            print(f"\n[오류] JSON 파싱 실패: {e}")
            print("─" * 60)
            print("  귀책분석_data.json 파일이 JSON 형식 오류입니다.")
            print(f"  오류 위치: {e}")
            print("  JSON 전체를 다시 검토하고 올바른 형식으로 재출력해 주세요.")
            print("─" * 60)
            sys.exit(1)

    # ── 분류: hard_errors(generate 중단) / soft_warnings(경고만) ──────
    hard_errors:  list[str] = []   # Claude 수정 필요 → generate 불가
    soft_warnings: list[str] = []  # 확인 권고 → generate는 가능

    # ── scan_result.json 로드 (교차검증용) ───────────────────────────
    scan_items_by_no: dict[int, dict] = {}
    scan_doc_numbers: set[str] = set()
    scan_result_path = project_dir / "scan_result.json"
    if scan_result_path.exists():
        try:
            with open(scan_result_path, encoding="utf-8") as _sf:
                _sdata = json.load(_sf)
            for idx_s, _si in enumerate(_sdata.get("items", []), 1):
                _no = _si.get("no", idx_s)
                scan_items_by_no[int(_no)] = _si
                _dn = (_si.get("doc_number") or "").strip()
                if _dn:
                    scan_doc_numbers.add(_dn)
        except Exception:
            pass

    # ── [경고] items scan_no 교차검증 ────────────────────────────────
    data_items = data.get("items", [])
    if isinstance(data_items, list) and scan_items_by_no:
        for i, di in enumerate(data_items):
            sno = di.get("scan_no") or di.get("no")
            if sno is not None:
                try:
                    sno = int(sno)
                except (ValueError, TypeError):
                    continue
                if sno not in scan_items_by_no:
                    soft_warnings.append(
                        f"  [경고] items[{i}] scan_no={sno} 가 scan_result.json에 없습니다.\n"
                        f"     수동 추가 항목이면 무시하세요. 아니라면 번호를 확인하세요."
                    )

    # ── [오류] accountability_diagram 합계 행 삽입 감지 ─────────────
    diagram = data.get("accountability_diagram", [])
    _SUM_ROW_KW = {"합계", "소계", "계", "total", "sum"}
    if isinstance(diagram, list):
        for i, row in enumerate(diagram):
            if not isinstance(row, dict):
                continue
            cause_raw = (row.get("cause") or row.get("delay_cause") or "").strip()
            if cause_raw.lower() in _SUM_ROW_KW or cause_raw in _SUM_ROW_KW:
                total = data.get("total_delay_days", "?")
                hard_errors.append(
                    f"  [오류] accountability_diagram[{i}] cause='{cause_raw}' — 합계 행이 삽입되어 있습니다.\n"
                    f"     이 행을 삭제하세요. report_generator가 delay_days를 자동 합산하므로\n"
                    f"     합계 행이 있으면 total_delay_days({total})의 2배로 오류 발생합니다.\n"
                    f"     accountability_diagram에는 실제 지연 사유 항목만 포함하세요."
                )

    # ── [오류] accountability_diagram delay_days 누락 ────────────────
    if isinstance(diagram, list):
        for i, row in enumerate(diagram):
            if not isinstance(row, dict):
                continue
            dd = row.get("delay_days")
            if dd is None:
                cause = (row.get("cause") or row.get("delay_cause") or f"항목 {i+1}")[:30]
                hard_errors.append(
                    f"  [오류] accountability_diagram[{i}] ('{cause}')에\n"
                    f"     delay_days 필드가 없습니다. 일수 미확정 시 0으로 기재하세요.\n"
                    f"     ※ 모든 항목 delay_days 합계 = total_delay_days({data.get('total_delay_days','?')})\n"
                    f"     ※ '합계' 행 별도 추가 금지 (이중 계산됨)"
                )

    # ── [경고] accountability_diagram source_docs 누락 ───────────────
    if isinstance(diagram, list):
        for i, row in enumerate(diagram):
            if not isinstance(row, dict):
                continue
            cause_label = (row.get("cause") or row.get("delay_cause") or f"항목 {i+1}")[:30]
            sd = row.get("source_docs")
            if sd is None or (isinstance(sd, list) and len(sd) == 0):
                soft_warnings.append(
                    f"  [경고] accountability_diagram[{i}] ('{cause_label}')\n"
                    f"     source_docs 필드가 없거나 비어 있습니다.\n"
                    f"     근거 공문번호·변경계약 차수를 리스트로 추가하세요.\n"
                    f"     예: \"source_docs\": [\"제22-0123호\", \"3차 변경계약\"]"
                )
            elif isinstance(sd, list) and scan_doc_numbers:
                # source_docs 내 번호가 scan에 있는지 교차검증
                for ref in sd:
                    ref_str = str(ref).strip()
                    if ref_str and ref_str not in scan_doc_numbers:
                        # 부분 매칭 허용 (공문번호가 앞뒤 텍스트와 붙어있을 수 있음)
                        matched = any(ref_str in dn or dn in ref_str for dn in scan_doc_numbers)
                        if not matched:
                            soft_warnings.append(
                                f"  [경고] accountability_diagram[{i}] source_docs의\n"
                                f"     '{ref_str}' 가 scan_result.json 공문번호 목록에 없습니다.\n"
                                f"     공문번호 표기가 다를 수 있으니 직접 확인하세요."
                            )

    # ── [오류/경고] detail_narratives 소결 누락 ─────────────────────
    # 결론 문구는 계약 유형 무관하게 동일 → hard_error
    # 조항 인용은 민간계약 등 비표준 조항 가능 → soft_warning
    _SOGYEOL_CONCLUSION = "계약상대자의 책임 없는 사유"
    _SOGYEOL_CLAUSE_KW  = [
        "제8절", "제22조", "제25조", "제26조", "제27조", "제74조",
        "일반조건", "계약조건", "해당 조항", "계약 조항",
    ]
    narratives = data.get("detail_narratives", [])
    if isinstance(narratives, list):
        for i, block in enumerate(narratives):
            if not isinstance(block, dict):
                continue
            paras = block.get("paragraphs", [])
            if not isinstance(paras, list) or len(paras) == 0:
                continue
            block_title = block.get("title") or block.get("label", f"[{i}]")
            paras_text = " ".join(str(p) for p in paras)
            if _SOGYEOL_CONCLUSION not in paras_text:
                hard_errors.append(
                    f"  [오류] detail_narratives '{block_title}' 블록에 소결 결론이 없습니다.\n"
                    f"     마지막 단락에 \"계약상대자의 책임 없는 사유에 해당합니다\"를 포함시키세요."
                )
            elif not any(kw in paras_text for kw in _SOGYEOL_CLAUSE_KW):
                soft_warnings.append(
                    f"  [경고] detail_narratives '{block_title}' 블록에 근거 조항 인용이 없습니다.\n"
                    f"     소결에 공사계약일반조건 제XX조 등 근거 조항을 인용하세요."
                )

    # ── [경고] responsible_party 유효값 확인 ─────────────────────────
    _VALID_PARTIES = {"발주처", "시공사", "감리", "불가항력", "제3자"}
    if isinstance(diagram, list):
        for i, row in enumerate(diagram):
            if not isinstance(row, dict):
                continue
            rp = (row.get("responsible_party") or "").strip()
            if rp and rp not in _VALID_PARTIES:
                cause_label = (row.get("cause") or row.get("delay_cause") or f"항목 {i+1}")[:30]
                soft_warnings.append(
                    f"  [경고] accountability_diagram[{i}] ('{cause_label}')\n"
                    f"     responsible_party='{rp}' 가 표준값이 아닙니다.\n"
                    f"     표준값: {' / '.join(sorted(_VALID_PARTIES))}"
                )

    # ── [경고] items 날짜 미확정 / doc_number OCR 아티팩트 / scan_no null ─
    import re as _re
    _DATE_UNCERTAIN = ("xx", "?", "미확정", "불명", "확인필요")
    if isinstance(data_items, list):
        for i, di in enumerate(data_items):
            if not isinstance(di, dict):
                continue
            item_label = f"items[{i}](no={di.get('no','?')})"

            # 날짜 미확정
            dt = (di.get("date") or "").strip()
            if any(tok in dt for tok in _DATE_UNCERTAIN):
                soft_warnings.append(
                    f"  [경고] {item_label} date='{dt}' — 미확정 날짜입니다.\n"
                    f"     원문 공문에서 실제 발신일을 확인 후 수정하세요."
                )

            # doc_number OCR 아티팩트 의심
            # 한글이 포함된 번호에 소문자 영문 2자 이상이 연속 등장하면 OCR 오인식 가능성
            dn = (di.get("doc_number") or "").strip()
            if (dn
                    and _re.search(r"[가-힣]", dn)
                    and _re.search(r"[a-z]{2,}", dn)):
                soft_warnings.append(
                    f"  [경고] {item_label} doc_number='{dn}' — OCR 오인식 의심.\n"
                    f"     한글 번호에 소문자 영문이 섞여 있습니다 (예: -ys, -bts).\n"
                    f"     원문 공문에서 실제 문서번호를 확인 후 수정하세요."
                )

            # scan_no null → 수동 추가 항목
            sno = di.get("scan_no")
            if sno is None:
                soft_warnings.append(
                    f"  [경고] {item_label} scan_no=null — 스캔 미수록 수동 추가 항목.\n"
                    f"     원본 파일이 scan_result.json에 없으므로 내용을 직접 검토하세요."
                )

    # ── 스키마 검증 (_validate_data) — 출력 캡처 ─────────────────────
    captured = _io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = captured
    schema_error = None
    try:
        _validate_data(data, data_path)
    except ValueError as e:
        schema_error = str(e)
    finally:
        sys.stdout = old_stdout
    schema_output = captured.getvalue()

    # ── 결과 출력 ─────────────────────────────────────────────────────
    has_schema_error = schema_error is not None

    if schema_output.strip():
        print(schema_output, end="")

    # soft_warnings 출력 (generate는 가능)
    if soft_warnings:
        print()
        print("─" * 60)
        print("  ⚠️  경고 (generate는 가능하지만 확인 권고)")
        print("─" * 60)
        for w in soft_warnings:
            print(w)

    # 오류 없으면 통과
    if not hard_errors and not has_schema_error:
        print("\n" + "─" * 60)
        if soft_warnings:
            print("⚠️  경고 있음 — 위 항목 확인 후 generate 를 실행하세요.")
        else:
            print("✅ 검증 통과 — 오류 없음.")
        print("─" * 60)
        return True

    # hard_errors 있으면 Claude 수정 요청 출력 후 중단
    print()
    print("=" * 60)
    print("  ▼ Claude에게 전달할 수정 요청 (아래를 복사하여 붙여넣기)")
    print("=" * 60)
    print()
    print("귀책분석_data.json에 다음 오류가 있습니다. 수정 후 전체 JSON을 다시 출력해 주세요.\n")
    for err in hard_errors:
        print(err)
    if has_schema_error:
        print(f"\n[스키마 오류] {schema_error}")
    print()
    print("=" * 60)
    if _stop_on_error:
        sys.exit(1)
    return False


# ── 콤보 명령 ────────────────────────────────────────────────────────────────

def cmd_scan_prepare(raw_args: list[str]):
    """scan + prepare 한 번에 실행."""
    cmd_scan(raw_args)
    # scan이 current_project를 저장하므로 project_name=None으로 자동 참조
    print()
    cmd_prepare(None)


def cmd_finish(project_name: str | None):
    """validate 후 오류 없으면 generate 자동 실행."""
    ok = cmd_validate(project_name, _stop_on_error=False)
    if ok:
        print()
        cmd_generate(project_name)


# ── CLI 파싱 ─────────────────────────────────────────────────────────────────

def main():
    args = sys.argv[1:]

    if not args:
        print(__doc__)
        sys.exit(0)

    cmd = args[0].lower()

    if cmd == "learn":
        cmd_learn()

    elif cmd == "scan":
        cmd_scan(args[1:])

    elif cmd == "prepare":
        project_name = args[1] if len(args) > 1 else None
        cmd_prepare(project_name)

    elif cmd == "validate":
        project_name = args[1] if len(args) > 1 else None
        cmd_validate(project_name)

    elif cmd == "generate":
        project_name = args[1] if len(args) > 1 else None
        cmd_generate(project_name)

    elif cmd in ("scanprepare", "scan-prepare", "sp"):
        cmd_scan_prepare(args[1:])

    elif cmd in ("finish", "checkgenerate", "check-generate", "cg"):
        project_name = args[1] if len(args) > 1 else None
        cmd_finish(project_name)

    else:
        print(f"알 수 없는 명령: {cmd}")
        print("사용 가능한 명령: learn | scan | prepare | validate | generate")
        print("콤보 명령:        scanprepare (scan+prepare)  |  finish (validate+generate)")
        sys.exit(1)


if __name__ == "__main__":
    main()

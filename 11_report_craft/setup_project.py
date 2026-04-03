"""
setup_project.py — 새 프로젝트 초기화
======================================
수신자료/ 폴더를 스캔하여 config.py TARGETS를 자동 생성하고
sections_cause.py / sections.py / 참고예시/를 소스 프로젝트에서 복사합니다.

사용법:
    python setup_project.py 새프로젝트명
    python setup_project.py 새프로젝트명 --source 창원용원  # 복사 기준 프로젝트 (기본값: 창원용원)
    python setup_project.py 새프로젝트명 --rescan           # 기존 프로젝트 TARGETS 재스캔
"""

import sys
import io
import re
import shutil
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE = Path(__file__).parent

SUPPORTED_EXT = {
    ".pdf": "pdf",
    ".xlsx": "xlsx",
    ".xls": "xlsx",
    ".hwp": "hwp",
    ".hwpx": "hwp",
}


def scan_targets(수신자료_dir: Path) -> list[tuple]:
    """수신자료/ 폴더를 재귀 스캔하여 TARGETS 엔트리 생성."""
    targets = []

    files = sorted(수신자료_dir.rglob("*"))
    for idx, f in enumerate(files, start=1):
        if not f.is_file():
            continue
        ext = f.suffix.lower()
        if ext not in SUPPORTED_EXT:
            continue
        mode = SUPPORTED_EXT[ext]
        rel = f.relative_to(수신자료_dir)
        safe_name = f.stem[:30].replace(" ", "_").replace("/", "_")
        out_name = f"{idx:02d}_{safe_name}.txt"
        targets.append((out_name, str(rel), mode, f.stem[:40]))

    return targets


def build_targets_lines(targets: list) -> list[str]:
    """TARGETS 리스트 본문 라인들을 반환 (TARGETS = [ ... ] 포함)."""
    lines = [
        '# ── 추출 대상 파일 목록 ─────────────────────────────────',
        '# (출력파일명, 수신자료/ 내 상대경로, 추출방식, 비고)',
        'TARGETS = [',
    ]
    for out_name, rel, mode, note in targets:
        rel_escaped = rel.replace('\\', '/')
        lines.append(f'    ("{out_name}", "{rel_escaped}", "{mode}", "{note}"),')
    lines.append(']')
    return lines


def write_config(project_dir: Path, project_name: str, targets: list):
    lines = [
        '"""',
        f'config.py — {project_name} 프로젝트 설정',
        '=' * 42,
        '프로젝트별 설정만 여기서 수정하세요.',
        '"""',
        '',
        'SOURCE_DIR = "수신자료"',
        '',
        '# ── 프로젝트 메타데이터 ──────────────────────────────────',
        f'PROJECT_NAME    = "{project_name}"',
        'PLAINTIFF       = ""   # TODO: 원고명',
        'DEFENDANT       = ""   # TODO: 피고명',
        'CASE_NUMBER     = ""   # TODO: 사건번호',
        'CONTRACT_DATE   = ""   # TODO: 계약일',
        'CONTRACT_AMOUNT = ""   # TODO: 계약금액',
        '',
        'REFERENCE_DOC = ""   # TODO: 발주자측 감정보고서 processed 파일명',
        '',
        '# ── 감정보고서 페이지 범위 — 실제 보고서 확인 후 수정 ────',
        'GDOC_CAUSE_PAGES      = set(range(10,  91))',
        'GDOC_31_32_PAGES      = set(range(91,  133))',
        'GDOC_REST_PAGES       = set(range(446, 492))',
        'GDOC_SUM_PAGES        = set(range(491, 514))',
        'GDOC_GITUPIP_1_PAGES  = set(range(133, 290))',
        'GDOC_GITUPIP_2_PAGES  = set(range(290, 446))',
        '',
        '# ── 손실금액 확인 금액 ──────────────────────────────────',
        'CLAIM_AMOUNTS = {',
        '    "기투입_합계":      "",',
        '    "집행예정_합계":    "",',
        '    "원상복구_합계":    "",',
        '    "기타손실_합계":    "",',
        '    "발주자_청구_총계": "",',
        '}',
        '',
    ]
    lines += build_targets_lines(targets)
    lines += [
        '',
        '# ── 페이지 필터 ─────────────────────────────────────────',
        '# 지정 페이지 제거 (skip)',
        'FILE_PAGE_FILTERS: dict[str, set[int]] = {',
        '    # TODO: 필요 시 추가',
        '    # "파일명_stem": set(range(시작, 끝)),',
        '}',
        '',
        '# 지정 페이지만 보존',
        'FILE_KEEP_FILTERS: dict[str, set[int]] = {',
        '    # TODO: 필요 시 추가',
        '}',
    ]

    (project_dir / "config.py").write_text("\n".join(lines), encoding="utf-8-sig")


def rescan_config(project_dir: Path, targets: list):
    """기존 config.py에서 TARGETS 블록만 교체한다."""
    config_path = project_dir / "config.py"
    original = config_path.read_text(encoding="utf-8-sig")

    # TARGETS = [ ... ] 블록을 정규식으로 찾아 교체
    # 패턴: "TARGETS = [\n" ~ 다음 "]" (단독 줄)
    new_targets_block = "\n".join(build_targets_lines(targets))

    # 패턴: TARGETS = [ 부터 단독 ] 까지 (multiline)
    pattern = re.compile(
        r'(# ── 추출 대상 파일 목록[^\n]*\n# \(출력파일명[^\n]*\n)?TARGETS\s*=\s*\[.*?\n\]',
        re.DOTALL
    )

    if pattern.search(original):
        updated = pattern.sub(new_targets_block, original, count=1)
        config_path.write_text(updated, encoding="utf-8-sig")
        print(f"[갱신] config.py TARGETS 블록 교체 완료")
    else:
        print(f"[경고] config.py에서 TARGETS 블록을 찾지 못했습니다. 수동으로 수정하세요.")


def copy_sections(project_dir: Path, source_dir: Path, filename: str):
    """소스 프로젝트의 sections 파일을 복사하고 상단에 안내 주석을 추가."""
    src = source_dir / filename
    dst = project_dir / filename

    if not src.exists():
        # 소스 파일이 없으면 빈 템플릿으로 대체
        dst.write_text(
            f'# {filename} — 섹션 정의\n'
            f'# 소스 프로젝트({source_dir.name})에 {filename}이 없어 빈 파일로 생성됨\n\n'
            'SECTIONS = []\n',
            encoding="utf-8-sig"
        )
        print(f"  [경고] {source_dir.name}/{filename} 없음 — 빈 파일로 생성")
        return

    original = src.read_text(encoding="utf-8-sig")
    header = (
        f"# ※ {source_dir.name} 프로젝트에서 복사된 파일입니다.\n"
        f"# processed 파일명, output_request 내용을 이 프로젝트에 맞게 수정하세요.\n"
        f"# 파일 참조 형식: processed: / 예시: / 법령: / (\"filtered\", 소스, CONFIG_ATTR)\n"
        f"# {'=' * 60}\n\n"
    )
    dst.write_text(header + original, encoding="utf-8-sig")


def copy_참고예시(project_dir: Path, source_dir: Path):
    """소스 프로젝트의 참고예시 파일들을 복사."""
    src_dir = source_dir / "참고예시"
    dst_dir = project_dir / "참고예시"

    if not src_dir.exists():
        print(f"  [경고] {source_dir.name}/참고예시/ 없음 — 건너뜀")
        return

    copied = 0
    for f in sorted(src_dir.glob("*.txt")):
        if f.name == "README.txt":
            continue
        shutil.copy2(f, dst_dir / f.name)
        copied += 1

    if copied:
        print(f"[복사] 참고예시 {copied}개 파일 ({source_dir.name} 기준 — 새 프로젝트 예시 생기면 교체)")
    else:
        print(f"  [경고] {source_dir.name}/참고예시/에 복사할 파일 없음")


def main():
    rescan = "--rescan" in sys.argv

    # --source 값 파싱
    source_name = "창원용원"
    raw_args = sys.argv[1:]
    for i, a in enumerate(raw_args):
        if a == "--source" and i + 1 < len(raw_args):
            source_name = raw_args[i + 1]
    args = [a for a in raw_args if not a.startswith("--") and a != source_name or
            (a == source_name and raw_args[raw_args.index(a) - 1] != "--source"
             if raw_args.index(a) > 0 else True)]
    # 단순화: -- 옵션과 그 값을 제거하고 나머지만 취함
    clean_args = []
    skip_next = False
    for a in raw_args:
        if skip_next:
            skip_next = False
            continue
        if a == "--source":
            skip_next = True
            continue
        if a.startswith("--"):
            continue
        clean_args.append(a)

    if len(clean_args) < 1:
        print("사용법: python setup_project.py 새프로젝트명 [--source 소스프로젝트] [--rescan]")
        sys.exit(1)

    project_name = clean_args[0]
    project_dir  = BASE / "projects" / project_name
    source_dir   = BASE / "projects" / source_name

    # ── --rescan 모드 ──────────────────────────────────────────
    if rescan:
        if not project_dir.exists():
            print(f"[오류] 프로젝트가 존재하지 않습니다: {project_dir}")
            sys.exit(1)

        수신자료 = project_dir / "수신자료"
        targets = scan_targets(수신자료)
        if targets:
            print(f"[스캔] 수신자료 {len(targets)}개 파일 발견")
        else:
            print("[스캔] 수신자료 파일 없음")

        rescan_config(project_dir, targets)
        print(f"\n재스캔 완료. config.py를 확인하세요: {project_dir / 'config.py'}")
        return

    # ── 신규 생성 모드 ─────────────────────────────────────────
    if project_dir.exists():
        print(f"[오류] 이미 존재하는 프로젝트입니다: {project_dir}")
        print(f"       재스캔하려면: python setup_project.py {project_name} --rescan")
        sys.exit(1)

    # 폴더 생성
    for sub in ["수신자료", "processed", "참고예시", "prompts_cause", "prompts"]:
        (project_dir / sub).mkdir(parents=True)
    print(f"[생성] 폴더 구조 생성: {project_dir}")

    # 수신자료 스캔
    수신자료 = project_dir / "수신자료"
    targets = scan_targets(수신자료)
    if targets:
        print(f"[스캔] 수신자료 {len(targets)}개 파일 발견")
    else:
        print("[스캔] 수신자료 파일 없음 — config.py TARGETS는 비어있습니다.")
        print("       수신자료/ 폴더에 파일을 넣은 후 다시 실행하거나 직접 수정하세요.")

    # config.py 생성
    write_config(project_dir, project_name, targets)
    print(f"[생성] config.py")

    # sections 복사 (소스 프로젝트 기준)
    if source_dir.exists():
        copy_sections(project_dir, source_dir, "sections_cause.py")
        copy_sections(project_dir, source_dir, "sections.py")
        print(f"[복사] sections_cause.py, sections.py ({source_name} 기준 — 내용 수정 필요)")
    else:
        print(f"  [경고] 소스 프로젝트 없음({source_name}) — sections 빈 파일로 생성")
        for fname in ["sections_cause.py", "sections.py"]:
            (project_dir / fname).write_text(f"# {fname}\nSECTIONS = []\n", encoding="utf-8-sig")

    # 참고예시 복사 (소스 프로젝트 기준)
    copy_참고예시(project_dir, source_dir)

    print(f"\n완료! 다음 단계:")
    print(f"  1. projects/{project_name}/수신자료/ 에 원본 파일 배치")
    print(f"  2. python setup_project.py {project_name} --rescan  (파일 배치 후 재스캔)")
    print(f"  3. projects/{project_name}/config.py — 메타데이터, 페이지 범위 수정")
    print(f"  4. projects/{project_name}/sections_cause.py — 쟁점에 맞게 output_request 수정")
    print(f"  5. projects/{project_name}/참고예시/ — 이 사건 예시 파일로 교체 (선택)")
    print(f"  6. python extractor.py {project_name}")
    print(f"  7. python generate_prompts_cause.py {project_name}")


if __name__ == "__main__":
    main()

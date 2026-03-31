"""
generate_prompts_cause.py — 원인·과실책임 분석 보고서 프롬프트 생성기
======================================================================
templates/prompt_cause.txt + projects/{프로젝트}/sections_cause.py
→ projects/{프로젝트}/prompts_cause/ 저장

사용법:
    python generate_prompts_cause.py 창원용원
    python generate_prompts_cause.py          # 기본값: 창원용원
"""

import re, sys, io, importlib.util
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE         = Path(__file__).parent
project_name = sys.argv[1] if len(sys.argv) > 1 else "창원용원"
PROJECT_DIR  = BASE / "projects" / project_name

# config.py 동적 로드
_spec = importlib.util.spec_from_file_location("config", PROJECT_DIR / "config.py")
config = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(config)

# sections_cause.py 동적 로드
_spec2 = importlib.util.spec_from_file_location("sections_cause", PROJECT_DIR / "sections_cause.py")
sections_mod = importlib.util.module_from_spec(_spec2)
_spec2.loader.exec_module(sections_mod)
SECTIONS = sections_mod.SECTIONS

PROCESSED = PROJECT_DIR / "processed"
LAW_DIR   = BASE / "법령_processed"
PROMPTS   = PROJECT_DIR / "prompts_cause"
PROMPTS.mkdir(exist_ok=True)


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig", errors="replace")


def filter_pages(text: str, keep_pages: set) -> str:
    """지정 페이지만 남긴다."""
    parts = re.split(r"(\[\d+페이지\])", text)
    result = []
    keeping = True
    for part in parts:
        m = re.match(r"\[(\d+)페이지\]", part)
        if m:
            keeping = int(m.group(1)) in keep_pages
            if keeping:
                result.append(part)
        else:
            if keeping:
                result.append(part)
    return "".join(result)


def resolve_file_ref(ref) -> str | None:
    """ref를 실제 텍스트 내용으로 반환. 없으면 None."""
    if isinstance(ref, tuple) and ref[0] == "filtered":
        _, src_file, pages_key = ref
        path = PROCESSED / src_file
        if not path.exists():
            print(f"  [없음] {src_file}")
            return None
        pages = getattr(config, pages_key)
        return filter_pages(read(path), pages)
    if isinstance(ref, str):
        if ref.startswith("processed:"):
            path = PROCESSED / ref[len("processed:"):]
        elif ref.startswith("예시:"):
            path = PROJECT_DIR / "참고예시" / ref[len("예시:"):]
        elif ref.startswith("법령:"):
            path = LAW_DIR / ref[len("법령:"):]
        else:
            return ref  # 이미 문자열 내용
        if not path.exists():
            print(f"  [없음] {path.name}")
            return None
        return read(path)
    return None


TEMPLATE = read(BASE / "templates" / "prompt_cause.txt")
SPLIT_MARKER = "[여기에 processed/공사명.txt 내용 전체를 붙여넣기]"
if SPLIT_MARKER in TEMPLATE:
    _parts = TEMPLATE.split(SPLIT_MARKER, 1)
    TEMPLATE_TOP    = _parts[0]
    TEMPLATE_BOTTOM = _parts[1]
else:
    TEMPLATE_TOP    = TEMPLATE
    TEMPLATE_BOTTOM = ""

DIVIDER = "=" * 64


def build_prompt(section: dict) -> str:
    lines = [TEMPLATE_TOP.rstrip()]
    lines.append(f"\n{DIVIDER}\n## 공사 정보 — {section['title']}\n{DIVIDER}\n")

    for label, ref in section["files"]:
        content = resolve_file_ref(ref)
        if content is None:
            continue
        lines.append(f"\n{label}\n{'-'*40}\n{content.strip()}\n")

    lines.append(f"\n{DIVIDER}\n")
    lines.append(section["output_request"])
    return "\n".join(lines)


total_size = 0
for sec in SECTIONS:
    prompt = build_prompt(sec)
    out_path = PROMPTS / sec["filename"]
    out_path.write_text(prompt, encoding="utf-8-sig")
    sz = len(prompt)
    total_size += sz
    print(f"  [생성] {sz:>10,}자  {sec['filename']}")

print(f"\n프로젝트: {project_name}")
print(f"저장 위치: {PROMPTS}")
print(f"총 생성: {total_size:,}자 ({len(SECTIONS)}개 파일)")

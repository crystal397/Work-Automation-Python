"""
분석 입출력 관리자 — API 없음
역할:
  1. 추출된 텍스트를 Claude Code 분석용 파일로 저장
  2. Claude Code가 작성한 analysis_result.json 로드·검증
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
import config
from src.extractor.file_extractor import ExtractResult
from .prompts import ANALYSIS_SCHEMA_GUIDE


# ── 분석용 파일 준비 ──────────────────────────────────────────────────────────────

def prepare(results: list[ExtractResult],
            output_dir: str = "output") -> Path:
    """
    추출 결과를 Claude Code 분석용 마크다운으로 저장.
    저장 후 다음 단계 안내 메시지 출력.
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    path = out / "extracted_for_analysis.md"

    # 분석지시서 내용을 헤더에 포함
    guide_path = Path(__file__).parent.parent.parent / "분석지시서.md"
    guide_text = guide_path.read_text(encoding="utf-8") if guide_path.exists() else ""

    lines = [
        "# 공기연장 간접비 보고서 — 분석 대상 자료",
        "",
        "아래 자료를 분석하여 `analysis_result.json` 파일을 작성해 주세요.",
        "**작성 전에 아래 [분석 지시서] 와 문서 하단 [JSON 작성 가이드] 를 반드시 읽으세요.**",
        "",
    ]
    if guide_text:
        lines += [
            "---",
            "",
            "## [분석 지시서]",
            "",
            guide_text,
            "",
        ]
    lines += [
        "---",
        "",
        "## 추출된 자료",
        "",
    ]

    # 파일별 품질 요약
    lines += ["| 파일명 | 품질 | 비고 |", "|--------|------|------|"]
    for r in results:
        icon = {"OK": "✅", "WARN": "⚠️", "FAIL": "❌"}[r.quality]
        note = r.issues[0] if r.issues else ""
        lines.append(f"| {r.file} | {icon} {r.quality} | {note} |")
    lines += ["", "---", ""]

    # 출처 포함 전체 텍스트
    for r in results:
        lines += [f"### [{r.file}]  (품질: {r.quality})", ""]
        if r.quality == "FAIL":
            lines += [
                "> ⚠️ 이 파일은 모든 추출 방법 소진 후에도 내용을 읽지 못했습니다.",
                "> 관련 항목은 `[출처 확인 필요]`로 표시하세요.",
                ""
            ]
        else:
            for c in r.chunks:
                if c.text.strip():
                    quality_note = f"  <!-- {c.quality}: {c.note} -->" if c.quality != "OK" else ""
                    lines += [
                        f"**[출처: {c.source}]**{quality_note}",
                        "",
                        c.text.strip(),
                        "",
                    ]
        lines += ["---", ""]

    # JSON 작성 가이드
    lines += ["", "---", "", "## [JSON 작성 가이드]", "", ANALYSIS_SCHEMA_GUIDE]

    path.write_text("\n".join(lines), encoding="utf-8")

    print(f"\n📄 분석용 파일 저장 완료: {path}")
    print("\n" + "=" * 60)
    print("다음 단계 — Claude Code에 아래 문장을 그대로 입력하세요:\n")
    print(f'  "{path.name} 파일을 읽고,')
    print(f"   분석지시서.md의 지침에 따라")
    print(f'   analysis_result.json 으로 저장해줘"')
    print("=" * 60 + "\n")

    return path


# ── 분석 결과 로드 ────────────────────────────────────────────────────────────────

def load(output_dir: str = "output") -> dict:
    """
    Claude Code가 작성한 analysis_result.json 로드.
    파일이 없거나 형식이 잘못된 경우 안내 메시지 출력.
    """
    path = Path(output_dir) / "analysis_result.json"

    if not path.exists():
        print(f"\n❌ {path} 파일이 없습니다.")
        print("먼저 Claude Code에서 extracted_for_analysis.md 를 분석하여")
        print("analysis_result.json 을 생성해 주세요.\n")
        sys.exit(1)

    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    _validate(data)
    print(f"✅ 분석 결과 로드: {path}")
    print(f"   유형: {data.get('report_type', '?')} | "
          f"미확인 항목: {len(data.get('unresolved', []))}개")
    return data


# ── 검증 ─────────────────────────────────────────────────────────────────────────

REQUIRED_KEYS = ["report_type", "contract", "extension", "rates"]

def _validate(data: dict) -> None:
    missing = [k for k in REQUIRED_KEYS if k not in data]
    if missing:
        print(f"⚠️  analysis_result.json 에 필수 항목 누락: {missing}")
        print("   계산 결과가 불완전할 수 있습니다.")

    rtype = data.get("report_type")
    if rtype not in ("A", "B", "C"):
        print(f"⚠️  report_type 값이 A/B/C 가 아닙니다: {rtype!r}")

    # 미확인 항목 안내
    unresolved = data.get("unresolved", [])
    if unresolved:
        print(f"\n⚠️  미확인 항목 {len(unresolved)}개 — 보고서에 주의 표시됩니다:")
        for u in unresolved:
            print(f"   • {u.get('item')}: {u.get('reason')}")
        print()

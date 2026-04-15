"""
추출 품질 검증기
- 모든 폴백 시도 이후의 최종 결과를 보고
- 중단 없이 항상 extraction_report.md 생성
- FAIL 항목도 보고서 내 주의 표시로만 처리
"""

from datetime import datetime
from pathlib import Path
from .file_extractor import ExtractResult


def check_and_report(results: list[ExtractResult],
                     output_dir: str = "output") -> bool:
    """
    품질 보고서 생성. 중단하지 않고 항상 True 반환.
    FAIL 항목은 보고서에 주의 표시 후 계속 진행.
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    report_path = Path(output_dir) / "extraction_report.md"

    ok_list   = [r for r in results if r.quality == "OK"]
    warn_list = [r for r in results if r.quality == "WARN"]
    fail_list = [r for r in results if r.quality == "FAIL"]

    lines = _build_report(results, ok_list, warn_list, fail_list)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\n📄 품질 보고서: {report_path}")
    if fail_list:
        print(f"\n   ❌ 추출 실패 {len(fail_list)}개 파일 — 포맷 변환 후 재추출 권장:")
        for r in fail_list:
            action = _suggest_action(r)
            print(f"      • {r.file}")
            print(f"        → {action}")
        print(f"\n   변환 후 같은 input 폴더에 추가하고 'python main.py extract' 재실행")
        print(f"   (기존 파일은 캐시로 건너뛰고 새 파일만 처리)")
    return True   # 중단 없이 항상 계속


def _build_report(results, ok_list, warn_list, fail_list) -> list[str]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        "# 추출 품질 보고서",
        f"생성일시: {now}",
        "",
        "> 각 파일은 가능한 모든 방법을 순서대로 시도했습니다.",
        "> WARN/FAIL은 모든 방법 소진 후의 최종 결과입니다.",
        "",
        "---",
        "",
        "## 요약",
        "",
        "| 구분 | 건수 |",
        "|------|------|",
        f"| 전체 파일 | {len(results)}개 |",
        f"| ✅ 정상 (OK) | {len(ok_list)}개 |",
        f"| ⚠️ 일부 구간 미확인 (WARN) | {len(warn_list)}개 |",
        f"| ❌ 추출 실패 — 수동 확인 필요 (FAIL) | {len(fail_list)}개 |",
        "",
    ]

    # ── 정상 ──
    if ok_list:
        lines += ["---", "", "## ✅ 정상 파일", "",
                  "| 파일명 | 포맷 | 청크 수 | 추출 분량 |",
                  "|--------|------|---------|----------|"]
        for r in ok_list:
            chars = sum(len(c.text) for c in r.chunks)
            lines.append(f"| {r.file} | {r.format} | {len(r.chunks)} | {chars:,}자 |")
        lines.append("")

    # ── WARN ──
    if warn_list:
        lines += [
            "---", "",
            "## ⚠️ 일부 구간 미확인 (WARN)", "",
            "> 가능한 방법을 모두 시도했으나 일부 구간의 품질이 낮습니다.",
            "> 보고서 해당 항목을 수동으로 검토하세요.", "",
        ]
        for r in warn_list:
            lines += [f"### {r.file}", "",
                      "| 출처 | 시도한 방법 | 최종 결과 | 비고 |",
                      "|------|-----------|----------|------|"]
            for c in r.chunks:
                if c.quality != "OK":
                    tried = " → ".join(c.tried) if c.tried else c.method
                    lines.append(
                        f"| {c.source} | {tried} | {c.quality} | {c.note} |"
                    )
            lines.append("")

    # ── FAIL ──
    if fail_list:
        lines += [
            "---", "",
            "## ❌ 추출 실패 — 조치 필요 (FAIL)", "",
            "> 아래 파일은 모든 추출 방법을 시도했으나 내용을 읽지 못했습니다.",
            "> **포맷별 권장 조치** 를 따라 파일을 변환한 뒤 `python main.py extract` 를 재실행하세요.",
            "> 변환된 파일을 원본과 **같은 폴더에 추가**하면 캐시된 파일은 건너뛰고 새 파일만 처리합니다.",
            "",
            "| 파일명 | 시도한 방법 | 원인 | 권장 조치 |",
            "|--------|-----------|------|----------|",
        ]
        for r in fail_list:
            # 모든 청크에서 tried 수집
            all_tried = []
            for c in r.chunks:
                for m in (c.tried or [c.method]):
                    if m not in all_tried:
                        all_tried.append(m)
            tried_str  = " → ".join(all_tried) if all_tried else "알 수 없음"
            cause      = r.issues[0] if r.issues else "알 수 없음"
            action     = _suggest_action(r)
            lines.append(f"| {r.file} | {tried_str} | {cause} | {action} |")
        lines += [
            "",
            "**포맷별 변환 방법 (재추출 불필요 — 변환 후 input 폴더에 추가만)**",
            "",
            "| 포맷 | 변환 방법 |",
            "|------|----------|",
            "| HWP | 한글 → 파일 → 다른 이름으로 저장 → PDF 선택 |",
            "| 스캔 PDF / TIF | 300dpi 이상 재스캔하거나 원본 전자 파일 사용 |",
            "| 보안 PDF | PDF 작성자에게 보안 해제 요청, 또는 인쇄 → PDF 저장 |",
            "| XLS | Excel에서 열고 xlsx로 다시 저장 |",
            "",
        ]

    # ── 출처 인덱스 ──
    lines += [
        "---", "",
        "## 출처 인덱스 (전체 청크)", "",
        "| 파일명 | 출처 | 사용 방법 | 품질 | 내용 미리보기 |",
        "|--------|------|----------|------|-------------|",
    ]
    for r in results:
        for c in r.chunks:
            preview = c.text[:50].replace("\n", " ") if c.text else "(없음)"
            lines.append(
                f"| {r.file} | {c.source} | {c.method} | {c.quality} | {preview}... |"
            )
    lines.append("")

    return lines


def _suggest_action(result: ExtractResult) -> str:
    fmt = result.format
    if fmt == "hwp":
        return "한글에서 '다른 이름으로 저장 → PDF' 후 재투입"
    if fmt in ("tif", "tiff", "image"):
        return "300dpi 이상으로 재스캔, 또는 PDF로 변환 후 재투입"
    if fmt == "pdf":
        return "보안 잠금 해제 또는 텍스트 레이어 포함 여부 확인"
    if fmt == "excel":
        return ".xlsx로 재저장 후 재투입"
    return "파일 손상 여부 확인 후 재투입"


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))
    from src.extractor.file_extractor import extract_folder

    input_dir  = sys.argv[1] if len(sys.argv) > 1 else "input"
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    results    = extract_folder(input_dir)
    check_and_report(results, output_dir)

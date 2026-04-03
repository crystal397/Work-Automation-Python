"""
extractor.py — 보고서 작성용 핵심 파일 텍스트 추출기
=====================================================
projects/{프로젝트명}/수신자료/ 내 지정 파일을 텍스트 추출
→ projects/{프로젝트명}/processed/ 저장

사용법:
    python extractor.py 창원용원
    python extractor.py          # 기본값: 창원용원
"""

import io
import sys
import re
import importlib.util
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE         = Path(__file__).parent
project_name = sys.argv[1] if len(sys.argv) > 1 else "창원용원"
PROJECT_DIR  = BASE / "projects" / project_name

# config.py 동적 로드
_spec = importlib.util.spec_from_file_location("config", PROJECT_DIR / "config.py")
config = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(config)

SRC    = PROJECT_DIR / config.SOURCE_DIR
OUT    = PROJECT_DIR / "processed"

TARGETS           = config.TARGETS
FILE_PAGE_FILTERS = config.FILE_PAGE_FILTERS
FILE_KEEP_FILTERS = config.FILE_KEEP_FILTERS


def filter_pages(text: str, skip_pages: set[int]) -> str:
    """[N페이지] 마커 기준으로 지정 페이지를 제거한다."""
    parts = re.split(r"(\[\d+페이지\])", text)
    result = []
    skipping = False
    for part in parts:
        m = re.match(r"\[(\d+)페이지\]", part)
        if m:
            skipping = int(m.group(1)) in skip_pages
            if not skipping:
                result.append(part)
        else:
            if not skipping:
                result.append(part)
    return "".join(result)


def filter_pages_keep(text: str, keep_pages: set[int]) -> str:
    """[N페이지] 마커 기준으로 지정 페이지만 보존한다."""
    parts = re.split(r"(\[\d+페이지\])", text)
    result = []
    keeping = False
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


# ──────────────────────────────────────────────────────────────
# 추출 함수
# ──────────────────────────────────────────────────────────────
def _is_garbled(text: str) -> bool:
    """추출 텍스트가 깨진 인코딩인지 판정.
    한글·영숫자·기본 구두점 비율이 30% 미만이면 깨진 것으로 본다."""
    if not text or len(text) < 50:
        return False
    valid = sum(
        1 for c in text
        if '\uAC00' <= c <= '\uD7A3'          # 한글 완성형
        or '\u1100' <= c <= '\u11FF'          # 한글 자모
        or '\u3130' <= c <= '\u318F'          # 한글 호환 자모
        or (c.isascii() and (c.isprintable() or c in '\n\t '))
    )
    return (valid / len(text)) < 0.3


def _extract_pdf_pdfminer(path: Path) -> str:
    """pdfminer.six 를 이용한 PDF 텍스트 추출 (CID 폰트 폴백용)."""
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTTextContainer
    except ImportError:
        return ""
    try:
        pages = []
        for i, page_layout in enumerate(extract_pages(str(path))):
            texts = [el.get_text() for el in page_layout if isinstance(el, LTTextContainer)]
            txt = "".join(texts).strip()
            if txt:
                pages.append(f"[{i+1}페이지]\n{txt}")
        return "\n\n".join(pages)
    except Exception:
        return ""


def extract_pdf(path: Path) -> str:
    import fitz
    doc = fitz.open(path)
    pages = []
    for i in range(doc.page_count):
        txt = doc[i].get_text().strip()
        if txt:
            pages.append(f"[{i+1}페이지]\n{txt}")
    doc.close()
    result = "\n\n".join(pages)

    if _is_garbled(result):
        fallback = _extract_pdf_pdfminer(path)
        if fallback and not _is_garbled(fallback):
            print(f"    ↳ [폴백] pdfminer로 재추출 (PyMuPDF 인코딩 오류 감지)")
            return fallback

    return result


def extract_xlsx(path: Path) -> str:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    parts = []
    for ws in wb.worksheets:
        parts.append(f"=== 시트: {ws.title} ===")
        for row in ws.iter_rows(values_only=True):
            line = "\t".join("" if v is None else str(v) for v in row)
            if line.strip():
                parts.append(line)
    wb.close()
    return "\n".join(parts)


def extract_hwp(path: Path) -> str:
    """HWP 텍스트 추출 — 한글(win32com) HWP→PDF→PyMuPDF 파이프라인"""
    import shutil, tempfile

    try:
        import win32com.client, fitz
        tmp_dir = Path(tempfile.mkdtemp())
        tmp_hwp = tmp_dir / "input.hwp"
        shutil.copy2(path, tmp_hwp)

        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
        # 팝업 메시지 자동 승인
        try:
            hwp.XHwpApplication.SetMessageBoxMode(65535)
        except Exception:
            pass
        hwp.Open(str(tmp_hwp.resolve()))

        out_pdf = tmp_dir / "output.pdf"
        hwp.SaveAs(str(out_pdf.resolve()), "PDF")
        hwp.Quit()

        doc = fitz.open(str(out_pdf))
        pages = []
        n = doc.page_count
        for i in range(n):
            txt = doc[i].get_text().strip()
            if txt:
                pages.append(f"[{i+1}페이지]\n{txt}")
        doc.close()
        shutil.rmtree(tmp_dir, ignore_errors=True)

        text = "\n\n".join(pages)
        if text.strip():
            return text
    except Exception:
        pass

    return "[HWP 추출 실패 — 한글 프로그램 또는 win32com 오류]"


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main():
    OUT.mkdir(exist_ok=True)

    ok = skip = fail = 0
    for out_name, rel_path, mode, note in TARGETS:
        src = SRC / rel_path
        out = OUT / out_name

        if not src.exists():
            print(f"  [없음]  {out_name}  ← {rel_path[:60]}")
            skip += 1
            continue

        try:
            if mode == "pdf":
                text = extract_pdf(src)
            elif mode == "xlsx":
                text = extract_xlsx(src)
            elif mode == "hwp":
                text = extract_hwp(src)
            else:
                text = src.read_text(encoding="utf-8-sig", errors="replace")

            # 파일별 페이지 필터 적용
            out_stem = out_name.replace(".txt", "")
            for key, skip_set in FILE_PAGE_FILTERS.items():
                if key in out_stem:
                    before = len(text)
                    text = filter_pages(text, skip_set)
                    print(f"  [필터]  {before:,}자 → {len(text):,}자  ({key} p{min(skip_set)}-{max(skip_set)} 제거)")
                    break
            for key, keep_set in FILE_KEEP_FILTERS.items():
                if key in out_stem:
                    before = len(text)
                    text = filter_pages_keep(text, keep_set)
                    print(f"  [보존]  {before:,}자 → {len(text):,}자  ({key} p{min(keep_set)}-{max(keep_set)} 만 보존)")
                    break

            header = f"# {note}\n# 원본: {src.name}\n\n"
            out.write_text(header + text, encoding="utf-8-sig")
            print(f"  [OK]    {len(text):>8,}자  {out_name}  ({note})")
            ok += 1

        except Exception as e:
            print(f"  [오류]  {out_name}: {e}")
            fail += 1

    print(f"\n프로젝트: {project_name}")
    print(f"완료: 성공 {ok}건 / 없음 {skip}건 / 오류 {fail}건")
    print(f"저장 위치: {OUT}")


if __name__ == "__main__":
    main()

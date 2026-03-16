"""
PDF 자동 합치기 + 갑지 생성 도구
─────────────────────────────────
- 지정 폴더를 재귀 순회하며 폴더/파일명으로 갑지(표지) 자동 생성
- PDF / Excel(.xlsx .xls .xlsm) 파일을 순서대로 하나의 PDF로 합침
- Excel은 Microsoft Excel(설치 필요)로 PDF 변환

갑지 규칙:
  - 첫 번째 갑지 : "# 첨부자료"  (고정)
  - 두 번째 갑지 : "추가공사비 집계표"  (고정)
  - 이후 갑지    : 계층 번호 + 이름 자동 생성
                   1단계 폴더       →  "1. 업무지시서"
                   2단계 폴더       →  "1.1. 수직구"
                   3단계 파일       →  "1.1.1. 내역서"
  - 선(구분선) 없음
  - 확장자(.pdf .xlsx 등) 표시 안 함
  - 폰트: 한컴바탕 20pt (없으면 맑은 고딕으로 fallback)

필요 패키지 설치 (최초 1회):
    pip install reportlab pypdf pywin32

사용법:
    python pdf_merger.py <폴더경로>
    python pdf_merger.py <폴더경로> <출력파일.pdf>

예시:
    python pdf_merger.py C:\\산출물
    python pdf_merger.py C:\\산출물 C:\\결과.pdf
"""

import os
import re
import sys
import tempfile
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from pypdf import PdfWriter, PdfReader


# ──────────────────────────────────────────────
# 한글 폰트 등록 (한컴바탕 우선)
# ──────────────────────────────────────────────
def register_font():
    candidates = [
        # 한컴바탕 (한컴오피스 설치 시 생성되는 경로들)
        (r"C:\Windows\Fonts\HCR Batang.ttf",  "HCRBatang"),
        (r"C:\Windows\Fonts\HCRBatang.ttf",   "HCRBatang"),
        (r"C:\Program Files\HNC\Shared\HncFonts\HCR Batang.ttf", "HCRBatang"),
        (r"C:\Program Files (x86)\HNC\Shared\HncFonts\HCR Batang.ttf", "HCRBatang"),
        # 맑은 고딕 Bold (fallback)
        (r"C:\Windows\Fonts\malgunbd.ttf",     "MalgunGothicBold"),
        (r"C:\Windows\Fonts\malgun.ttf",       "MalgunGothic"),
        # 나눔명조 (fallback)
        (r"C:\Windows\Fonts\NanumMyeongjoBold.ttf", "NanumMyeongjoBold"),
        (r"C:\Windows\Fonts\NanumMyeongjo.ttf",     "NanumMyeongjo"),
    ]
    for path, name in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                print(f"[폰트] {name} 등록 완료")
                return name
            except Exception:
                continue
    print("[폰트] 한글 폰트를 찾지 못했습니다 → 한글이 깨질 수 있습니다.")
    print("       한컴오피스 또는 나눔폰트를 설치한 뒤 다시 실행하세요.")
    return "Helvetica"


FONT      = register_font()
FONT_SIZE = 20


# ──────────────────────────────────────────────
# 파일명 처리 유틸
# ──────────────────────────────────────────────
def remove_ext(name: str) -> str:
    """PDF/Excel 확장자만 제거. 마침표가 포함된 한글 폴더명은 건드리지 않음."""
    return re.sub(r'\.(pdf|xlsx|xls|xlsm)$', '', name, flags=re.IGNORECASE).strip()


def extract_label(name: str) -> str:
    """
    파일/폴더명에서 맨 앞 번호를 제거하고 실제 이름만 반환.
      '1. 업무지시서'   →  '업무지시서'
      '1.1. 수직구'     →  '수직구'
      '1. 내역서.pdf'   →  '내역서'
      '도면.pdf'        →  '도면'
    """
    base   = remove_ext(name)
    result = re.sub(r'^\d+(?:\.\d+)*\.?\s*', '', base).strip()
    return result if result else base


# ──────────────────────────────────────────────
# 파일명 앞 숫자 기준 정렬
# ──────────────────────────────────────────────
def sort_key(entry: Path):
    m = re.match(r'^(\d+(?:\.\d+)*)[.\s]', entry.name)
    if m:
        return tuple(int(x) for x in m.group(1).split('.'))
    return (9999,)


def sorted_entries(folder: Path):
    return sorted(folder.iterdir(), key=sort_key)


# ──────────────────────────────────────────────
# 계층 번호 문자열 생성
# ──────────────────────────────────────────────
def make_title(index_stack: list, label: str) -> str:
    """
    index_stack=[1, 2, 3], label='내역서'
    → '1.2.3. 내역서'
    """
    num = '.'.join(str(i) for i in index_stack) + '.'
    return f"{num} {label}"


# ──────────────────────────────────────────────
# 갑지(표지) PDF 생성
# ──────────────────────────────────────────────
def make_cover(title: str, tmp_dir: str) -> str:
    """갑지 한 장짜리 PDF를 만들어 경로를 반환합니다."""
    safe = re.sub(r'[<>:"/\\|?*\s#]', '_', title)[:50]
    out  = os.path.join(tmp_dir, f"cover_{safe}_{abs(hash(title)) % 99999}.pdf")

    w, h = A4   # 595pt × 841pt

    c = canvas.Canvas(out, pagesize=A4)

    # 흰 배경
    c.setFillColorRGB(1, 1, 1)
    c.rect(0, 0, w, h, fill=1, stroke=0)

    # 텍스트: 너비에 맞게 자동 줄바꿈 + 수직/수평 중앙 정렬
    c.setFont(FONT, FONT_SIZE)
    c.setFillColorRGB(0, 0, 0)

    max_w = w - 80 * mm
    lines, cur = [], ""
    for ch in title:
        if c.stringWidth(cur + ch, FONT, FONT_SIZE) <= max_w:
            cur += ch
        else:
            lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)

    line_h  = FONT_SIZE * 1.6
    total_h = len(lines) * line_h
    y       = h / 2 + total_h / 2 - FONT_SIZE

    for line in lines:
        tw = c.stringWidth(line, FONT, FONT_SIZE)
        c.drawString((w - tw) / 2, y, line)
        y -= line_h

    # 선 없음 (요청사항)

    c.save()
    return out


# ──────────────────────────────────────────────
# Excel → PDF 변환 (Microsoft Excel COM 사용)
# ──────────────────────────────────────────────
def excel_to_pdf(excel_path: str, tmp_dir: str) -> str | None:
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("  [오류] pywin32가 없습니다. → pip install pywin32")
        return None

    base    = re.sub(r'\.(xlsx|xls|xlsm)$', '', os.path.basename(excel_path), flags=re.IGNORECASE)
    out_pdf = os.path.abspath(os.path.join(tmp_dir, base + ".pdf"))
    abs_src = os.path.abspath(excel_path)
    xl      = None

    try:
        pythoncom.CoInitialize()
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible       = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(abs_src)
        wb.ExportAsFixedFormat(0, out_pdf)   # 0 = xlTypePDF (전체 시트)
        wb.Close(False)
        xl.Quit()
        pythoncom.CoUninitialize()
        return out_pdf if os.path.exists(out_pdf) else None
    except Exception as e:
        print(f"  [Excel 변환 오류] {os.path.basename(excel_path)}: {e}")
        try:
            xl.Quit()
        except Exception:
            pass
        return None


# ──────────────────────────────────────────────
# 폴더 재귀 순회 → PdfWriter에 페이지 추가
# ──────────────────────────────────────────────
def collect(folder: Path, index_stack: list, tmp_dir: str,
            writer: PdfWriter, stats: dict):
    """
    index_stack : 현재까지의 번호 리스트
      루트 직속 1번 폴더  →  [1]
      그 안의 2번 폴더    →  [1, 2]
      그 안의 3번 파일    →  [1, 2, 3]
    """
    idx = 0

    for entry in sorted_entries(folder):
        name = entry.name

        # 숨김·임시 파일 건너뜀
        if name.startswith('.') or name.startswith('~$'):
            continue

        label  = extract_label(name)
        ext    = entry.suffix.lower()
        indent = "  " * len(index_stack)

        # ── 폴더 ──
        if entry.is_dir():
            idx += 1
            stack = index_stack + [idx]
            title = make_title(stack, label)
            print(f"{indent}📁 {title}")
            add_pdf(writer, make_cover(title, tmp_dir))
            stats['covers'] += 1
            collect(entry, stack, tmp_dir, writer, stats)

        # ── PDF ──
        elif ext == '.pdf':
            idx += 1
            stack = index_stack + [idx]
            title = make_title(stack, label)
            print(f"{indent}📄 {title}")
            add_pdf(writer, make_cover(title, tmp_dir))
            stats['covers'] += 1
            add_pdf(writer, str(entry))
            stats['pdfs'] += 1

        # ── Excel ──
        elif ext in ('.xlsx', '.xls', '.xlsm'):
            idx += 1
            stack = index_stack + [idx]
            title = make_title(stack, label)
            print(f"{indent}📊 {title}")
            converted = excel_to_pdf(str(entry), tmp_dir)
            if converted:
                add_pdf(writer, make_cover(title, tmp_dir))
                stats['covers'] += 1
                add_pdf(writer, converted)
                stats['excels'] += 1
            else:
                stats['failed'] += 1


def add_pdf(writer: PdfWriter, path: str):
    try:
        for page in PdfReader(path).pages:
            writer.add_page(page)
    except Exception as e:
        print(f"  [경고] 추가 실패: {path}  ({e})")


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def run(root_folder: str, output_pdf: str | None = None):
    root = Path(root_folder)
    if not root.is_dir():
        print(f"오류: '{root_folder}' 폴더를 찾을 수 없습니다.")
        sys.exit(1)

    if output_pdf is None:
        output_pdf = str(root.parent / (root.name + "_merged.pdf"))

    print("=" * 55)
    print(f"  루트 폴더 : {root}")
    print(f"  출력 파일 : {output_pdf}")
    print("=" * 55)

    writer = PdfWriter()
    stats  = {'covers': 0, 'pdfs': 0, 'excels': 0, 'failed': 0}

    with tempfile.TemporaryDirectory() as tmp:

        # 고정 갑지 1, 2
        print("📋 갑지 #1 : # 첨부자료")
        add_pdf(writer, make_cover("# 첨부자료", tmp))
        stats['covers'] += 1

        print("📋 갑지 #2 : 추가공사비 집계표")
        add_pdf(writer, make_cover("추가공사비 집계표", tmp))
        stats['covers'] += 1

        # 폴더 순회
        collect(root, index_stack=[], tmp_dir=tmp, writer=writer, stats=stats)

        if not writer.pages:
            print("\n처리된 파일이 없습니다. 폴더 안에 PDF/Excel 파일이 있는지 확인해 주세요.")
            return

        print(f"\n저장 중... ({len(writer.pages)}페이지)")
        with open(output_pdf, "wb") as f:
            writer.write(f)

    print(f"\n✅ 완료!")
    print(f"   총 페이지  : {len(writer.pages)}")
    print(f"   갑지 수    : {stats['covers']}")
    print(f"   PDF        : {stats['pdfs']}개")
    print(f"   Excel      : {stats['excels']}개")
    if stats['failed']:
        print(f"   변환 실패  : {stats['failed']}개")
    print(f"\n   → {output_pdf}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    run(sys.argv[1], sys.argv[2] if len(sys.argv) >= 3 else None)

"""
PDF 자동 합치기 + 갑지 생성 도구
"""

import os
import re
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from pypdf import PdfWriter, PdfReader


# ──────────────────────────────────────────────
# 한글 폰트 등록
# ──────────────────────────────────────────────
def register_font():
    candidates = [
        (r"C:\Windows\Fonts\HCR Batang.ttf",  "HCRBatang"),
        (r"C:\Windows\Fonts\HCRBatang.ttf",   "HCRBatang"),
        (r"C:\Program Files\HNC\Shared\HncFonts\HCR Batang.ttf", "HCRBatang"),
        (r"C:\Program Files (x86)\HNC\Shared\HncFonts\HCR Batang.ttf", "HCRBatang"),
        (r"C:\Windows\Fonts\malgunbd.ttf",     "MalgunGothicBold"),
        (r"C:\Windows\Fonts\malgun.ttf",       "MalgunGothic"),
        (r"C:\Windows\Fonts\NanumMyeongjoBold.ttf", "NanumMyeongjoBold"),
        (r"C:\Windows\Fonts\NanumMyeongjo.ttf",     "NanumMyeongjo"),
    ]
    for path, name in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                return name
            except Exception:
                continue
    return "Helvetica"


FONT      = register_font()
FONT_SIZE = 20


# ──────────────────────────────────────────────
# 파일명 파싱 유틸
# ──────────────────────────────────────────────
def remove_ext(name):
    return re.sub(r'\.(pdf|xlsx|xls|xlsm)$', '', name, flags=re.IGNORECASE).strip()

def parse_file_number(name):
    base = remove_ext(name)
    m = re.match(r'^(\d+)((?:-\d+)+)[.\s]', base)
    if m:
        return int(m.group(1)), [int(x) for x in m.group(2).split('-') if x]
    m = re.match(r'^(\d+(?:\.\d+)*)[.\s]', base)
    if m:
        return int(m.group(1).split('.')[0]), []
    return None, []

def extract_label(name):
    base   = remove_ext(name)
    result = re.sub(r'^\d+(?:[.\-]\d+)*[.\s]\s*', '', base).strip()
    return result if result else base

def sort_key(entry):
    main, subs = parse_file_number(entry.name)
    if main is None:
        return (9999, 0, 0)
    return tuple([main] + subs + [0])

def sorted_entries(folder):
    return sorted(folder.iterdir(), key=sort_key)

def make_title(index_stack, label):
    return '.'.join(str(i) for i in index_stack) + '. ' + label


# ──────────────────────────────────────────────
# 갑지 생성
# ──────────────────────────────────────────────
def make_cover(title, tmp_dir):
    safe = re.sub(r'[<>:"/\\|?*\s#]', '_', title)[:50]
    out  = os.path.join(tmp_dir, f"cover_{safe}_{abs(hash(title)) % 99999}.pdf")
    w, h = A4
    c = canvas.Canvas(out, pagesize=A4)
    c.setFillColorRGB(1, 1, 1)
    c.rect(0, 0, w, h, fill=1, stroke=0)
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
    c.save()
    return out


# ──────────────────────────────────────────────
# Excel → PDF 변환
# ──────────────────────────────────────────────
def excel_to_pdf(excel_path, tmp_dir):
    try:
        import win32com.client, pythoncom
    except ImportError:
        return None
    base    = re.sub(r'\.(xlsx|xls|xlsm)$', '', os.path.basename(excel_path), flags=re.IGNORECASE)
    out_pdf = os.path.abspath(os.path.join(tmp_dir, base + ".pdf"))
    abs_src = os.path.abspath(excel_path)
    xl = None
    try:
        pythoncom.CoInitialize()
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            xl.AutomationSecurity = 1
        except Exception:
            pass
        try:
            for i in range(xl.Application.ProtectedViewWindows.Count, 0, -1):
                xl.Application.ProtectedViewWindows.Item(i).Close(True)
        except Exception:
            pass
        wb = xl.Workbooks.Open(abs_src, UpdateLinks=0, ReadOnly=True, CorruptLoad=2)
        wb.ExportAsFixedFormat(0, out_pdf)
        wb.Close(False)
        xl.Quit()
        pythoncom.CoUninitialize()
        return out_pdf if os.path.exists(out_pdf) else None
    except Exception as e:
        try:
            xl.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return None


# ──────────────────────────────────────────────
# PDF 추가
# ──────────────────────────────────────────────
def add_pdf(writer, path):
    start = len(writer.pages)
    try:
        for page in PdfReader(path).pages:
            writer.add_page(page)
    except Exception:
        pass
    return start


# ──────────────────────────────────────────────
# 폴더 재귀 순회
# ──────────────────────────────────────────────
def collect(folder, index_stack, tmp_dir, writer, stats, parent_bookmark, log):
    seen_main = {}
    idx = 0
    for entry in sorted_entries(folder):
        name = entry.name
        if name.startswith('.') or name.startswith('~$'):
            continue
        label   = extract_label(name)
        ext     = entry.suffix.lower()
        indent  = "  " * len(index_stack)
        is_file = ext in ('.pdf', '.xlsx', '.xls', '.xlsm')

        if entry.is_dir():
            idx += 1
            stack = index_stack + [idx]
            title = make_title(stack, label)
            log(f"{indent}📁 {title}")
            page_num = add_pdf(writer, make_cover(title, tmp_dir))
            bm = writer.add_outline_item(title, page_number=page_num, parent=parent_bookmark)
            stats['covers'] += 1
            collect(entry, stack, tmp_dir, writer, stats, bm, log)

        elif is_file:
            main_num, sub_nums = parse_file_number(name)
            if sub_nums:
                if main_num not in seen_main:
                    idx += 1
                    seen_main[main_num] = idx
                stack = index_stack + [seen_main[main_num]] + sub_nums
            else:
                idx += 1
                stack = index_stack + [idx]
            title = make_title(stack, label)

            if ext == '.pdf':
                log(f"{indent}📄 {title}")
                page_num = add_pdf(writer, make_cover(title, tmp_dir))
                writer.add_outline_item(title, page_number=page_num, parent=parent_bookmark)
                stats['covers'] += 1
                add_pdf(writer, str(entry))
                stats['pdfs'] += 1
            else:
                log(f"{indent}📊 {title}")
                converted = excel_to_pdf(str(entry), tmp_dir)
                if converted:
                    page_num = add_pdf(writer, make_cover(title, tmp_dir))
                    writer.add_outline_item(title, page_number=page_num, parent=parent_bookmark)
                    stats['covers'] += 1
                    add_pdf(writer, converted)
                    stats['excels'] += 1
                else:
                    log(f"  [변환 실패] {name}")
                    stats['failed'] += 1


# ──────────────────────────────────────────────
# 핵심 병합 로직
# ──────────────────────────────────────────────
def run_merge(root_folder, output_pdf, log):
    root = Path(root_folder)
    writer = PdfWriter()
    stats  = {'covers': 0, 'pdfs': 0, 'excels': 0, 'failed': 0}

    with tempfile.TemporaryDirectory() as tmp:
        log("📋 갑지 #1 : # 첨부자료")
        p1 = add_pdf(writer, make_cover("# 첨부자료", tmp))
        writer.add_outline_item("# 첨부자료", page_number=p1)
        stats['covers'] += 1

        log("📋 갑지 #2 : 추가공사비 집계표")
        p2 = add_pdf(writer, make_cover("추가공사비 집계표", tmp))
        writer.add_outline_item("추가공사비 집계표", page_number=p2)
        stats['covers'] += 1

        collect(root, [], tmp, writer, stats, None, log)

        if not writer.pages:
            log("처리된 파일이 없습니다.")
            return

        log(f"\n💾 저장 중... ({len(writer.pages)}페이지)")
        with open(output_pdf, "wb") as f:
            writer.write(f)

    log(f"\n✅ 완료!")
    log(f"   총 페이지  : {len(writer.pages)}")
    log(f"   갑지 수    : {stats['covers']}")
    log(f"   PDF        : {stats['pdfs']}개")
    log(f"   Excel      : {stats['excels']}개")
    if stats['failed']:
        log(f"   변환 실패  : {stats['failed']}개")
    log(f"\n   → {output_pdf}")


# ──────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 합치기 도구")
        self.geometry("700x520")
        self.resizable(True, True)
        self._build_ui()

    def _build_ui(self):
        pad = dict(padx=10, pady=5)

        # ── 입력 폴더 ──
        tk.Label(self, text="📂 루트 폴더", anchor="w").pack(fill="x", **pad)
        row1 = tk.Frame(self)
        row1.pack(fill="x", padx=10)
        self.folder_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.folder_var, font=("맑은 고딕", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="찾아보기", command=self._browse_folder).pack(side="left", padx=(5, 0))

        # ── 출력 파일 ──
        tk.Label(self, text="💾 출력 PDF 파일", anchor="w").pack(fill="x", **pad)
        row2 = tk.Frame(self)
        row2.pack(fill="x", padx=10)
        self.output_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.output_var, font=("맑은 고딕", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="저장 위치", command=self._browse_output).pack(side="left", padx=(5, 0))

        # ── 실행 버튼 ──
        self.run_btn = tk.Button(
            self, text="▶  합치기 시작", font=("맑은 고딕", 11, "bold"),
            bg="#0078D7", fg="white", relief="flat",
            command=self._start, pady=6
        )
        self.run_btn.pack(fill="x", padx=10, pady=(10, 5))

        # ── 로그창 ──
        tk.Label(self, text="진행 로그", anchor="w").pack(fill="x", padx=10)
        self.log_box = scrolledtext.ScrolledText(
            self, font=("맑은 고딕", 9), state="disabled",
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white"
        )
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _browse_folder(self):
        path = filedialog.askdirectory(title="루트 폴더 선택")
        if path:
            self.folder_var.set(path)
            # 출력 경로 자동 설정
            if not self.output_var.get():
                p = Path(path)
                self.output_var.set(str(p.parent / (p.name + "_merged.pdf")))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="출력 PDF 저장 위치",
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
        )
        if path:
            self.output_var.set(path)

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _start(self):
        folder = self.folder_var.get().strip()
        output = self.output_var.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("오류", "유효한 루트 폴더를 선택해 주세요.")
            return
        if not output:
            messagebox.showerror("오류", "출력 PDF 파일 경로를 입력해 주세요.")
            return

        self.run_btn.config(state="disabled", text="⏳  처리 중...")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        def worker():
            try:
                self._log(f"루트 폴더 : {folder}")
                self._log(f"출력 파일 : {output}")
                self._log("=" * 45)
                run_merge(folder, output, self._log)
            except Exception as e:
                self._log(f"\n[오류] {e}")
            finally:
                self.run_btn.config(state="normal", text="▶  합치기 시작")

        threading.Thread(target=worker, daemon=True).start()


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
if __name__ == "__main__":
    # CLI 모드 (인수가 있을 때)
    if len(sys.argv) >= 2:
        root   = sys.argv[1]
        output = sys.argv[2] if len(sys.argv) >= 3 else None
        if output is None:
            p = Path(root)
            output = str(p.parent / (p.name + "_merged.pdf"))
        print(f"루트 폴더 : {root}")
        print(f"출력 파일 : {output}")
        run_merge(root, output, print)
    else:
        # GUI 모드
        App().mainloop()

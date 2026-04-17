"""
PDF 자동 합치기 + 갑지 생성 도구 v4
- v3 대비 변경: 폴더에만 간지 삽입, 파일(PDF/Excel)은 간지 없이 바로 추가
- 초기 갑지를 사용자가 자유롭게 설정 (추가/삭제/순서변경)
- 갑지 번호 시작 번호 지정 가능
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
def remove_ext(name: str) -> str:
    return re.sub(r'\.(pdf|xlsx|xls|xlsm)$', '', name, flags=re.IGNORECASE).strip()


def extract_label(name: str) -> str:
    base   = remove_ext(name)
    result = re.sub(r'^\d+(?:\.\d+)*\.?\s*', '', base).strip()
    return result if result else base


def sort_key(entry: Path):
    m = re.match(r'^(\d+(?:\.\d+)*)[.\s]', entry.name)
    if m:
        return tuple(int(x) for x in m.group(1).split('.'))
    return (9999,)


def sorted_entries(folder: Path):
    return sorted(folder.iterdir(), key=sort_key)


def make_title(index_stack: list, label: str) -> str:
    return '.'.join(str(i) for i in index_stack) + '. ' + label


# ──────────────────────────────────────────────
# 갑지 생성
# ──────────────────────────────────────────────
def make_cover(title: str, tmp_dir: str) -> str:
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
def excel_to_pdf(excel_path: str, tmp_dir: str):
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
    except Exception:
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
def add_pdf(writer: PdfWriter, path: str) -> int:
    start = len(writer.pages)
    try:
        for page in PdfReader(path).pages:
            writer.add_page(page)
    except Exception:
        pass
    return start


# ──────────────────────────────────────────────
# 번호 유무 판단
# ──────────────────────────────────────────────
def has_leading_number(name: str) -> bool:
    return bool(re.match(r'^\d+[.\s]', name))


# ──────────────────────────────────────────────
# 폴더 재귀 순회
# ──────────────────────────────────────────────
def collect(folder: Path, index_stack: list, tmp_dir: str,
            writer: PdfWriter, stats: dict, parent_bookmark, log,
            start_number: int = 1):
    """
    v4 변경점: 폴더에만 간지 삽입, 파일(PDF/Excel)은 간지 없이 바로 추가.
    파일 북마크는 파일 첫 페이지를 직접 가리킨다.

    번호 있는 항목 → 계층 번호 생성 (4.1.1. 당초 도면)
    번호 없는 항목 → 번호 없이 이름만 (당초 도면)
    start_number: 최상위 레벨에서만 적용되는 시작 번호 오프셋.
    """
    idx = (start_number - 1) if len(index_stack) == 0 else 0

    for entry in sorted_entries(folder):
        name = entry.name

        if name.startswith('.') or name.startswith('~$'):
            continue

        ext     = entry.suffix.lower()
        is_pdf  = ext == '.pdf'
        is_xlsx = ext in ('.xlsx', '.xls', '.xlsm')
        indent  = "  " * len(index_stack)

        if not entry.is_dir() and not is_pdf and not is_xlsx:
            continue

        label    = extract_label(name)
        numbered = has_leading_number(name)

        if numbered:
            idx += 1
            stack = index_stack + [idx]
            title = make_title(stack, label)
        else:
            stack = index_stack
            title = label

        # ── 폴더: 간지 삽입 ──
        if entry.is_dir():
            log(f"{indent}📁 {title}")
            page_num = add_pdf(writer, make_cover(title, tmp_dir))
            bm = writer.add_outline_item(title, page_number=page_num, parent=parent_bookmark)
            stats['covers'] += 1
            collect(entry, stack, tmp_dir, writer, stats, bm, log, start_number=1)

        # ── PDF: 간지 없이 바로 추가 ──
        elif is_pdf:
            log(f"{indent}📄 {title}")
            page_num = add_pdf(writer, str(entry))
            writer.add_outline_item(title, page_number=page_num, parent=parent_bookmark)
            stats['pdfs'] += 1

        # ── Excel: 간지 없이 변환본 바로 추가 ──
        elif is_xlsx:
            log(f"{indent}📊 {title}")
            converted = excel_to_pdf(str(entry), tmp_dir)
            if converted:
                page_num = add_pdf(writer, converted)
                writer.add_outline_item(title, page_number=page_num, parent=parent_bookmark)
                stats['excels'] += 1
            else:
                log(f"{indent}  [변환 실패] {name}")
                stats['failed'] += 1


# ──────────────────────────────────────────────
# 핵심 병합 로직
# ──────────────────────────────────────────────
def run_merge(root_folder: str, output_pdf: str, log,
              initial_covers: list = None, start_number: int = 1):
    """
    initial_covers: 초기 갑지 이름 리스트 (예: ["# 첨부자료", "추가공사비 집계표"])
                    None이면 초기 갑지 없이 바로 폴더 순회
    start_number:   폴더 내 번호 매기기 시작 번호 (기본 1)
    """
    if initial_covers is None:
        initial_covers = []

    root   = Path(root_folder)
    writer = PdfWriter()
    stats  = {'covers': 0, 'pdfs': 0, 'excels': 0, 'failed': 0}

    with tempfile.TemporaryDirectory() as tmp:
        # 초기 갑지 생성
        for i, cover_name in enumerate(initial_covers, 1):
            cover_name = cover_name.strip()
            if not cover_name:
                continue
            log(f"📋 초기 갑지 #{i} : {cover_name}")
            p = add_pdf(writer, make_cover(cover_name, tmp))
            writer.add_outline_item(cover_name, page_number=p)
            stats['covers'] += 1

        collect(root, [], tmp, writer, stats, None, log, start_number=start_number)

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
        self.title("PDF 합치기 도구 v4")
        self.geometry("720x700")
        self.resizable(True, True)
        self.cover_entries = []
        self._build_ui()
        self._add_cover_entry("# 첨부자료")
        self._add_cover_entry("추가공사비 집계표")

    def _build_ui(self):
        pad = dict(padx=10, pady=5)

        # ── 루트 폴더 ──
        tk.Label(self, text="📂 루트 폴더", anchor="w").pack(fill="x", **pad)
        row1 = tk.Frame(self)
        row1.pack(fill="x", padx=10)
        self.folder_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.folder_var, font=("맑은 고딕", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="찾아보기", command=self._browse_folder).pack(side="left", padx=(5, 0))

        # ── 출력 PDF ──
        tk.Label(self, text="💾 출력 PDF 파일", anchor="w").pack(fill="x", **pad)
        row2 = tk.Frame(self)
        row2.pack(fill="x", padx=10)
        self.output_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.output_var, font=("맑은 고딕", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="저장 위치", command=self._browse_output).pack(side="left", padx=(5, 0))

        # ── 시작 번호 ──
        tk.Label(self, text="🔢 시작 번호 (폴더 내 첫 번째 항목의 번호)", anchor="w").pack(fill="x", **pad)
        row_num = tk.Frame(self)
        row_num.pack(fill="x", padx=10)
        self.start_num_var = tk.StringVar(value="1")
        tk.Spinbox(row_num, from_=1, to=9999, textvariable=self.start_num_var,
                    font=("맑은 고딕", 10), width=8).pack(side="left")
        tk.Label(row_num, text="  (예: 3 입력 시 → 3. xxx, 4. xxx, 5. xxx ...)",
                 fg="#666666").pack(side="left", padx=(5, 0))

        # ── 초기 갑지 설정 ──
        sep = tk.Frame(self, height=1, bg="#cccccc")
        sep.pack(fill="x", padx=10, pady=(10, 5))

        header_row = tk.Frame(self)
        header_row.pack(fill="x", padx=10, pady=(0, 2))
        tk.Label(header_row, text="📋 초기 갑지 설정 (합치기 전에 맨 앞에 삽입할 갑지)",
                 anchor="w").pack(side="left")
        tk.Button(header_row, text="＋ 갑지 추가", command=lambda: self._add_cover_entry(""),
                  bg="#4CAF50", fg="white", relief="flat", padx=8).pack(side="right")

        self.cover_canvas = tk.Canvas(self, height=120, highlightthickness=0)
        self.cover_scrollbar = tk.Scrollbar(self, orient="vertical", command=self.cover_canvas.yview)
        self.cover_frame = tk.Frame(self.cover_canvas)

        self.cover_frame.bind("<Configure>",
            lambda e: self.cover_canvas.configure(scrollregion=self.cover_canvas.bbox("all")))
        self.cover_canvas.create_window((0, 0), window=self.cover_frame, anchor="nw")
        self.cover_canvas.configure(yscrollcommand=self.cover_scrollbar.set)

        self.cover_canvas.pack(fill="x", padx=10, pady=(0, 5))
        self.cover_scrollbar.pack_forget()

        # ── 실행 버튼 ──
        self.run_btn = tk.Button(
            self, text="▶  합치기 시작", font=("맑은 고딕", 11, "bold"),
            bg="#0078D7", fg="white", relief="flat",
            command=self._start, pady=6
        )
        self.run_btn.pack(fill="x", padx=10, pady=(10, 5))

        # ── 로그 ──
        tk.Label(self, text="진행 로그", anchor="w").pack(fill="x", padx=10)
        self.log_box = scrolledtext.ScrolledText(
            self, font=("맑은 고딕", 9), state="disabled",
            bg="#1e1e1e", fg="#d4d4d4"
        )
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _add_cover_entry(self, default_text: str = ""):
        row = tk.Frame(self.cover_frame)
        row.pack(fill="x", pady=1)

        idx_label = tk.Label(row, text=f"#{len(self.cover_entries) + 1}", width=3,
                             font=("맑은 고딕", 9), fg="#888888")
        idx_label.pack(side="left")

        var = tk.StringVar(value=default_text)
        entry = tk.Entry(row, textvariable=var, font=("맑은 고딕", 10))
        entry.pack(side="left", fill="x", expand=True, padx=(2, 5))

        btn_up = tk.Button(row, text="▲", width=2,
                           command=lambda r=row: self._move_cover(r, -1))
        btn_up.pack(side="left", padx=1)

        btn_down = tk.Button(row, text="▼", width=2,
                             command=lambda r=row: self._move_cover(r, 1))
        btn_down.pack(side="left", padx=1)

        btn_del = tk.Button(row, text="✕", width=2, fg="red",
                            command=lambda r=row: self._remove_cover(r))
        btn_del.pack(side="left", padx=(1, 0))

        self.cover_entries.append({'row': row, 'var': var, 'label': idx_label})
        self._update_cover_numbers()
        self.after(50, self._adjust_cover_canvas)

    def _remove_cover(self, row_frame):
        for i, item in enumerate(self.cover_entries):
            if item['row'] == row_frame:
                row_frame.destroy()
                self.cover_entries.pop(i)
                break
        self._update_cover_numbers()
        self.after(50, self._adjust_cover_canvas)

    def _move_cover(self, row_frame, direction):
        idx = None
        for i, item in enumerate(self.cover_entries):
            if item['row'] == row_frame:
                idx = i
                break
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.cover_entries):
            return

        self.cover_entries[idx], self.cover_entries[new_idx] = \
            self.cover_entries[new_idx], self.cover_entries[idx]

        for item in self.cover_entries:
            item['row'].pack_forget()
        for item in self.cover_entries:
            item['row'].pack(fill="x", pady=1)

        self._update_cover_numbers()

    def _update_cover_numbers(self):
        for i, item in enumerate(self.cover_entries):
            item['label'].config(text=f"#{i + 1}")

    def _adjust_cover_canvas(self):
        self.cover_frame.update_idletasks()
        needed = self.cover_frame.winfo_reqheight()
        max_h = 150
        self.cover_canvas.config(height=min(needed + 5, max_h))
        if needed > max_h:
            self.cover_scrollbar.pack(side="right", fill="y")
        else:
            self.cover_scrollbar.pack_forget()

    def _get_cover_list(self) -> list:
        return [item['var'].get() for item in self.cover_entries if item['var'].get().strip()]

    def _browse_folder(self):
        path = filedialog.askdirectory(title="루트 폴더 선택")
        if path:
            self.folder_var.set(path)
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

        try:
            start_num = int(self.start_num_var.get())
            if start_num < 1:
                start_num = 1
        except ValueError:
            start_num = 1

        covers = self._get_cover_list()

        self.run_btn.config(state="disabled", text="⏳  처리 중...")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        def worker():
            try:
                self._log(f"루트 폴더 : {folder}")
                self._log(f"출력 파일 : {output}")
                self._log(f"시작 번호 : {start_num}")
                if covers:
                    self._log(f"초기 갑지 : {', '.join(covers)}")
                else:
                    self._log(f"초기 갑지 : (없음)")
                self._log("=" * 45)
                run_merge(folder, output, self._log,
                          initial_covers=covers, start_number=start_num)
            except Exception as e:
                self._log(f"\n[오류] {e}")
            finally:
                self.run_btn.config(state="normal", text="▶  합치기 시작")

        threading.Thread(target=worker, daemon=True).start()


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) >= 2:
        import argparse
        parser = argparse.ArgumentParser(description="PDF 합치기 도구 v4")
        parser.add_argument("root", help="루트 폴더 경로")
        parser.add_argument("output", nargs="?", default=None, help="출력 PDF 경로")
        parser.add_argument("--covers", type=str, default="",
                            help="초기 갑지 이름 (쉼표 구분)")
        parser.add_argument("--start", type=int, default=1,
                            help="번호 시작 번호 (기본 1)")
        args = parser.parse_args()

        output = args.output
        if output is None:
            p = Path(args.root)
            output = str(p.parent / (p.name + "_merged.pdf"))

        covers = [c.strip() for c in args.covers.split(",") if c.strip()] if args.covers else []

        print(f"루트 폴더 : {args.root}")
        print(f"출력 파일 : {output}")
        print(f"시작 번호 : {args.start}")
        if covers:
            print(f"초기 갑지 : {', '.join(covers)}")
        run_merge(args.root, output, print,
                  initial_covers=covers, start_number=args.start)
    else:
        App().mainloop()

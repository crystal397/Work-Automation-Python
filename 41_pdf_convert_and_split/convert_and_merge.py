"""
폴더 PDF 변환 + 병합 도구 (GUI/CLI 통합).

폴더 내 모든 파일을 PDF로 변환한 뒤, 하나의 PDF로 병합합니다.
- 병합 순서: 폴더구조 순서 (rglob 기준, 하위 폴더 깊이 우선)
- 출력 파일명: <폴더명>_merged.pdf
- 병합 성공 시 개별 PDF는 삭제됨

사용법:
  - 인수 없이 실행 → GUI 창 (드래그&드롭 지원)
  - 폴더 경로 인수 → GUI에 경로 자동 채움
  - --cli 플래그 → 순수 CLI 모드
      convert_and_merge.exe --cli "C:\\folder"
"""

import os
import sys
import time
import threading
from pathlib import Path
from typing import Callable, Optional

WORD_EXTS = {".doc", ".docx", ".rtf"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".csv"}
PPT_EXTS = {".ppt", ".pptx"}
HWP_EXTS = {".hwp", ".hwpx"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif"}
TEXT_EXTS = {".txt", ".md", ".log"}
PDF_EXT = ".pdf"

LogFn = Callable[[str], None]


# ──────────────────────────────────────────────────────────────
# 변환기들 (convert_and_split.py와 동일)
# ──────────────────────────────────────────────────────────────

def convert_word_to_pdf(src: Path, dst: Path, log: LogFn) -> bool:
    import win32com.client
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=True)
        doc.SaveAs(str(dst), FileFormat=17)
        return True
    except Exception as e:
        log(f"  [Word 변환 실패] {src.name}: {e}")
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def convert_excel_to_pdf(src: Path, dst: Path, log: LogFn) -> bool:
    import win32com.client
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(src), ReadOnly=True)
        wb.ExportAsFixedFormat(0, str(dst))
        return True
    except Exception as e:
        log(f"  [Excel 변환 실패] {src.name}: {e}")
        return False
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def convert_ppt_to_pdf(src: Path, dst: Path, log: LogFn) -> bool:
    import win32com.client
    ppt = None
    pres = None
    try:
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        pres = ppt.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
        pres.SaveAs(str(dst), 32)
        return True
    except Exception as e:
        log(f"  [PowerPoint 변환 실패] {src.name}: {e}")
        return False
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:
            pass


def convert_hwp_to_pdf(src: Path, dst: Path, log: LogFn) -> bool:
    """한글(.hwp/.hwpx) → PDF.
    
    전략:
      1) simple-hwp2pdf 라이브러리 우선 (보안 다이얼로그 없음, .hwpx에 강함)
      2) 실패 시 한컴 COM 직접 호출 (재시도 3회)
    """
    # ─── 1차: simple-hwp2pdf 사용 ───
    try:
        from simple_hwp2pdf import convert as shc_convert
        try:
            shc_convert(str(src), str(dst), method="auto")
            if dst.exists() and dst.stat().st_size > 0:
                return True
        except Exception as e:
            log(f"     (simple-hwp2pdf 실패: {e} — 한컴 COM으로 폴백)")
    except ImportError:
        pass

    # ─── 2차: 한컴 COM 직접 호출 ───
    import win32com.client
    last_err = None
    for attempt in range(1, 4):
        hwp = None
        try:
            hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
            try:
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass
            hwp.Open(str(src), "", "forceopen:true")
            try:
                pset = hwp.HParameterSet.HFileOpenSave
                hwp.HAction.GetDefault("FileSaveAsPdf", pset.HSet)
                pset.filename = str(dst)
                pset.Format = "PDF"
                ok = hwp.HAction.Execute("FileSaveAsPdf", pset.HSet)
                if ok and dst.exists():
                    return True
            except Exception as e1:
                last_err = e1
            try:
                hwp.SaveAs(str(dst), "PDF", "")
                if dst.exists():
                    return True
            except Exception as e2:
                last_err = e2
        except Exception as e:
            last_err = e
        finally:
            try:
                if hwp is not None:
                    hwp.XHwpDocuments.Item(0).Close(isDirty=False)
            except Exception:
                pass
            try:
                if hwp is not None:
                    hwp.Quit()
            except Exception:
                pass
            hwp = None
            time.sleep(1.0)

        if attempt < 3:
            log(f"     (COM 시도 {attempt} 실패: {last_err} — 재시도)")
            time.sleep(1.5)

    log(f"  [한글 변환 실패] {src.name}: {last_err}")
    return False


def convert_image_to_pdf(src: Path, dst: Path, log: LogFn) -> bool:
    try:
        from PIL import Image
        img = Image.open(src)
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        img.save(str(dst), "PDF", resolution=100.0)
        return True
    except Exception as e:
        log(f"  [이미지 변환 실패] {src.name}: {e}")
        return False


def convert_text_to_pdf(src: Path, dst: Path, log: LogFn) -> bool:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        font_name = "Helvetica"
        for font_path in [
            r"C:\Windows\Fonts\malgun.ttf",
            r"C:\Windows\Fonts\gulim.ttc",
            r"C:\Windows\Fonts\batang.ttc",
        ]:
            if os.path.exists(font_path):
                try:
                    pdfmetrics.registerFont(TTFont("KoreanFont", font_path))
                    font_name = "KoreanFont"
                    break
                except Exception:
                    continue

        text = None
        for enc in ("utf-8", "cp949", "euc-kr", "latin-1"):
            try:
                with open(src, "r", encoding=enc) as f:
                    text = f.read()
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            log(f"  [텍스트 변환 실패] {src.name}: 인코딩 감지 실패")
            return False

        c = canvas.Canvas(str(dst), pagesize=A4)
        width, height = A4
        c.setFont(font_name, 10)
        margin = 40
        y = height - margin
        line_height = 14

        for line in text.splitlines():
            while len(line) > 100:
                c.drawString(margin, y, line[:100])
                line = line[100:]
                y -= line_height
                if y < margin:
                    c.showPage()
                    c.setFont(font_name, 10)
                    y = height - margin
            c.drawString(margin, y, line)
            y -= line_height
            if y < margin:
                c.showPage()
                c.setFont(font_name, 10)
                y = height - margin

        c.save()
        return True
    except Exception as e:
        log(f"  [텍스트 변환 실패] {src.name}: {e}")
        return False


def convert_to_pdf(src: Path, log: LogFn) -> Optional[Path]:
    """파일을 PDF로 변환. 같은 이름의 PDF가 이미 있으면 건너뜀."""
    ext = src.suffix.lower()
    dst = src.with_suffix(".pdf")
    if dst.exists():
        log(f"  → 건너뜀 (이미 존재): {dst.name}")
        return dst

    log(f"  → PDF 변환 중: {src.name}{dst.name}")

    if ext in WORD_EXTS:
        ok = convert_word_to_pdf(src, dst, log)
    elif ext in EXCEL_EXTS:
        ok = convert_excel_to_pdf(src, dst, log)
    elif ext in PPT_EXTS:
        ok = convert_ppt_to_pdf(src, dst, log)
    elif ext in HWP_EXTS:
        ok = convert_hwp_to_pdf(src, dst, log)
    elif ext in IMAGE_EXTS:
        ok = convert_image_to_pdf(src, dst, log)
    elif ext in TEXT_EXTS:
        ok = convert_text_to_pdf(src, dst, log)
    else:
        return None

    if not ok and dst.exists():
        try:
            dst.unlink()
        except Exception:
            pass

    return dst if ok and dst.exists() else None


# ──────────────────────────────────────────────────────────────
# 정렬 키: 폴더구조 순서대로
# ──────────────────────────────────────────────────────────────

def folder_structure_sort_key(path: Path, root: Path):
    """폴더 구조 순서대로 정렬하기 위한 키.
    
    - 루트에 가까운 폴더가 먼저
    - 같은 폴더 내에서는 파일명 순
    - 폴더 → 그 하위 폴더 → ... 순서로 깊이 우선 탐색
    """
    try:
        rel = path.relative_to(root)
    except ValueError:
        rel = path
    parts = rel.parts
    # 각 part를 소문자 + 자연 정렬 가능한 형태로
    # 디렉터리 부분과 파일명 부분을 분리
    return tuple(p.lower() for p in parts)


# ──────────────────────────────────────────────────────────────
# PDF 병합
# ──────────────────────────────────────────────────────────────

def merge_pdfs(pdf_files: list, output_path: Path, log: LogFn) -> bool:
    """여러 PDF를 하나로 병합."""
    from pypdf import PdfWriter, PdfReader

    writer = PdfWriter()
    merged_count = 0
    failed = []

    for pdf in pdf_files:
        try:
            reader = PdfReader(str(pdf))
            for page in reader.pages:
                writer.add_page(page)
            merged_count += 1
            log(f"     + {pdf.name} ({len(reader.pages)} 페이지)")
        except Exception as e:
            log(f"     [건너뜀] {pdf.name}: {e}")
            failed.append(pdf)

    if merged_count == 0:
        log("     [병합 실패] 병합할 PDF가 없습니다.")
        return False

    try:
        with open(output_path, "wb") as f:
            writer.write(f)
        log(f"\n  ✓ 병합 완료: {output_path.name} ({merged_count}개 PDF)")
        if failed:
            log(f"  ⚠ 병합 실패 {len(failed)}개: 개별 PDF로 남겨둡니다.")
        return True
    except Exception as e:
        log(f"     [병합 실패] 저장 오류: {e}")
        return False


# ──────────────────────────────────────────────────────────────
# 메인 처리
# ──────────────────────────────────────────────────────────────

def process_folder(folder: Path, recursive: bool = True,
                   log: Optional[LogFn] = None):
    if log is None:
        log = print

    if not folder.exists() or not folder.is_dir():
        log(f"오류: 유효한 폴더가 아닙니다: {folder}")
        return

    log(f"\n작업 폴더: {folder}")
    log(f"하위 폴더 포함: {'예' if recursive else '아니오'}")
    log("=" * 60)

    # ─── 1단계: PDF 변환 ───
    log("\n[1단계] 비-PDF 파일을 PDF로 변환")
    log("-" * 60)

    iterator = folder.rglob("*") if recursive else folder.glob("*")
    files_to_convert = sorted(
        [p for p in iterator if p.is_file() and p.suffix.lower() != PDF_EXT],
        key=lambda p: folder_structure_sort_key(p, folder),
    )

    converted = 0
    already_exists = 0
    unsupported = 0
    failed = 0
    for src in files_to_convert:
        ext = src.suffix.lower()
        supported = (
            ext in WORD_EXTS or ext in EXCEL_EXTS or ext in PPT_EXTS
            or ext in HWP_EXTS or ext in IMAGE_EXTS or ext in TEXT_EXTS
        )
        if not supported:
            unsupported += 1
            continue
        try:
            rel = src.relative_to(folder)
        except ValueError:
            rel = src
        log(f"\n파일: {rel}")

        will_skip = src.with_suffix(".pdf").exists()
        result = convert_to_pdf(src, log)
        if result is None:
            failed += 1
        elif will_skip:
            already_exists += 1
        else:
            converted += 1
    log(f"\n  변환됨: {converted}개 / 이미 존재(skip): {already_exists}개 "
        f"/ 미지원(skip): {unsupported}개 / 실패: {failed}개")

    # ─── 2단계: PDF 수집 및 정렬 ───
    log("\n" + "=" * 60)
    log("\n[2단계] PDF 파일들을 폴더 구조 순으로 정렬")
    log("-" * 60)

    iterator = folder.rglob("*.pdf") if recursive else folder.glob("*.pdf")
    output_name = f"{folder.name}_merged.pdf"
    output_path = folder / output_name

    pdf_files = sorted(
        [p for p in iterator if p.is_file() and p.name != output_name],
        key=lambda p: folder_structure_sort_key(p, folder),
    )

    if not pdf_files:
        log("  병합할 PDF가 없습니다.")
        return

    log(f"  병합 대상 PDF: {len(pdf_files)}개")
    for i, p in enumerate(pdf_files, 1):
        try:
            rel = p.relative_to(folder)
        except ValueError:
            rel = p
        log(f"    {i:>3}. {rel}")

    # ─── 3단계: 병합 ───
    log("\n" + "=" * 60)
    log(f"\n[3단계] 하나의 PDF로 병합 → {output_name}")
    log("-" * 60)

    # 출력 파일이 이미 있으면 백업 (덮어쓰기 방지)
    if output_path.exists():
        idx = 1
        while True:
            backup = folder / f"{folder.name}_merged_{idx}.pdf"
            if not backup.exists():
                output_path = backup
                break
            idx += 1
        log(f"  기존 병합 파일이 있어 새 이름으로 저장: {output_path.name}")

    success = merge_pdfs(pdf_files, output_path, log)

    # ─── 4단계: 개별 PDF 삭제 ───
    if success:
        log("\n" + "=" * 60)
        log("\n[4단계] 개별 PDF 파일 삭제")
        log("-" * 60)
        deleted = 0
        for pdf in pdf_files:
            try:
                pdf.unlink()
                deleted += 1
            except Exception as e:
                log(f"     [삭제 실패] {pdf.name}: {e}")
        log(f"  삭제 완료: {deleted}개")

    log("\n" + "=" * 60)
    log("모든 작업이 완료되었습니다.")


# ──────────────────────────────────────────────────────────────
# GUI (convert_and_split.py와 거의 동일)
# ──────────────────────────────────────────────────────────────

def run_gui(initial_folder: Optional[str] = None):
    import tkinter as tk
    from tkinter import ttk, filedialog, scrolledtext, messagebox

    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        dnd_available = True
    except ImportError:
        root = tk.Tk()
        dnd_available = False

    root.title("폴더 PDF 변환 + 병합 도구")
    root.geometry("780x560")
    root.minsize(640, 480)

    top_frame = ttk.Frame(root, padding=10)
    top_frame.pack(fill="x")

    ttk.Label(
        top_frame,
        text="폴더 내 모든 파일을 PDF로 변환 후 하나의 PDF로 병합합니다.",
        font=("맑은 고딕", 11, "bold"),
    ).pack(anchor="w")

    ttk.Label(
        top_frame,
        text="결과물: <폴더명>_merged.pdf · 개별 PDF는 병합 후 삭제됩니다",
        foreground="#555",
    ).pack(anchor="w", pady=(2, 0))

    folder_frame = ttk.Frame(root, padding=(10, 0))
    folder_frame.pack(fill="x")

    ttk.Label(folder_frame, text="폴더 경로:").pack(side="left")
    folder_var = tk.StringVar()
    folder_entry = ttk.Entry(folder_frame, textvariable=folder_var)
    folder_entry.pack(side="left", fill="x", expand=True, padx=6)

    def browse():
        path = filedialog.askdirectory(title="처리할 폴더 선택")
        if path:
            folder_var.set(path)

    ttk.Button(folder_frame, text="찾아보기...", command=browse).pack(side="left")

    opt_frame = ttk.Frame(root, padding=(10, 6))
    opt_frame.pack(fill="x")

    recursive_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        opt_frame, text="하위 폴더 포함", variable=recursive_var
    ).pack(side="left")

    dnd_status = "✔ 드래그&드롭 가능" if dnd_available else "✖ 드래그&드롭 비활성"
    ttk.Label(opt_frame, text=dnd_status, foreground="#888").pack(side="right")

    progress = ttk.Progressbar(root, mode="indeterminate")
    progress.pack(fill="x", padx=10, pady=(4, 0))

    log_frame = ttk.Frame(root, padding=10)
    log_frame.pack(fill="both", expand=True)

    log_text = scrolledtext.ScrolledText(
        log_frame, wrap="word", font=("Consolas", 9)
    )
    log_text.pack(fill="both", expand=True)

    btn_frame = ttk.Frame(root, padding=(10, 0, 10, 10))
    btn_frame.pack(fill="x")

    is_running = {"value": False}

    def append_log(msg: str):
        def _do():
            log_text.insert("end", msg + "\n")
            log_text.see("end")
        root.after(0, _do)

    def start():
        if is_running["value"]:
            return
        folder = folder_var.get().strip().strip('"')
        if not folder:
            messagebox.showwarning("경고", "폴더를 선택해 주세요.")
            return
        p = Path(folder)
        if not p.exists() or not p.is_dir():
            messagebox.showerror("오류", f"유효한 폴더가 아닙니다:\n{folder}")
            return

        # 사용자에게 동작 확인
        confirm = messagebox.askyesno(
            "확인",
            f"'{p.name}' 폴더의 모든 파일을 PDF로 변환한 뒤,\n"
            f"하나의 PDF({p.name}_merged.pdf)로 병합합니다.\n\n"
            f"병합이 성공하면 개별 PDF는 삭제됩니다.\n\n"
            f"계속할까요?"
        )
        if not confirm:
            return

        is_running["value"] = True
        log_text.delete("1.0", "end")
        progress.start(10)
        run_btn.config(state="disabled")

        def worker():
            try:
                process_folder(p, recursive=recursive_var.get(), log=append_log)
            except Exception as e:
                append_log(f"\n[치명적 오류] {e}")
            finally:
                root.after(0, _on_done)

        def _on_done():
            progress.stop()
            run_btn.config(state="normal")
            is_running["value"] = False

        threading.Thread(target=worker, daemon=True).start()

    run_btn = ttk.Button(btn_frame, text="실행", command=start)
    run_btn.pack(side="right")

    if dnd_available:
        def on_drop(event):
            data = event.data.strip()
            if data.startswith("{"):
                end = data.find("}")
                path = data[1:end] if end != -1 else data
            else:
                path = data.split()[0] if " " not in data else data
            folder_var.set(path)

        root.drop_target_register(DND_FILES)
        root.dnd_bind("<<Drop>>", on_drop)

    if initial_folder:
        folder_var.set(initial_folder)

    root.mainloop()


# ──────────────────────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────────────────────

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = [a for a in sys.argv[1:] if a.startswith("--")]

    if "--cli" in flags:
        if not args:
            print("사용법: convert_and_merge.exe --cli <폴더경로>")
            sys.exit(1)
        folder = Path(args[0]).resolve()
        recursive = "--no-recursive" not in flags
        process_folder(folder, recursive=recursive, log=print)
        return

    initial = args[0] if args else None
    run_gui(initial_folder=initial)


if __name__ == "__main__":
    main()

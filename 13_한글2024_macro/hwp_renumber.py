"""
============================================================
한글 2024 표/그림 번호 재정렬 v5
============================================================

[방식]
  HWPX 파일(ZIP+XML)을 직접 수정 — COM Find/Replace 우회
  → 표 셀 안, 글상자 안 등 어디에 있어도 동작

[사용법]
  1. python hwp_renumber.py (또는 exe 실행)
  2. 파일 선택 → 실행
  3. _renumbered.hwpx 파일이 생성됨
============================================================
"""

import win32com.client
import zipfile
import re
import os
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox


# ==============================================================
# 설정값
# ==============================================================
TABLE_START  = 1
FIGURE_START = 1


# ==============================================================
# 한글 연결 (선택 사항 — 열려 있으면 자동 저장)
# ==============================================================
def try_connect_hwp():
    """한글이 실행 중이면 연결해서 저장, 아니면 None 반환."""
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        except Exception:
            pass
        return hwp
    except Exception:
        return None


# ==============================================================
# HWPX XML 직접 수정
# ==============================================================
def count_autonums(src_path):
    """
    HWPX 내 HWP 자동번호 필드(autoNum) 개수 반환.
    자동번호 필드는 정규식으로 처리 불가 → 그 수만큼 시작 번호를 뒤로 밀어야 중복 방지.
    """
    auto_table = 0
    auto_figure = 0
    with zipfile.ZipFile(src_path, 'r') as z:
        for name in z.namelist():
            if not name.endswith('.xml'):
                continue
            text = z.read(name).decode('utf-8', errors='ignore')
            auto_table  += len(re.findall(r'<hp:autoNum\b[^>]*numType="TABLE"', text))
            auto_figure += len(re.findall(r'<hp:autoNum\b[^>]*numType="FIGURE"', text))
    return auto_table, auto_figure


def renumber_hwpx(src_path, dst_path, log):
    """
    HWPX(ZIP) 내 XML 파일을 열어 표/그림 번호를 순차 재정렬.
    log: 메시지 출력 콜백 함수
    반환: (table_count, figure_count)
    """
    auto_t, auto_f = count_autonums(src_path)
    tstart = TABLE_START + auto_t
    fstart = FIGURE_START + auto_f
    if auto_t:
        log(f"  [정보] 표 자동번호 필드 {auto_t}개 감지 → 일반 텍스트 표 번호를 {tstart}부터 시작")
    if auto_f:
        log(f"  [정보] 그림 자동번호 필드 {auto_f}개 감지 → 일반 텍스트 그림 번호를 {fstart}부터 시작")

    state = {
        "tnum": tstart, "tcnt": 0,
        "fnum": fstart, "fcnt": 0,
    }

    pat_table   = re.compile(r'\[표 \d+\]')
    pat_fig_enc = re.compile(r'&lt;그림 \d+&gt;')
    pat_fig_raw = re.compile(r'<그림 \d+>')

    def repl_table(_m):
        r = f'[표 {state["tnum"]}]'
        state["tnum"] += 1; state["tcnt"] += 1
        return r

    def repl_fig_enc(_m):
        r = f'&lt;그림 {state["fnum"]}&gt;'
        state["fnum"] += 1; state["fcnt"] += 1
        return r

    def repl_fig_raw(_m):
        r = f'<그림 {state["fnum"]}>'
        state["fnum"] += 1; state["fcnt"] += 1
        return r

    with zipfile.ZipFile(src_path, 'r') as zin:
        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith('.xml'):
                    try:
                        text = data.decode('utf-8')
                        text = pat_table.sub(repl_table, text)
                        text = pat_fig_enc.sub(repl_fig_enc, text)
                        text = pat_fig_raw.sub(repl_fig_raw, text)
                        data = text.encode('utf-8')
                    except Exception:
                        pass
                zout.writestr(item, data)

    return state["tcnt"], state["fcnt"]


# ==============================================================
# GUI
# ==============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("한글 표/그림 번호 재정렬")
        self.resizable(False, False)
        self._build()
        self._try_autofill()

    # ── 화면 구성 ──────────────────────────────────────────────
    def _build(self):
        pad = {"padx": 10, "pady": 6}

        # 제목
        tk.Label(self, text="한글 2024  표/그림 번호 재정렬",
                 font=("맑은 고딕", 13, "bold")).grid(
            row=0, column=0, columnspan=3, pady=(14, 4))
        tk.Label(self, text="[표 N] / <그림 N> 번호를 문서 순서대로 1부터 재정렬합니다.",
                 font=("맑은 고딕", 9), fg="#555").grid(
            row=1, column=0, columnspan=3, pady=(0, 10))

        # 파일 선택
        tk.Label(self, text="HWPX 파일", font=("맑은 고딕", 10)).grid(
            row=2, column=0, sticky="e", **pad)
        self.path_var = tk.StringVar()
        tk.Entry(self, textvariable=self.path_var, width=52,
                 font=("맑은 고딕", 9)).grid(row=2, column=1, **pad)
        tk.Button(self, text="찾아보기", command=self._browse,
                  width=9).grid(row=2, column=2, **pad)

        # 실행 버튼
        self.run_btn = tk.Button(
            self, text="▶  실행", font=("맑은 고딕", 11, "bold"),
            bg="#1a6fc4", fg="white", activebackground="#145ea8",
            width=16, height=1, command=self._run)
        self.run_btn.grid(row=3, column=0, columnspan=3, pady=(2, 8))

        # 로그
        tk.Label(self, text="진행 로그", font=("맑은 고딕", 9),
                 fg="#555").grid(row=4, column=0, columnspan=3, sticky="w", padx=10)
        self.log_box = scrolledtext.ScrolledText(
            self, width=70, height=14, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white")
        self.log_box.grid(row=5, column=0, columnspan=3, padx=10, pady=(0, 12))

    # ── 한글에서 파일 경로 자동 감지 ──────────────────────────
    def _try_autofill(self):
        hwp = try_connect_hwp()
        if not hwp:
            return
        for attr in ("Path", "FileName", "FullName", "DocumentPath"):
            try:
                val = getattr(hwp, attr, None)
                if val and str(val).strip() and os.path.isfile(str(val)):
                    self.path_var.set(str(val).strip())
                    self._log("[OK] 한글에서 파일 경로를 자동으로 가져왔습니다.")
                    return
            except Exception:
                pass

    # ── 파일 선택 대화상자 ─────────────────────────────────────
    def _browse(self):
        path = filedialog.askopenfilename(
            title="HWPX 파일 선택",
            filetypes=[("한글 문서", "*.hwpx"), ("모든 파일", "*.*")])
        if path:
            self.path_var.set(path)

    # ── 로그 출력 ──────────────────────────────────────────────
    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    # ── 실행 ───────────────────────────────────────────────────
    def _run(self):
        filepath = self.path_var.get().strip().strip('"')
        if not filepath:
            messagebox.showwarning("경고", "HWPX 파일을 선택하세요.")
            return
        if not os.path.isfile(filepath):
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{filepath}")
            return
        if not filepath.lower().endswith('.hwpx'):
            messagebox.showerror("오류", "HWPX 파일만 지원합니다.\n한글에서 'HWPX' 형식으로 저장 후 선택하세요.")
            return

        self.run_btn.configure(state="disabled")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        # 별도 스레드에서 실행 (UI 멈춤 방지)
        threading.Thread(target=self._worker, args=(filepath,), daemon=True).start()

    def _worker(self, filepath):
        log = self._log
        try:
            log(f"파일: {filepath}")

            # 한글 자동 저장 시도
            hwp = try_connect_hwp()
            if hwp:
                try:
                    hwp.Save()
                    log("[OK] 한글 문서 저장 완료")
                except Exception:
                    pass

            # 출력 경로
            base, ext = os.path.splitext(filepath)
            dst_path = base + "_renumbered" + ext
            if os.path.exists(dst_path):
                os.remove(dst_path)

            log("\nHWPX 파일 수정 중...")
            tc, fc = renumber_hwpx(filepath, dst_path, log)

            if tc == 0 and fc == 0:
                os.remove(dst_path)
                log("\n[경고] 표/그림 패턴을 찾지 못했습니다.")
                log("  → 문서 내 형식이 [표 N] / <그림 N> 인지 확인하세요.")
                messagebox.showwarning("결과 없음",
                    "표/그림 패턴을 찾지 못했습니다.\n"
                    "형식이 [표 N] 또는 <그림 N>인지 확인하세요.")
                return

            log(f"\n{'='*46}")
            log(f"  완료!  표 {tc}개,  그림 {fc}개  재정렬")
            log(f"{'='*46}")
            log(f"\n저장 위치:\n  {dst_path}")
            messagebox.showinfo("완료",
                f"표 {tc}개, 그림 {fc}개 재정렬 완료!\n\n저장 위치:\n{dst_path}")

        except Exception as e:
            log(f"\n[ERROR] {e}")
            messagebox.showerror("오류", str(e))

        finally:
            self.run_btn.configure(state="normal")


# ==============================================================
# 진입점
# ==============================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()

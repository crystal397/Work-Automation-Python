"""
귀책분석 자동화 시스템 — GUI
더블클릭 실행 시 이 창이 표시됩니다.
"""

from __future__ import annotations

import queue
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

sys.path.insert(0, str(Path(__file__).parent))
import config


# ── stdout → 큐 리디렉션 (스레드 안전) ─────────────────────────────────────

class _RedirectText:
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, msg: str):
        if msg:
            self._q.put(msg)

    def flush(self):
        pass


# ── 메인 애플리케이션 ────────────────────────────────────────────────────────

class App(tk.Tk):
    _BTN_PRIMARY  = {"bg": "#1a5fb4", "fg": "white",  "activebackground": "#1650a0",
                     "activeforeground": "white", "relief": "flat", "bd": 0,
                     "font": ("맑은 고딕", 10, "bold"), "pady": 7, "cursor": "hand2"}
    _BTN_SECONDARY= {"bg": "#e0e0e0", "fg": "#333",   "activebackground": "#cccccc",
                     "relief": "flat", "bd": 0,
                     "font": ("맑은 고딕", 9), "pady": 6, "cursor": "hand2"}

    def __init__(self):
        super().__init__()
        self.title("귀책분석 자동화 시스템")
        self.geometry("900x640")
        self.minsize(760, 500)
        self.configure(bg="#f5f5f5")

        self._q: queue.Queue = queue.Queue()
        self._orig_stdout = sys.stdout
        self._orig_stderr = sys.stderr
        sys.stdout = _RedirectText(self._q)
        sys.stderr = _RedirectText(self._q)

        self._build_ui()
        self._refresh_project()
        self._poll_queue()

    # ── UI 구성 ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 헤더
        hdr = tk.Frame(self, bg="#1a5fb4", height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="귀책분석 자동화 시스템  v1.0",
                 font=("맑은 고딕", 13, "bold"), fg="white", bg="#1a5fb4"
                 ).pack(side="left", padx=18, pady=14)
        self._hdr_proj = tk.Label(hdr, text="", font=("맑은 고딕", 9),
                                   fg="#a8c8f0", bg="#1a5fb4")
        self._hdr_proj.pack(side="right", padx=18)

        # 본문
        body = tk.Frame(self, bg="#f5f5f5")
        body.pack(fill="both", expand=True, padx=12, pady=10)

        # 좌측 컨트롤
        left = tk.Frame(body, bg="#f5f5f5", width=290)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)
        self._build_left(left)

        # 우측 출력
        right = tk.Frame(body, bg="#f5f5f5")
        right.pack(side="left", fill="both", expand=True)
        self._build_right(right)

    def _build_left(self, parent):
        pad = {"padx": 10, "pady": 4}

        # ── STEP 1 ──────────────────────────────────────────
        f1 = tk.LabelFrame(parent, text="  STEP 1  스캔 + 프롬프트 생성  ",
                            font=("맑은 고딕", 9, "bold"), bg="#f5f5f5",
                            fg="#1a5fb4", relief="groove", bd=1)
        f1.pack(fill="x", pady=(0, 8))

        tk.Label(f1, text="수신자료 폴더", font=("맑은 고딕", 9),
                 bg="#f5f5f5", fg="#555").grid(row=0, column=0, sticky="w", **pad)

        path_row = tk.Frame(f1, bg="#f5f5f5")
        path_row.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 2))
        f1.columnconfigure(0, weight=1)
        path_row.columnconfigure(0, weight=1)

        self._path_var = tk.StringVar()
        tk.Entry(path_row, textvariable=self._path_var, font=("맑은 고딕", 9)
                 ).grid(row=0, column=0, sticky="ew", ipady=3)
        tk.Button(path_row, text="찾기", font=("맑은 고딕", 9), width=4,
                  command=self._browse, cursor="hand2", relief="flat",
                  bg="#e0e0e0", activebackground="#ccc"
                  ).grid(row=0, column=1, padx=(4, 0), ipady=3)

        tk.Label(f1, text="프로젝트명  (비워두면 자동 감지)", font=("맑은 고딕", 9),
                 bg="#f5f5f5", fg="#555").grid(row=2, column=0, sticky="w", **pad)
        self._proj_var = tk.StringVar()
        tk.Entry(f1, textvariable=self._proj_var, font=("맑은 고딕", 9)
                 ).grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 6), ipady=3)

        self._btn_scan = tk.Button(f1, text="▶   스캔 시작",
                                    command=self._run_scan, **self._BTN_PRIMARY)
        self._btn_scan.grid(row=4, column=0, sticky="ew", padx=10, pady=(2, 10))

        # ── STEP 2 ──────────────────────────────────────────
        f2 = tk.LabelFrame(parent, text="  STEP 2  docx 생성  ",
                            font=("맑은 고딕", 9, "bold"), bg="#f5f5f5",
                            fg="#1a5fb4", relief="groove", bd=1)
        f2.pack(fill="x", pady=(0, 8))

        tk.Label(f2, text="JSON 작성 완료 후 실행하세요.",
                 font=("맑은 고딕", 9), bg="#f5f5f5", fg="#555"
                 ).pack(anchor="w", padx=10, pady=(8, 4))

        self._btn_finish = tk.Button(f2, text="▶   docx 생성",
                                      command=self._run_finish, **self._BTN_PRIMARY)
        self._btn_finish.pack(fill="x", padx=10, pady=(0, 10))

        # ── STEP 3 ──────────────────────────────────────────
        f3 = tk.LabelFrame(parent, text="  STEP 3  품질 검증  ",
                            font=("맑은 고딕", 9, "bold"), bg="#f5f5f5",
                            fg="#1a5fb4", relief="groove", bd=1)
        f3.pack(fill="x")

        row3 = tk.Frame(f3, bg="#f5f5f5")
        row3.pack(fill="x", padx=10, pady=10)

        self._btn_cmp = tk.Button(row3, text="단일 검증",
                                   command=self._run_compare, **self._BTN_SECONDARY)
        self._btn_cmp.pack(side="left", expand=True, fill="x", padx=(0, 4))

        self._btn_cmp_all = tk.Button(row3, text="전체 검증",
                                       command=self._run_compare_all, **self._BTN_SECONDARY)
        self._btn_cmp_all.pack(side="left", expand=True, fill="x")

        # 버튼 목록 (일괄 활성/비활성용)
        self._all_buttons = [self._btn_scan, self._btn_finish,
                              self._btn_cmp, self._btn_cmp_all]

    def _build_right(self, parent):
        lf = tk.LabelFrame(parent, text="  실행 결과  ",
                            font=("맑은 고딕", 9), bg="#f5f5f5", relief="groove", bd=1)
        lf.pack(fill="both", expand=True)

        self._out = scrolledtext.ScrolledText(
            lf, state="disabled", wrap="word",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white", relief="flat",
        )
        self._out.pack(fill="both", expand=True, padx=4, pady=4)

    # ── 큐 폴링 (stdout → 텍스트 위젯, 메인 스레드에서 실행) ───────────────

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                self._out.configure(state="normal")
                self._out.insert(tk.END, msg)
                self._out.see(tk.END)
                self._out.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(50, self._poll_queue)

    # ── 유틸리티 ────────────────────────────────────────────────────────────

    def _refresh_project(self):
        proj = config.load_current_project()
        self._hdr_proj.config(text=f"현재 프로젝트: {proj}" if proj else "")

    def _set_buttons(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for btn in self._all_buttons:
            btn.config(state=state)

    def _clear_output(self):
        self._out.configure(state="normal")
        self._out.delete("1.0", tk.END)
        self._out.configure(state="disabled")

    def _browse(self):
        folder = filedialog.askdirectory(title="수신자료 폴더 선택")
        if folder:
            self._path_var.set(folder)

    def _run_thread(self, fn):
        """fn 을 별도 스레드에서 실행하고, 완료 시 버튼 복원."""
        self._set_buttons(False)
        self._clear_output()

        def target():
            try:
                fn()
            except SystemExit:
                pass
            except Exception as e:
                print(f"\n[오류] {e}")
            finally:
                self.after(0, self._set_buttons, True)
                self.after(0, self._refresh_project)

        threading.Thread(target=target, daemon=True).start()

    # ── 버튼 핸들러 ─────────────────────────────────────────────────────────

    def _run_scan(self):
        path_str = self._path_var.get().strip()
        if not path_str:
            messagebox.showwarning("경로 필요", "수신자료 폴더 경로를 입력하거나\n[찾기] 버튼으로 선택하세요.")
            return

        project = self._proj_var.get().strip()

        def do():
            from main import cmd_scan_prepare
            raw = [p.strip().strip('"').strip("'")
                   for p in path_str.split(";") if p.strip()]
            args = raw[:]
            if project:
                args += ["--project", project]
            cmd_scan_prepare(args)

            # ── 완료 안내 ─────────────────────────────────────
            proj_now = config.load_current_project()
            if not proj_now:
                return
            proj_dir = config.get_project_dir(proj_now)
            prompt_ok = (proj_dir / "prompt_for_claude.md").exists()

            print("\n" + "═" * 52)
            if prompt_ok:
                print("★ 스캔 완료 — 지금 해야 할 일")
                print("═" * 52)
                print(f"\n저장 위치: output\\{proj_dir.name}\\")
                print()
                print("① scan_summary.md 를 열어 공문 목록을 확인하세요.")
                print("  → 관련 없는 공문이 있으면 scan_result.json 에서 해당 줄 삭제")
                print("  → OCR ⚠️ 항목은 원본 파일로 날짜·공문번호 확인")
                print()
                print("② Claude Code 에 아래 내용을 그대로 붙여넣으세요:")
                print()
                print(f'  output\\{proj_dir.name}\\prompt_for_claude.md 읽고')
                print(f'  귀책분석_data.json 생성해줘')
                print()
                print(f"  → JSON 저장 위치: output\\{proj_dir.name}\\귀책분석_data.json")
                print()
                print("③ JSON 저장 완료 후  [▶ docx 생성]  버튼을 누르세요.")
            else:
                print("⚠️  스캔은 완료됐으나 프롬프트 생성에 실패했습니다.")
                print("═" * 52)
                print("\noutput\\reference_patterns.md 파일이 있는지 확인하세요.")
                print("없으면 관리자에게 문의하세요.")
            print("═" * 52)

        self._run_thread(do)

    def _run_finish(self):
        def do():
            from main import cmd_finish
            cmd_finish(None)
        self._run_thread(do)

    def _run_compare(self):
        def do():
            from main import cmd_compare
            cmd_compare(None)
        self._run_thread(do)

    def _run_compare_all(self):
        def do():
            from main import cmd_compare_all
            cmd_compare_all()
        self._run_thread(do)

    def destroy(self):
        sys.stdout = self._orig_stdout
        sys.stderr = self._orig_stderr
        super().destroy()


def run_gui():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    run_gui()

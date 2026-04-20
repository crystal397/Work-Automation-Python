"""CustomTkinter GUI — 법령 시점 매칭 시스템 (모던 UI)

레이아웃:
  좌측 패널(290px): 기관코드 / 입찰공고일 / 법령 체크박스 / 실행 버튼
  우측 패널(확장):  CTkTabview(매칭결과 / 조문상세 / 실행로그) + Word 저장
  하단:            상태바
"""
import logging
import threading
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk

try:
    from tkcalendar import DateEntry
    _HAS_CAL = True
except ImportError:
    _HAS_CAL = False

import config
from engine import LawMatcher, MatchResult
from report import WordReportGenerator

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

logger = logging.getLogger(__name__)

# ── 색상 상수 ──────────────────────────────────────────────────────────────────
_C_PANEL  = "#f5f5f5"   # 좌측 패널 배경
_C_SEP    = "#d8d8d8"   # 구분선
_C_TEXT   = "#1a1a1a"   # 기본 텍스트
_C_GRAY   = "#888888"   # 보조 텍스트
_C_BLUE   = "#1a6eb5"   # 강조 텍스트
_C_GREEN  = "#157347"   # 저장 버튼


def _sep(parent: ctk.CTkFrame) -> None:
    """가로 구분선"""
    ctk.CTkFrame(parent, height=1, fg_color=_C_SEP).pack(
        fill="x", padx=12, pady=4
    )


class App(ctk.CTk):
    """메인 애플리케이션"""

    def __init__(self) -> None:
        super().__init__()
        self.title("법령 시점 매칭 시스템 v1.0")
        self.geometry("1160x740")
        self.minsize(920, 580)
        self._results: list[MatchResult] = []
        self._law_vars: list[tk.BooleanVar] = []
        self._build_ui()
        self._apply_treeview_style()
        self._wire_logging()

    # ── 스타일 ────────────────────────────────────────────────────────────────

    def _apply_treeview_style(self) -> None:
        """ttk.Treeview를 모던하게 스타일링"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Treeview",
            background="#ffffff",
            foreground=_C_TEXT,
            fieldbackground="#ffffff",
            rowheight=28,
            font=("맑은 고딕", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#e4e4e4",
            foreground="#333333",
            font=("맑은 고딕", 9, "bold"),
            relief="flat",
            padding=6,
        )
        style.map(
            "Treeview",
            background=[("selected", "#0078d4")],
            foreground=[("selected", "#ffffff")],
        )
        style.map("Treeview.Heading", background=[("active", "#d0d0d0")])

    # ── UI 구성 ───────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── 좌측 패널 ──────────────────────────────────────────────────────────
        left = ctk.CTkFrame(self, width=290, corner_radius=0, fg_color=_C_PANEL)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        # 기관코드
        ctk.CTkLabel(
            left, text="법제처 기관코드 (OC)", anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(fill="x", padx=14, pady=(16, 2))
        self.oc_var = tk.StringVar(value=config.LAW_API_OC)
        ctk.CTkEntry(
            left, textvariable=self.oc_var, show="*",
            placeholder_text="이메일 형식 OC",
        ).pack(fill="x", padx=14, pady=(0, 2))
        ctk.CTkLabel(
            left, text="open.law.go.kr 가입 이메일 ID",
            text_color=_C_GRAY, font=ctk.CTkFont(size=10), anchor="w",
        ).pack(fill="x", padx=16, pady=(0, 10))

        _sep(left)

        # 입찰공고일
        ctk.CTkLabel(
            left, text="입찰공고일", anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(fill="x", padx=14, pady=(10, 2))
        if _HAS_CAL:
            self.date_entry = DateEntry(
                left, width=18, date_pattern="yyyy-mm-dd",
                locale="ko_KR", background="#1f6aa5", foreground="white",
            )
            self.date_entry.pack(fill="x", padx=14, pady=(0, 10))
        else:
            self.date_var = tk.StringVar(value=str(date.today()))
            ctk.CTkEntry(
                left, textvariable=self.date_var,
                placeholder_text="YYYY-MM-DD",
            ).pack(fill="x", padx=14, pady=(0, 2))
            ctk.CTkLabel(
                left, text="형식: YYYY-MM-DD",
                text_color=_C_GRAY, font=ctk.CTkFont(size=10), anchor="w",
            ).pack(fill="x", padx=16, pady=(0, 10))

        _sep(left)

        # 대상 법령 선택
        ctk.CTkLabel(
            left, text="대상 법령 선택", anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(fill="x", padx=14, pady=(10, 4))

        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.pack(fill="x", padx=14, pady=(0, 4))
        ctk.CTkButton(
            btn_row, text="전체 선택", command=self._select_all,
            width=118, height=28,
        ).pack(side="left")
        ctk.CTkButton(
            btn_row, text="전체 해제", command=self._deselect_all,
            width=118, height=28, fg_color="#6c757d", hover_color="#5a6268",
        ).pack(side="left", padx=(6, 0))

        scroll_frame = ctk.CTkScrollableFrame(
            left, fg_color="#ebebeb",
            scrollbar_button_color="#c8c8c8",
            scrollbar_button_hover_color="#aaaaaa",
        )
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        self._law_vars = []
        for name, _, _ in config.TARGET_LAWS:
            var = tk.BooleanVar(value=True)
            ctk.CTkCheckBox(
                scroll_frame, text=name, variable=var,
                font=ctk.CTkFont(size=11),
                checkbox_width=18, checkbox_height=18,
            ).pack(anchor="w", pady=2, padx=4)
            self._law_vars.append(var)

        _sep(left)

        # 실행 버튼 + 프로그레스
        ctk.CTkButton(
            left, text="▶  매칭 실행", command=self._run_matching,
            height=38, font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=6,
        ).pack(fill="x", padx=14, pady=(8, 6))
        self.progress = ctk.CTkProgressBar(left, mode="determinate", height=8)
        self.progress.pack(fill="x", padx=14, pady=(0, 12))
        self.progress.set(0)

        # ── 우측 패널 ──────────────────────────────────────────────────────────
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.pack(side="right", fill="both", expand=True, padx=(4, 10), pady=(8, 4))

        self.tabview = ctk.CTkTabview(right, anchor="nw")
        self.tabview.pack(fill="both", expand=True)

        self._build_result_tab(self.tabview.add("  매칭 결과  "))
        self._build_article_tab(self.tabview.add("  조문 상세  "))
        self._build_log_tab(self.tabview.add("  실행 로그  "))

        # Word 저장 버튼
        ctk.CTkButton(
            right, text="💾  Word 리포트 생성 (.docx)",
            command=self._save_report, height=34,
            fg_color=_C_GREEN, hover_color="#116137",
            corner_radius=6,
        ).pack(fill="x", pady=(6, 0))

        # ── 상태바 ──────────────────────────────────────────────────────────────
        self.status_var = tk.StringVar(value="준비")
        ctk.CTkLabel(
            self, textvariable=self.status_var, anchor="w",
            height=26, fg_color="#e0e0e0", text_color="#444444",
            corner_radius=0, font=ctk.CTkFont(size=10),
        ).pack(side="bottom", fill="x")

    def _build_result_tab(self, parent) -> None:
        cols = ("법령명", "공포번호", "시행일", "부칙 경과규정", "비고")

        vsb = ttk.Scrollbar(parent, orient="vertical")
        hsb = ttk.Scrollbar(parent, orient="horizontal")
        self.tree = ttk.Treeview(
            parent, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        widths = [200, 110, 90, 100, 600]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=60, anchor="center", stretch=False)
        self.tree.column("법령명", anchor="w", stretch=True)
        self.tree.column("비고", anchor="w", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def _build_article_tab(self, parent) -> None:
        vsb = ttk.Scrollbar(parent, orient="vertical")
        self.article_text = tk.Text(
            parent, wrap="word", font=("맑은 고딕", 9),
            yscrollcommand=vsb.set, state="disabled",
            relief="flat", padx=12, pady=10,
            background="#ffffff", foreground=_C_TEXT,
        )
        vsb.config(command=self.article_text.yview)
        vsb.pack(side="right", fill="y")
        self.article_text.pack(fill="both", expand=True)

        self.article_text.tag_configure(
            "title", font=("맑은 고딕", 12, "bold"), foreground="#111111"
        )
        self.article_text.tag_configure(
            "label", font=("맑은 고딕", 9, "bold"), foreground=_C_BLUE
        )
        self.article_text.tag_configure(
            "warn", foreground="#c0392b", font=("맑은 고딕", 9, "bold")
        )
        self.article_text.tag_configure("url", foreground=_C_GRAY)
        self.article_text.tag_configure(
            "art_title", font=("맑은 고딕", 10, "bold"), foreground="#154360"
        )

    def _build_log_tab(self, parent) -> None:
        vsb = ttk.Scrollbar(parent, orient="vertical")
        self.log_text = tk.Text(
            parent, wrap="word", font=("Consolas", 8),
            yscrollcommand=vsb.set, state="disabled",
            background="#1e1e1e", foreground="#d4d4d4",
            insertbackground="white", relief="flat", padx=6, pady=4,
        )
        vsb.config(command=self.log_text.yview)
        vsb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

    # ── 이벤트 핸들러 ─────────────────────────────────────────────────────────

    def _select_all(self) -> None:
        for var in self._law_vars:
            var.set(True)

    def _deselect_all(self) -> None:
        for var in self._law_vars:
            var.set(False)

    def _get_bid_date(self) -> date:
        if _HAS_CAL:
            return self.date_entry.get_date()
        return date.fromisoformat(self.date_var.get().strip())

    def _get_selected_laws(self) -> list[tuple[str, str, str]]:
        return [
            config.TARGET_LAWS[i]
            for i, var in enumerate(self._law_vars)
            if var.get()
        ]

    def _run_matching(self) -> None:
        oc = self.oc_var.get().strip()
        if not oc:
            messagebox.showerror("오류", "법제처 기관코드(OC)를 입력하세요.")
            return

        selected = self._get_selected_laws()
        if not selected:
            messagebox.showwarning("경고", "대상 법령을 하나 이상 선택하세요.")
            return

        try:
            bid_date = self._get_bid_date()
        except ValueError:
            messagebox.showerror("오류", "입찰공고일 형식이 올바르지 않습니다 (YYYY-MM-DD).")
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        self._set_text(self.article_text, "")
        self._results = []

        threading.Thread(
            target=self._thread_match,
            args=(oc, bid_date, selected),
            daemon=True,
        ).start()

    def _thread_match(self, oc: str, bid_date: date, laws: list) -> None:
        total = len(laws)

        def progress(current: int, _total: int, name: str) -> None:
            frac = current / _total if _total else 0
            self.after(0, lambda f=frac: self.progress.set(f))
            self.after(
                0,
                lambda: self.status_var.set(
                    f"처리 중… {name} ({current}/{_total})"
                ),
            )

        self.after(0, lambda: self.status_var.set("매칭 실행 중…"))
        try:
            matcher = LawMatcher(oc=oc)
            results = matcher.match_all(bid_date, laws, progress_callback=progress)
            self._results = results
            self.after(0, lambda: self._display_results(results))
        except Exception as exc:
            logger.exception("매칭 중 오류 발생")
            self.after(0, lambda: messagebox.showerror("오류", str(exc)))
        finally:
            self.after(0, lambda: self.status_var.set(f"완료 — {total}개 법령 처리"))
            self.after(0, lambda: self.progress.set(0))

    def _display_results(self, results: list[MatchResult]) -> None:
        for r in results:
            if r.selected:
                v = r.selected
                values = (
                    r.display_name,
                    v.announce_num,
                    str(v.enforce_date),
                    (
                        (f"있음 ⚠ ({r.transitional_type}형)" if r.transitional_type else "있음 ⚠")
                        if r.transitional_flag
                        else "없음"
                    ),
                    r.consistency_warning or r.warning or (
                        "확인 필요" if r.needs_user_review else ""
                    ),
                )
                if r.transitional_flag or r.consistency_warning:
                    tag = "warn"
                elif v.target == "admrul" and "연혁 조회 불가" in (r.warning or ""):
                    tag = "admrul"
                elif r.needs_user_review:
                    tag = "review"
                else:
                    tag = ""
            else:
                values = (r.display_name, "-", "-", "-", r.warning or "조회 실패")
                tag = "error"

            self.tree.insert("", "end", values=values, tags=(tag,) if tag else ())

        self.tree.tag_configure("warn",   background="#fff3cd")
        self.tree.tag_configure("admrul", background="#fff0e0")
        self.tree.tag_configure("review", background="#fff8e7")
        self.tree.tag_configure("error",  background="#fde8e8")

    def _on_select(self, _event) -> None:
        sel = self.tree.selection()
        if not sel or not self._results:
            return
        idx = self.tree.index(sel[0])
        if idx >= len(self._results):
            return
        self._show_article(self._results[idx])

    def _show_article(self, r: MatchResult) -> None:
        w = self.article_text
        w.config(state="normal")
        w.delete("1.0", "end")

        w.insert("end", f"【 {r.display_name} 】\n", "title")
        w.insert("end", "\n")

        if not r.selected:
            w.insert("end", f"⚠ {r.warning}\n", "warn")
            w.config(state="disabled")
            return

        v = r.selected
        w.insert("end", "법령명: ",   "label"); w.insert("end", f"{v.name}\n")
        w.insert("end", "공포번호: ", "label"); w.insert("end", f"{v.announce_num}\n")
        w.insert("end", "공포일: ",   "label"); w.insert("end", f"{v.announce_date}\n")
        w.insert("end", "시행일: ",   "label"); w.insert("end", f"{v.enforce_date}\n")
        w.insert("end", "출처 URL: ", "label"); w.insert("end", f"{v.source_url}\n", "url")
        w.insert("end", "\n")

        if r.warning:
            w.insert("end", f"ℹ {r.warning}\n", "warn")
            w.insert("end", "\n")

        if r.consistency_warning:
            w.insert("end", f"{r.consistency_warning}\n", "warn")
            w.insert("end", "\n")

        if r.transitional_flag:
            type_label = {
                "A": "유형 A (법령 전체 경과규정)",
                "B": "유형 B (조문 단위 경과규정)",
            }.get(r.transitional_type, "유형 미확인")
            w.insert("end", f"⚠ 부칙 경과규정 탐지 [{type_label}] — 사용자 확인 필요\n", "warn")
            w.insert("end", f"  발견 문장: {r.transitional_text}\n", "warn")
            if r.transitional_type == "B" and r.transitional_articles:
                art_list = ", ".join(f"제{n}조" for n in r.transitional_articles)
                w.insert("end", f"  → 영향 조문: {art_list}\n", "warn")
            if r.prev_version:
                pv = r.prev_version
                w.insert("end", f"  직전 버전: {pv.announce_num} | 시행일: {pv.enforce_date}\n")
            w.insert("end", "\n")

        if r.relevant_articles:
            w.insert("end", f"공기연장 관련 조문 ({len(r.relevant_articles)}개)\n", "label")
            w.insert("end", "─" * 50 + "\n")
            for art in r.relevant_articles:
                w.insert("end", f"제{art['조번호']}조 {art['조제목']}\n", "art_title")
                if art["조문내용"]:
                    w.insert("end", f"{art['조문내용']}\n")
                for para in art.get("항", []):
                    try:
                        num = chr(0x245F + int(str(para.get("항번호") or "")))
                    except (ValueError, OverflowError):
                        num = str(para.get("항번호") or "")
                    content = str(para.get("항내용") or "")
                    if content:
                        w.insert("end", f"  {num} {content}\n")
                    for sub in para.get("호", []):
                        sub_num = str(sub.get("호번호") or "")
                        sub_content = str(sub.get("호내용") or "")
                        if sub_content:
                            w.insert("end", f"    {sub_num} {sub_content}\n")
                        for ss in sub.get("목", []):
                            ss_num = str(ss.get("목번호") or "")
                            ss_content = str(ss.get("목내용") or "")
                            if ss_content:
                                w.insert("end", f"      {ss_num}) {ss_content}\n")
                w.insert("end", "\n")
        else:
            w.insert("end", "(공기연장 관련 조문 없음 또는 키워드 미탐지)\n")

        w.config(state="disabled")

    def _save_report(self) -> None:
        if not self._results:
            messagebox.showwarning("경고", "먼저 매칭을 실행하세요.")
            return
        try:
            bid_date = self._get_bid_date()
        except ValueError:
            messagebox.showerror("오류", "입찰공고일을 확인하세요.")
            return

        default_name = f"법령매칭_{bid_date}.docx"
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word 문서", "*.docx")],
            initialfile=default_name,
            title="리포트 저장 위치 선택",
        )
        if not path:
            return

        try:
            gen = WordReportGenerator()
            gen.generate(bid_date, self._results, Path(path))
            messagebox.showinfo("저장 완료", f"리포트가 저장되었습니다:\n{path}")
        except Exception as exc:
            logger.exception("리포트 생성 오류")
            messagebox.showerror("오류", str(exc))

    # ── 유틸리티 ─────────────────────────────────────────────────────────────

    def _set_text(self, widget: tk.Text, content: str) -> None:
        widget.config(state="normal")
        widget.delete("1.0", "end")
        if content:
            widget.insert("1.0", content)
        widget.config(state="disabled")

    def _wire_logging(self) -> None:
        """실행 로그 탭으로 로그 스트리밍"""

        class _GuiHandler(logging.Handler):
            def __init__(self_, widget: tk.Text, app: "App") -> None:
                super().__init__()
                self_._widget = widget
                self_._app = app

            def emit(self_, record: logging.LogRecord) -> None:
                msg = self_.format(record) + "\n"
                self_._app.after(0, lambda m=msg: self_._append(m))

            def _append(self_, msg: str) -> None:
                w = self_._widget
                w.config(state="normal")
                w.insert("end", msg)
                w.see("end")
                w.config(state="disabled")

        handler = _GuiHandler(self.log_text, self)
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
        )
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)


if __name__ == "__main__":
    app = App()
    app.mainloop()

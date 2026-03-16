"""
Excel → PDF 변환기 v2
- 선택한 시트만 PDF로 변환
- 제외 규칙(키워드/정확한 이름) 사전 설정 및 JSON 저장
실행 환경: Windows + Microsoft Excel 설치 필요
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import json


# ──────────────────────────────────────────────
# 설정 파일 경로 (실행 파일 옆에 저장)
# ──────────────────────────────────────────────
def _config_path() -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))
    return os.path.join(base, "excel_pdf_config.json")


DEFAULT_CONFIG = {
    "rules": []
    # rule 형태: {"type": "keyword" | "exact", "value": str, "enabled": bool}
}


def load_config() -> dict:
    path = _config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    path = _config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def should_exclude(sheet_name: str, rules: list) -> bool:
    """활성화된 규칙 중 하나라도 매칭되면 제외"""
    for rule in rules:
        if not rule.get("enabled", True):
            continue
        val = rule.get("value", "")
        if not val:
            continue
        if rule.get("type") == "exact":
            if sheet_name == val:
                return True
        else:  # keyword
            if val.lower() in sheet_name.lower():
                return True
    return False


# ──────────────────────────────────────────────
# 의존성 체크
# ──────────────────────────────────────────────
def check_dependencies():
    missing = []
    try:
        import win32com.client  # noqa
    except ImportError:
        missing.append("pywin32")
    try:
        import openpyxl  # noqa
    except ImportError:
        missing.append("openpyxl")
    return missing


# ──────────────────────────────────────────────
# Excel 핵심 로직
# ──────────────────────────────────────────────
def get_sheet_names(excel_path: str) -> list:
    import openpyxl
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def convert_sheets_to_pdf(
    excel_path: str,
    output_pdf: str,
    selected_sheets: list,
    progress_cb=None,
):
    """
    변환 전략:
    1) 선택되지 않은 시트를 xlSheetHidden(0)으로 숨김
       (xlSheetVeryHidden 시트는 이미 숨겨져 있으므로 건드리지 않음)
    2) 워크북 전체 ExportAsFixedFormat → 보이는 시트만 PDF 출력
    3) 숨긴 시트 원래 상태로 복원
    """
    import win32com.client

    XL_SHEET_VISIBLE    = -1   # xlSheetVisible
    XL_SHEET_HIDDEN     =  0   # xlSheetHidden
    XL_SHEET_VERY_HIDDEN = 2   # xlSheetVeryHidden

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible        = False
    excel.DisplayAlerts  = False
    excel.ScreenUpdating = False

    wb = None
    try:
        wb = excel.Workbooks.Open(
            os.path.abspath(excel_path),
            UpdateLinks=False,
            ReadOnly=True,
        )

        # ── 숨길 시트 목록 및 원래 상태 기록 ──────────────
        hidden_map: dict = {}   # {sheet_name: original_visible_state}

        for sh in wb.Sheets:
            if sh.Name not in selected_sheets:
                orig = sh.Visible
                # 이미 완전히 숨겨진(VeryHidden) 시트는 건드리지 않음
                if orig != XL_SHEET_VERY_HIDDEN:
                    hidden_map[sh.Name] = orig
                    sh.Visible = XL_SHEET_HIDDEN

        # ── 첫 번째 선택 시트 활성화 (안전용) ───────────
        try:
            wb.Sheets(selected_sheets[0]).Activate()
        except Exception:
            pass

        if progress_cb:
            progress_cb(40, "PDF 변환 중...")

        # ── PDF 출력 (보이는 시트 전체) ─────────────────
        wb.ExportAsFixedFormat(
            Type=0,                          # xlTypePDF
            Filename=os.path.abspath(output_pdf),
            Quality=0,                       # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

        if progress_cb:
            progress_cb(80, "시트 상태 복원 중...")

        # ── 숨김 원복 ────────────────────────────────
        for sh in wb.Sheets:
            if sh.Name in hidden_map:
                sh.Visible = hidden_map[sh.Name]

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass

    if progress_cb:
        progress_cb(100, "완료!")


# ──────────────────────────────────────────────
# 제외 규칙 설정 창
# ──────────────────────────────────────────────
class RulesDialog(tk.Toplevel):
    """
    제외할 시트 규칙을 관리하는 모달 창.
    - 키워드 포함: 시트 이름에 특정 문자열이 포함되면 제외
    - 정확한 이름: 시트 이름이 정확히 일치하면 제외
    - 각 규칙별 활성/비활성 토글
    """

    BG     = "#F0F4F8"
    CARD   = "#FFFFFF"
    ACCENT = "#2563EB"
    RED    = "#EF4444"
    LABEL_FG = "#1E293B"
    SUB_FG   = "#64748B"

    def __init__(self, parent, config, on_save):
        super().__init__(parent)
        self.title("제외 규칙 설정")
        self.resizable(False, False)
        self.configure(bg=self.BG)
        self.grab_set()

        self.config_data = config
        self.on_save = on_save
        self.rules = [dict(r) for r in config.get("rules", [])]

        self._build_ui()
        self._center(parent, 520, 500)

    def _center(self, parent, w, h):
        px = parent.winfo_x() + parent.winfo_width() // 2
        py = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"{w}x{h}+{px - w // 2}+{py - h // 2}")

    def _build_ui(self):
        # 헤더
        hdr = tk.Frame(self, bg=self.ACCENT, height=48)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🚫  제외 규칙 설정",
                 font=("Malgun Gothic", 12, "bold"),
                 fg="white", bg=self.ACCENT).pack(side="left", padx=14, pady=10)

        # 설명
        desc = tk.Frame(self, bg="#EFF6FF")
        desc.pack(fill="x")
        tk.Label(
            desc,
            text="아래 규칙에 해당하는 시트는 Excel 파일을 열 때 자동으로 체크 해제됩니다.",
            font=("Malgun Gothic", 9), fg="#1D4ED8", bg="#EFF6FF",
        ).pack(anchor="w", padx=14, pady=6)

        outer = tk.Frame(self, bg=self.BG)
        outer.pack(fill="both", expand=True, padx=14, pady=(10, 0))

        tk.Label(outer, text="규칙 목록", font=("Malgun Gothic", 10, "bold"),
                 fg=self.LABEL_FG, bg=self.BG).pack(anchor="w")

        list_card = tk.Frame(outer, bg=self.CARD,
                              highlightthickness=1, highlightbackground="#CBD5E1")
        list_card.pack(fill="both", expand=True, pady=(4, 0))

        canvas = tk.Canvas(list_card, bg=self.CARD, highlightthickness=0)
        sb = ttk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
        self.list_frame = tk.Frame(canvas, bg=self.CARD)
        self.list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._render_rules()

        # 새 규칙 추가
        add_frame = tk.Frame(outer, bg=self.BG)
        add_frame.pack(fill="x", pady=(10, 0))
        tk.Label(add_frame, text="새 규칙 추가", font=("Malgun Gothic", 10, "bold"),
                 fg=self.LABEL_FG, bg=self.BG).pack(anchor="w")

        input_row = tk.Frame(add_frame, bg=self.BG)
        input_row.pack(fill="x", pady=(4, 0))

        self.new_type = tk.StringVar(value="keyword")
        type_menu = ttk.Combobox(
            input_row, textvariable=self.new_type,
            values=["keyword", "exact"], state="readonly", width=10,
            font=("Malgun Gothic", 10),
        )
        type_menu.pack(side="left", padx=(0, 6), ipady=3)

        self.new_value = tk.StringVar()
        tk.Entry(
            input_row, textvariable=self.new_value,
            font=("Malgun Gothic", 10), relief="flat",
            bg=self.CARD, width=28,
        ).pack(side="left", ipady=5, padx=(0, 6))

        tk.Button(
            input_row, text="+ 추가", command=self._add_rule,
            bg=self.ACCENT, fg="white",
            font=("Malgun Gothic", 10, "bold"),
            relief="flat", cursor="hand2",
            activebackground=self.ACCENT, activeforeground="white",
            padx=10,
        ).pack(side="left")

        tk.Label(
            add_frame,
            text="keyword: 이름에 해당 문자열 포함 시 제외  |  exact: 이름이 정확히 일치 시 제외",
            font=("Malgun Gothic", 8), fg=self.SUB_FG, bg=self.BG,
        ).pack(anchor="w", pady=(3, 0))

        # 저장/취소
        btn_row = tk.Frame(self, bg=self.BG)
        btn_row.pack(fill="x", padx=14, pady=10)
        tk.Button(
            btn_row, text="저장", command=self._save,
            bg=self.ACCENT, fg="white",
            font=("Malgun Gothic", 10, "bold"),
            relief="flat", cursor="hand2",
            activebackground=self.ACCENT, activeforeground="white",
            padx=20, pady=4,
        ).pack(side="right", padx=(6, 0))
        tk.Button(
            btn_row, text="취소", command=self.destroy,
            bg="#94A3B8", fg="white",
            font=("Malgun Gothic", 10, "bold"),
            relief="flat", cursor="hand2",
            activebackground="#94A3B8", activeforeground="white",
            padx=20, pady=4,
        ).pack(side="right")

    def _render_rules(self):
        for w in self.list_frame.winfo_children():
            w.destroy()

        if not self.rules:
            tk.Label(
                self.list_frame,
                text="등록된 규칙이 없습니다. 아래에서 추가해보세요.",
                font=("Malgun Gothic", 9), fg=self.SUB_FG, bg=self.CARD,
            ).pack(pady=20)
            return

        for i, rule in enumerate(self.rules):
            row = tk.Frame(self.list_frame, bg=self.CARD,
                            highlightthickness=1, highlightbackground="#E2E8F0")
            row.pack(fill="x", padx=6, pady=3)

            enabled_var = tk.BooleanVar(value=rule.get("enabled", True))

            def _toggle(idx=i, var=enabled_var):
                self.rules[idx]["enabled"] = var.get()
                self._render_rules()

            tk.Checkbutton(
                row, variable=enabled_var, command=_toggle,
                bg=self.CARD,
            ).pack(side="left", padx=(6, 0))

            t = rule.get("type", "keyword")
            badge_bg = "#DBEAFE" if t == "keyword" else "#F3E8FF"
            badge_fg = "#1D4ED8" if t == "keyword" else "#7C3AED"
            tk.Label(
                row, text=t,
                font=("Malgun Gothic", 8, "bold"),
                fg=badge_fg, bg=badge_bg,
                padx=6, pady=2,
            ).pack(side="left", padx=(4, 0))

            fg_color = self.LABEL_FG if rule.get("enabled", True) else "#94A3B8"
            tk.Label(
                row, text=rule.get("value", ""),
                font=("Malgun Gothic", 10),
                fg=fg_color, bg=self.CARD, anchor="w",
            ).pack(side="left", padx=10, fill="x", expand=True)

            def _delete(idx=i):
                self.rules.pop(idx)
                self._render_rules()

            tk.Button(
                row, text="✕", command=_delete,
                bg=self.CARD, fg=self.RED,
                font=("Malgun Gothic", 9, "bold"),
                relief="flat", cursor="hand2",
                activebackground=self.CARD, activeforeground=self.RED,
                bd=0, padx=8,
            ).pack(side="right")

    def _add_rule(self):
        val = self.new_value.get().strip()
        if not val:
            messagebox.showwarning("입력 오류", "규칙 값을 입력해주세요.", parent=self)
            return
        self.rules.append({
            "type": self.new_type.get(),
            "value": val,
            "enabled": True,
        })
        self.new_value.set("")
        self._render_rules()

    def _save(self):
        clean = [
            {"type": r["type"], "value": r["value"], "enabled": r.get("enabled", True)}
            for r in self.rules
        ]
        self.config_data["rules"] = clean
        save_config(self.config_data)
        self.on_save(self.config_data)
        self.destroy()


# ──────────────────────────────────────────────
# 메인 앱
# ──────────────────────────────────────────────
class App(tk.Tk):
    BG     = "#F0F4F8"
    CARD   = "#FFFFFF"
    ACCENT = "#2563EB"
    LABEL_FG = "#1E293B"
    SUB_FG   = "#64748B"
    GREEN  = "#22C55E"
    RED    = "#EF4444"
    ORANGE = "#F59E0B"

    def __init__(self):
        super().__init__()
        self.title("Excel → PDF 변환기")
        self.resizable(False, False)
        self.configure(bg=self.BG)

        self.config_data = load_config()
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_vars = {}

        self._build_ui()
        self._center_window(630, 640)

    def _center_window(self, w, h):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        PAD = 16

        # 헤더
        hdr = tk.Frame(self, bg=self.ACCENT, height=56)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📊  Excel → PDF 변환기",
                 font=("Malgun Gothic", 14, "bold"),
                 fg="white", bg=self.ACCENT).pack(side="left", padx=PAD, pady=12)
        tk.Button(
            hdr, text="⚙  제외 규칙 설정",
            command=self._open_rules,
            bg="#1D4ED8", fg="white",
            font=("Malgun Gothic", 9, "bold"),
            relief="flat", cursor="hand2",
            activebackground="#1D4ED8", activeforeground="white",
            padx=10,
        ).pack(side="right", padx=PAD, pady=16)

        main = tk.Frame(self, bg=self.BG)
        main.pack(fill="both", expand=True, padx=PAD, pady=PAD)

        # ① 파일 선택
        self._section(main, "① Excel 파일 선택")
        row1 = tk.Frame(main, bg=self.BG)
        row1.pack(fill="x", pady=(4, 10))
        tk.Entry(row1, textvariable=self.excel_path,
                 font=("Malgun Gothic", 10), state="readonly",
                 relief="flat", bg=self.CARD, width=52,
                 ).pack(side="left", ipady=6, padx=(0, 6))
        self._btn(row1, "찾아보기", self._browse_excel, self.ACCENT).pack(side="left")

        # ② 시트 선택
        self._section(main, "② 변환할 시트 선택 (체크된 시트만 PDF에 포함)")
        sheet_card = tk.Frame(main, bg=self.CARD,
                               highlightthickness=1, highlightbackground="#CBD5E1")
        sheet_card.pack(fill="x", pady=(4, 6))

        canvas = tk.Canvas(sheet_card, bg=self.CARD, height=190, highlightthickness=0)
        sb = ttk.Scrollbar(sheet_card, orient="vertical", command=canvas.yview)
        self.sheet_frame = tk.Frame(canvas, bg=self.CARD)
        self.sheet_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.sheet_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 버튼 행
        btn_row = tk.Frame(main, bg=self.BG)
        btn_row.pack(fill="x", pady=(0, 4))
        self._btn(btn_row, "전체 선택",   self._select_all,     self.GREEN,  small=True).pack(side="left", padx=(0, 6))
        self._btn(btn_row, "전체 해제",   self._deselect_all,   self.RED,    small=True).pack(side="left", padx=(0, 10))
        self._btn(btn_row, "↺ 규칙 재적용", self._reapply_rules, self.ORANGE, small=True).pack(side="left")
        self.sheet_count_label = tk.Label(btn_row, text="",
                                           font=("Malgun Gothic", 9),
                                           fg=self.SUB_FG, bg=self.BG)
        self.sheet_count_label.pack(side="right")

        # 활성 규칙 힌트
        self.rule_hint = tk.Label(main, text="",
                                   font=("Malgun Gothic", 8),
                                   fg="#7C3AED", bg=self.BG,
                                   wraplength=570, justify="left")
        self.rule_hint.pack(anchor="w", pady=(0, 6))
        self._update_rule_hint()

        # ③ 저장 위치
        self._section(main, "③ PDF 저장 위치")
        row2 = tk.Frame(main, bg=self.BG)
        row2.pack(fill="x", pady=(4, 14))
        tk.Entry(row2, textvariable=self.output_path,
                 font=("Malgun Gothic", 10), state="readonly",
                 relief="flat", bg=self.CARD, width=52,
                 ).pack(side="left", ipady=6, padx=(0, 6))
        self._btn(row2, "찾아보기", self._browse_output, self.ACCENT).pack(side="left")

        # 진행 바
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main, variable=self.progress_var,
                        maximum=100, length=570).pack(pady=(0, 4))
        self.status_label = tk.Label(main, text="Excel 파일을 선택하세요.",
                                      font=("Malgun Gothic", 9),
                                      fg=self.SUB_FG, bg=self.BG)
        self.status_label.pack()

        # 변환 버튼
        self.convert_btn = self._btn(
            main, "  PDF로 변환하기  ", self._start_convert,
            self.ACCENT, font_size=12
        )
        self.convert_btn.pack(pady=12, ipadx=8, ipady=6)

    def _section(self, parent, text):
        tk.Label(parent, text=text,
                 font=("Malgun Gothic", 10, "bold"),
                 fg=self.LABEL_FG, bg=self.BG).pack(anchor="w", pady=(6, 0))

    def _btn(self, parent, text, command, color, small=False, font_size=10):
        return tk.Button(
            parent, text=text, command=command,
            bg=color, fg="white",
            font=("Malgun Gothic", font_size, "bold"),
            relief="flat", cursor="hand2",
            activebackground=color, activeforeground="white",
            bd=0, padx=10 if not small else 6,
        )

    # ── 규칙 힌트 ────────────────────────────────
    def _update_rule_hint(self):
        rules = self.config_data.get("rules", [])
        active = [r for r in rules if r.get("enabled", True) and r.get("value")]
        if not active:
            self.rule_hint.config(text="")
            return
        parts = []
        for r in active:
            prefix = "포함" if r["type"] == "keyword" else "일치"
            parts.append(f'"{r["value"]}"({prefix})')
        self.rule_hint.config(text="🚫 자동 제외 규칙: " + "  ·  ".join(parts))

    # ── 규칙 설정 창 ─────────────────────────────
    def _open_rules(self):
        def on_save(new_cfg):
            self.config_data = new_cfg
            self._update_rule_hint()
            if self.sheet_vars:
                self._reapply_rules()

        RulesDialog(self, self.config_data, on_save)

    # ── 파일 탐색 ─────────────────────────────────
    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        self.excel_path.set(path)
        base = os.path.splitext(path)[0]
        self.output_path.set(base + "_converted.pdf")
        self._load_sheets(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="PDF 저장 위치",
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
        )
        if path:
            self.output_path.set(path)

    # ── 시트 목록 로드 ────────────────────────────
    def _load_sheets(self, path):
        for w in self.sheet_frame.winfo_children():
            w.destroy()
        self.sheet_vars.clear()

        try:
            names = get_sheet_names(path)
        except Exception as e:
            messagebox.showerror("오류", f"시트 목록을 읽을 수 없습니다.\n{e}")
            return

        rules = self.config_data.get("rules", [])
        excluded_count = 0

        for name in names:
            exclude = should_exclude(name, rules)
            var = tk.BooleanVar(value=not exclude)
            self.sheet_vars[name] = var
            if exclude:
                excluded_count += 1

            row = tk.Frame(self.sheet_frame, bg=self.CARD)
            row.pack(fill="x", padx=8, pady=2)

            tk.Checkbutton(
                row, variable=var, bg=self.CARD,
                command=self._update_count,
            ).pack(side="left")

            tk.Label(
                row, text=name,
                font=("Malgun Gothic", 10),
                fg=self.LABEL_FG if not exclude else "#94A3B8",
                bg=self.CARD, anchor="w",
            ).pack(side="left")

            if exclude:
                tk.Label(
                    row, text="자동 제외",
                    font=("Malgun Gothic", 8),
                    fg="#EF4444", bg="#FEF2F2",
                    padx=5, pady=1,
                ).pack(side="left", padx=(8, 0))

        self._update_count()
        if excluded_count:
            self._set_status(f"규칙에 따라 {excluded_count}개 시트가 자동으로 제외되었습니다.")
        else:
            self._set_status("시트를 선택 후 변환 버튼을 누르세요.")

    # ── 규칙 재적용 ──────────────────────────────
    def _reapply_rules(self):
        if not self.sheet_vars:
            return
        rules = self.config_data.get("rules", [])
        excluded = sum(
            1 for name, var in self.sheet_vars.items()
            if should_exclude(name, rules) or not var.set(not should_exclude(name, rules))
        )
        # 위 표현식이 복잡하므로 명시적으로 처리
        excluded = 0
        for name, var in self.sheet_vars.items():
            ex = should_exclude(name, rules)
            var.set(not ex)
            if ex:
                excluded += 1
        self._update_count()
        if excluded:
            self._set_status(f"규칙 재적용 완료 — {excluded}개 시트 자동 제외")
        else:
            self._set_status("재적용 완료. 일치하는 규칙이 없습니다.")

    def _update_count(self):
        total = len(self.sheet_vars)
        checked = sum(1 for v in self.sheet_vars.values() if v.get())
        self.sheet_count_label.config(text=f"선택: {checked} / {total}")

    def _select_all(self):
        for v in self.sheet_vars.values():
            v.set(True)
        self._update_count()

    def _deselect_all(self):
        for v in self.sheet_vars.values():
            v.set(False)
        self._update_count()

    # ── 상태 / 진행 ──────────────────────────────
    def _set_status(self, msg):
        self.status_label.config(text=msg)
        self.update_idletasks()

    def _progress(self, value, msg=""):
        self.progress_var.set(value)
        if msg:
            self._set_status(msg)
        self.update_idletasks()

    # ── 변환 ─────────────────────────────────────
    def _start_convert(self):
        excel = self.excel_path.get()
        output = self.output_path.get()
        selected = [n for n, v in self.sheet_vars.items() if v.get()]

        if not excel:
            messagebox.showwarning("경고", "Excel 파일을 선택해주세요.")
            return
        if not output:
            messagebox.showwarning("경고", "저장 위치를 선택해주세요.")
            return
        if not selected:
            messagebox.showwarning("경고", "최소 1개 이상의 시트를 선택해주세요.")
            return

        self.convert_btn.config(state="disabled")
        self._progress(10, "변환 준비 중...")

        def worker():
            try:
                convert_sheets_to_pdf(excel, output, selected, self._progress)
                self.after(0, self._done, output)
            except Exception as e:
                self.after(0, self._error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, output):
        self.convert_btn.config(state="normal")
        self._progress(100, f"저장 완료: {os.path.basename(output)}")
        if messagebox.askyesno("완료", f"PDF 변환이 완료되었습니다!\n\n{output}\n\n파일을 열어보시겠습니까?"):
            os.startfile(output)

    def _error(self, msg):
        self.convert_btn.config(state="normal")
        self._progress(0, "변환 실패.")
        messagebox.showerror(
            "변환 오류",
            f"변환 중 오류가 발생했습니다.\n\n{msg}\n\n"
            "Microsoft Excel이 설치되어 있는지 확인해주세요.",
        )


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
def main():
    missing = check_dependencies()
    if missing:
        root = tk.Tk()
        root.withdraw()
        pkg_list = "\n".join(f"  pip install {p}" for p in missing)
        messagebox.showerror(
            "필수 패키지 누락",
            f"다음 패키지를 먼저 설치해주세요:\n\n{pkg_list}\n\n설치 후 다시 실행하세요.",
        )
        root.destroy()
        sys.exit(1)

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

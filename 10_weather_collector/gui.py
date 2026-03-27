"""
건설현장 기상데이터 수집 & 작업불가일 산정 — GUI
실행: python gui.py
"""

import math
import re
import threading
from datetime import date

import customtkinter as ctk
from tkinter import messagebox

from station_mapper import ASOS_STATIONS
from kma_client import fetch_daily_weather, parse_weather, validate_station
from storage import init_db, upsert_weather
from analyzer import summarize

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ── 상수 ──────────────────────────────────────────────────────
WORK_PRESETS = {
    "토공사":          ["is_rain_day", "is_snow_day", "is_freeze_day", "is_cold_day"],
    "철근콘크리트공사": ["is_rain_day", "is_heat_day", "is_cold_day", "is_freeze_day", "is_wind_day"],
    "타워크레인작업":   ["is_wind_crane", "is_wind_day", "fog_yn"],
    "도장·방수공사":   ["is_rain_day", "rain_yn", "is_no_sunshine", "is_cold_day", "is_freeze_day"],
    "강구조물공사":    ["is_rain_day", "is_wind_day", "is_cold_day", "is_freeze_day", "is_heat_day"],
    "포장공사":        ["is_rain_day", "is_snow_day", "is_freeze_day", "is_cold_day", "is_heat_day"],
    "직접 입력":       [],
}

ALL_FLAGS = [
    ("is_rain_day",      "우천 (10mm 이상)"),
    ("is_wind_day",      "강풍 (14m/s 이상)"),
    ("is_wind_crane",    "크레인 제한 (순간 10m/s)"),
    ("is_snow_day",      "적설 (1cm 이상)"),
    ("is_heat_day",      "폭염 (35℃ 이상)"),
    ("is_cold_day",      "한파 (-10℃ 이하)"),
    ("is_no_sunshine",   "일조 부족 (2시간 미만)"),
    ("is_freeze_day",    "지면 동결 (0℃ 이하)"),
    ("is_high_evap_day", "증발 과다 (10mm 이상)"),
    ("rain_yn",          "강수 유무 (소량 포함)"),
    ("snow_yn",          "강설 유무"),
    ("fog_yn",           "안개"),
]


def _haversine(lat1, lon1, lat2, lon2) -> float:
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2
    return 6371 * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


# ── 공종 추가 다이얼로그 ────────────────────────────────────────
class WorkDialog(ctk.CTkToplevel):
    def __init__(self, parent, site_start: str, site_end: str, on_save):
        super().__init__(parent)
        self.title("공종 추가")
        self.geometry("480x580")
        self.resizable(False, False)
        self.grab_set()

        self.site_start = site_start
        self.site_end = site_end
        self.on_save = on_save
        self.flag_vars: dict[str, ctk.BooleanVar] = {}

        self._build()
        self.after(50, self.lift)

    def _build(self):
        # 공종명
        ctk.CTkLabel(self, text="공종명 *", anchor="w").pack(fill="x", padx=20, pady=(20, 2))
        self.var_name = ctk.StringVar()
        self.entry_name = ctk.CTkEntry(self, textvariable=self.var_name, placeholder_text="예: 토공사")
        self.entry_name.pack(fill="x", padx=20, pady=(0, 12))

        # 프리셋
        ctk.CTkLabel(self, text="프리셋", anchor="w").pack(fill="x", padx=20, pady=(0, 2))
        preset_names = list(WORK_PRESETS.keys())
        self.var_preset = ctk.StringVar(value=preset_names[0])
        self.var_name.set(preset_names[0])
        ctk.CTkOptionMenu(
            self, values=preset_names, variable=self.var_preset,
            command=self._on_preset_change,
        ).pack(fill="x", padx=20, pady=(0, 12))

        # 작업 기간
        date_row = ctk.CTkFrame(self, fg_color="transparent")
        date_row.pack(fill="x", padx=20, pady=(0, 12))
        date_row.columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(date_row, text="시작일").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkLabel(date_row, text="종료일").grid(row=0, column=1, sticky="w")
        self.var_ws = ctk.StringVar(value=self.site_start)
        self.var_we = ctk.StringVar(value=self.site_end)
        ctk.CTkEntry(date_row, textvariable=self.var_ws).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkEntry(date_row, textvariable=self.var_we).grid(row=1, column=1, sticky="ew")

        # 플래그
        ctk.CTkLabel(self, text="작업불가일 판정 플래그 *", anchor="w").pack(fill="x", padx=20, pady=(0, 2))
        flag_scroll = ctk.CTkScrollableFrame(self, height=170)
        flag_scroll.pack(fill="x", padx=20, pady=(0, 12))

        default_flags = WORK_PRESETS[self.var_preset.get()]
        for flag_id, label in ALL_FLAGS:
            var = ctk.BooleanVar(value=flag_id in default_flags)
            self.flag_vars[flag_id] = var
            ctk.CTkCheckBox(flag_scroll, text=label, variable=var).pack(anchor="w", pady=2)

        # 버튼
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkButton(btn_row, text="취소", fg_color="gray", hover_color="dimgray",
                      command=self.destroy).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(btn_row, text="추가", command=self._save).pack(side="left", expand=True, fill="x")

    def _on_preset_change(self, preset: str):
        if preset != "직접 입력":
            self.var_name.set(preset)
        flags = WORK_PRESETS.get(preset, [])
        for flag_id, var in self.flag_vars.items():
            var.set(flag_id in flags)

    def _save(self):
        from datetime import datetime
        name = self.var_name.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "공종명을 입력해 주세요.", parent=self)
            return

        try:
            datetime.strptime(self.var_ws.get().strip(), "%Y-%m-%d")
            datetime.strptime(self.var_we.get().strip(), "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("날짜 오류", "날짜를 YYYY-MM-DD 형식으로 입력해 주세요.", parent=self)
            return

        flags = [fid for fid, var in self.flag_vars.items() if var.get()]
        if not flags:
            messagebox.showwarning("플래그 오류", "플래그를 1개 이상 선택해 주세요.", parent=self)
            return

        self.on_save({"name": name, "start": self.var_ws.get().strip(),
                      "end": self.var_we.get().strip(), "flags": flags})
        self.destroy()


# ── 메인 앱 ────────────────────────────────────────────────────
class App(ctk.CTk):
    STEPS = ["현장 정보", "관측소 선택", "수집 기간", "공종 설정", "실행"]

    def __init__(self):
        super().__init__()
        self.title("건설현장 기상데이터 수집 & 작업불가일 산정")
        self.geometry("860x640")
        self.resizable(False, False)

        self.current_step = 0
        self.selected_station: dict | None = None
        self.works: list[dict] = []

        self._build_layout()
        self._show_step(0)

    # ── 레이아웃 구성 ───────────────────────────────────────────
    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 사이드바
        sidebar = ctk.CTkFrame(self, width=170, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        ctk.CTkLabel(sidebar, text="기상데이터\n수집기",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(28, 20))

        self._step_btns: list[ctk.CTkButton] = []
        for i, label in enumerate(self.STEPS):
            btn = ctk.CTkButton(
                sidebar, text=f"  {i + 1}.  {label}",
                width=150, height=34, anchor="w",
                fg_color="transparent",
                text_color=("gray50", "gray50"),
                hover_color=("gray85", "gray25"),
                command=lambda i=i: self._show_step(i),
            )
            btn.pack(padx=10, pady=2)
            self._step_btns.append(btn)

        # 메인 영역
        self.main_area = ctk.CTkFrame(self)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        self.main_area.grid_rowconfigure(0, weight=1)
        self.main_area.grid_columnconfigure(0, weight=1)

        # 스텝 프레임 (모두 생성 후 포개기)
        self._frames = [
            self._build_step1(),
            self._build_step2(),
            self._build_step3(),
            self._build_step4(),
            self._build_step5(),
        ]
        for f in self._frames:
            f.grid(row=0, column=0, sticky="nsew")

        # 하단 네비게이션
        nav = ctk.CTkFrame(self.main_area, fg_color="transparent", height=48)
        nav.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        nav.grid_columnconfigure(1, weight=1)

        self.btn_prev = ctk.CTkButton(nav, text="← 이전", width=100, command=self._prev_step)
        self.btn_prev.grid(row=0, column=0)
        self.btn_next = ctk.CTkButton(nav, text="다음 →", width=100, command=self._next_step)
        self.btn_next.grid(row=0, column=2)

    # ── 스텝 전환 ───────────────────────────────────────────────
    def _show_step(self, step: int):
        self.current_step = step
        self._frames[step].tkraise()

        for i, btn in enumerate(self._step_btns):
            if i == step:
                btn.configure(fg_color=("gray75", "gray30"),
                               text_color=("black", "white"))
            else:
                btn.configure(fg_color="transparent",
                               text_color=("gray50", "gray50"))

        self.btn_prev.configure(state="normal" if step > 0 else "disabled")
        self.btn_next.configure(state="normal" if step < len(self.STEPS) - 1 else "disabled")

    def _prev_step(self):
        if self.current_step > 0:
            self._show_step(self.current_step - 1)

    def _next_step(self):
        if not self._validate_step():
            return
        if self.current_step < len(self.STEPS) - 1:
            next_step = self.current_step + 1
            if next_step == 4:
                self._refresh_summary()
            self._show_step(next_step)

    def _validate_step(self) -> bool:
        from datetime import datetime
        step = self.current_step
        if step == 0:
            if not self.var_site_name.get().strip():
                messagebox.showwarning("입력 오류", "현장명을 입력해 주세요.")
                return False
        elif step == 1:
            if not self.selected_station:
                messagebox.showwarning("관측소 미선택", "관측소를 선택하고 유효성을 확인해 주세요.")
                return False
        elif step == 2:
            try:
                s = datetime.strptime(self.var_start.get().strip(), "%Y-%m-%d")
                e = datetime.strptime(self.var_end.get().strip(), "%Y-%m-%d")
                if s > e:
                    messagebox.showwarning("날짜 오류", "종료일이 시작일보다 이전입니다.")
                    return False
            except ValueError:
                messagebox.showwarning("날짜 오류", "날짜를 YYYY-MM-DD 형식으로 입력해 주세요.")
                return False
        return True

    # ── Step 1: 현장 정보 ───────────────────────────────────────
    def _build_step1(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="현장 정보 입력",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=24, pady=(24, 20))

        form = ctk.CTkFrame(frame, fg_color="transparent")
        form.pack(anchor="w", padx=24, fill="x")

        ctk.CTkLabel(form, text="현장명 *", anchor="w").pack(fill="x", pady=(0, 2))
        self.var_site_name = ctk.StringVar()
        self.entry_site_name = ctk.CTkEntry(
            form, textvariable=self.var_site_name,
            placeholder_text="예: 용인 파장동 현장", width=340)
        self.entry_site_name.pack(anchor="w", pady=(0, 14))

        ctk.CTkLabel(form, text="현장 ID", anchor="w").pack(fill="x", pady=(0, 2))
        self.var_site_id = ctk.StringVar()
        ctk.CTkEntry(form, textvariable=self.var_site_id,
                     placeholder_text="비워두면 현장명에서 자동 생성", width=340).pack(anchor="w", pady=(0, 14))

        coord = ctk.CTkFrame(form, fg_color="transparent")
        coord.pack(anchor="w", pady=(0, 8))
        coord.columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(coord, text="위도 (선택)").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ctk.CTkLabel(coord, text="경도 (선택)").grid(row=0, column=1, sticky="w")
        self.var_lat = ctk.StringVar()
        self.var_lon = ctk.StringVar()
        ctk.CTkEntry(coord, textvariable=self.var_lat, placeholder_text="예: 37.2723", width=160).grid(row=1, column=0, padx=(0, 12))
        ctk.CTkEntry(coord, textvariable=self.var_lon, placeholder_text="예: 127.4842", width=160).grid(row=1, column=1)

        ctk.CTkLabel(frame,
                     text="※ 위도/경도를 입력하면 관측소 선택 시 가까운 관측소를 추천해 드립니다.",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(anchor="w", padx=24, pady=(12, 0))
        return frame

    # ── Step 2: 관측소 선택 ─────────────────────────────────────
    def _build_step2(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="관측소 선택",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=24, pady=(24, 10))

        # 검색창
        search_row = ctk.CTkFrame(frame, fg_color="transparent")
        search_row.pack(fill="x", padx=24, pady=(0, 4))

        self.var_search = ctk.StringVar()
        entry_search = ctk.CTkEntry(search_row, textvariable=self.var_search,
                                    placeholder_text="관측소 이름 검색 (예: 수원, 서울)", width=260)
        entry_search.pack(side="left", padx=(0, 8))
        entry_search.bind("<Return>", lambda e: self._search_stations())

        ctk.CTkButton(search_row, text="검색", width=72,
                      command=self._search_stations).pack(side="left", padx=(0, 8))
        ctk.CTkButton(search_row, text="좌표 기반 추천", width=110,
                      fg_color="gray", hover_color="dimgray",
                      command=self._show_nearest).pack(side="left")

        # 상태 메시지 (검증 진행 상황)
        self.lbl_search_status = ctk.CTkLabel(frame, text="", text_color="gray",
                                               font=ctk.CTkFont(size=11), anchor="w")
        self.lbl_search_status.pack(fill="x", padx=26, pady=(0, 4))

        # 결과 목록
        self._result_frame = ctk.CTkScrollableFrame(frame, label_text="검색 결과")
        self._result_frame.pack(fill="both", expand=True, padx=24, pady=(0, 8))
        self._radio_var = ctk.StringVar()

        # 선택된 관측소 표시
        self.lbl_selected = ctk.CTkLabel(frame, text="선택된 관측소: 없음", anchor="w")
        self.lbl_selected.pack(fill="x", padx=24, pady=(0, 6))

        return frame

    def _search_stations(self):
        keyword = self.var_search.get().strip()
        if not keyword:
            return
        candidates = [s for s in ASOS_STATIONS if keyword in s["name"]]
        if not candidates:
            self._clear_results()
            self.lbl_search_status.configure(text="검색 결과가 없습니다.")
            return
        self._validate_and_show(candidates)

    def _show_nearest(self):
        try:
            lat = float(self.var_lat.get())
            lon = float(self.var_lon.get())
        except ValueError:
            messagebox.showwarning("좌표 오류", "현장 정보 탭에서 위도/경도를 먼저 입력해 주세요.")
            return
        candidates = sorted(ASOS_STATIONS,
                             key=lambda s: _haversine(lat, lon, s["lat"], s["lon"]))[:25]
        self._validate_and_show(candidates, lat, lon)

    def _validate_and_show(self, candidates: list[dict], lat=None, lon=None):
        """후보 관측소를 병렬로 검증한 뒤 유효한 것만 표시.
        유효 결과가 없으면 후보 좌표(또는 현장 좌표) 기준 인근 관측소로 자동 대체."""
        from concurrent.futures import ThreadPoolExecutor, as_completed

        self._clear_results()
        total = len(candidates)
        self.lbl_search_status.configure(text=f"데이터 제공 관측소 확인 중... (0 / {total})")
        self.selected_station = None
        self.lbl_selected.configure(text="선택된 관측소: 없음")

        checked = [0]

        def worker():
            valid = []
            with ThreadPoolExecutor(max_workers=6) as ex:
                futures = {ex.submit(validate_station, s["code"]): s for s in candidates}
                for future in as_completed(futures):
                    s = futures[future]
                    checked[0] += 1
                    self.after(0, lambda n=checked[0]: self.lbl_search_status.configure(
                        text=f"데이터 제공 관측소 확인 중... ({n} / {total})"
                    ))
                    if future.result():
                        valid.append(s)

            if valid:
                ref_lat, ref_lon = lat, lon
                if ref_lat is not None:
                    valid.sort(key=lambda s: _haversine(ref_lat, ref_lon, s["lat"], s["lon"]))
                self.after(0, lambda: self._show_results(valid, ref_lat, ref_lon, total))
            else:
                # 유효 결과 없음 → 기준 좌표로 인근 재검색
                ref_lat = lat
                ref_lon = lon
                # 현장 좌표 없으면 후보들의 평균 좌표 사용
                if ref_lat is None and candidates:
                    ref_lat = sum(s["lat"] for s in candidates) / len(candidates)
                    ref_lon = sum(s["lon"] for s in candidates) / len(candidates)

                if ref_lat is None:
                    self.after(0, lambda: self._show_results([], None, None, total))
                    return

                self.after(0, lambda: self.lbl_search_status.configure(
                    text="해당 지역 관측소 데이터 없음 → 인근 관측소 검색 중..."
                ))
                nearby = sorted(ASOS_STATIONS,
                                key=lambda s: _haversine(ref_lat, ref_lon, s["lat"], s["lon"]))[:30]
                # 후보 제외 후 재검증
                candidate_codes = {s["code"] for s in candidates}
                nearby = [s for s in nearby if s["code"] not in candidate_codes]

                valid2 = []
                with ThreadPoolExecutor(max_workers=6) as ex:
                    futures2 = {ex.submit(validate_station, s["code"]): s for s in nearby}
                    for future in as_completed(futures2):
                        s = futures2[future]
                        if future.result():
                            valid2.append(s)

                valid2.sort(key=lambda s: _haversine(ref_lat, ref_lon, s["lat"], s["lon"]))
                self.after(0, lambda: self._show_results(
                    valid2, ref_lat, ref_lon, total, fallback=True
                ))

        threading.Thread(target=worker, daemon=True).start()

    def _clear_results(self):
        for w in self._result_frame.winfo_children():
            w.destroy()
        self._radio_var.set("")

    def _show_results(self, results: list[dict], lat=None, lon=None, total=0, fallback=False):
        self._clear_results()
        found = len(results)

        if fallback:
            self.lbl_search_status.configure(
                text=f"검색 지역에 데이터 제공 관측소가 없어 인근 관측소 {found}개를 표시합니다."
            )
        else:
            self.lbl_search_status.configure(
                text=f"총 {total}개 중 데이터 제공 관측소 {found}개" if total else f"{found}개 관측소"
            )

        if not results:
            ctk.CTkLabel(self._result_frame, text="인근에서도 관측소를 찾을 수 없습니다.").pack()
            return

        for s in results:
            dist_str = ""
            if lat is not None and lon is not None:
                d = _haversine(lat, lon, s["lat"], s["lon"])
                dist_str = f"  {d:.1f}km"
            label = f"[{s['code']:>4s}] {s['name']:<14s} ({s['lat']:.4f}, {s['lon']:.4f}){dist_str}"
            ctk.CTkRadioButton(
                self._result_frame, text=label,
                variable=self._radio_var, value=s["code"],
                font=ctk.CTkFont(family="Consolas", size=12),
                command=lambda s=s: self._on_select(s),
            ).pack(anchor="w", pady=2)

    def _on_select(self, station: dict):
        self.selected_station = station
        self.lbl_selected.configure(text=f"✓  [{station['code']}] {station['name']} 선택됨")

    # ── Step 3: 수집 기간 ───────────────────────────────────────
    def _build_step3(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self.main_area, fg_color="transparent")

        ctk.CTkLabel(frame, text="수집 기간 설정",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=24, pady=(24, 20))

        form = ctk.CTkFrame(frame, fg_color="transparent")
        form.pack(anchor="w", padx=24)

        ctk.CTkLabel(form, text="시작일 (YYYY-MM-DD)", anchor="w").pack(fill="x", pady=(0, 2))
        self.var_start = ctk.StringVar(value=f"{date.today().year - 1}-01-01")
        ctk.CTkEntry(form, textvariable=self.var_start, width=220).pack(anchor="w", pady=(0, 16))

        ctk.CTkLabel(form, text="종료일 (YYYY-MM-DD)", anchor="w").pack(fill="x", pady=(0, 2))
        self.var_end = ctk.StringVar(value=f"{date.today().year}-01-01")
        ctk.CTkEntry(form, textvariable=self.var_end, width=220).pack(anchor="w")
        return frame

    # ── Step 4: 공종 설정 ───────────────────────────────────────
    def _build_step4(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(24, 8))
        ctk.CTkLabel(header, text="공종 설정",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(header, text="+ 공종 추가", width=110,
                      command=self._open_work_dialog).pack(side="right")

        self._work_list = ctk.CTkScrollableFrame(frame, label_text="추가된 공종")
        self._work_list.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        ctk.CTkLabel(frame, text="※ 공종을 추가하지 않으면 기상 데이터만 수집합니다.",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(anchor="w", padx=24)
        return frame

    def _open_work_dialog(self):
        s = self.var_start.get().strip() or f"{date.today().year - 1}-01-01"
        e = self.var_end.get().strip() or f"{date.today().year}-01-01"
        WorkDialog(self, s, e, on_save=self._on_work_saved)

    def _on_work_saved(self, work: dict):
        self.works.append(work)
        self._render_works()

    def _render_works(self):
        for w in self._work_list.winfo_children():
            w.destroy()
        for i, work in enumerate(self.works):
            card = ctk.CTkFrame(self._work_list)
            card.pack(fill="x", pady=3)
            card.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                card,
                text=f"{work['name']}   {work['start']} ~ {work['end']}\n플래그: {', '.join(work['flags'])}",
                anchor="w", justify="left", font=ctk.CTkFont(size=12),
            ).grid(row=0, column=0, sticky="w", padx=10, pady=6)

            ctk.CTkButton(
                card, text="삭제", width=60,
                fg_color="#c0392b", hover_color="#96281b",
                command=lambda i=i: self._delete_work(i),
            ).grid(row=0, column=1, padx=8)

    def _delete_work(self, idx: int):
        self.works.pop(idx)
        self._render_works()

    # ── Step 5: 실행 ────────────────────────────────────────────
    def _build_step5(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(frame, text="설정 확인 & 실행",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=24, pady=(24, 8))

        self.summary_box = ctk.CTkTextbox(frame, height=130, state="disabled")
        self.summary_box.pack(fill="x", padx=24, pady=(0, 10))

        self.progress_bar = ctk.CTkProgressBar(frame)
        self.progress_bar.pack(fill="x", padx=24, pady=(0, 8))
        self.progress_bar.set(0)

        self.log_box = ctk.CTkTextbox(frame, state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        self.btn_run = ctk.CTkButton(frame, text="▶  데이터 수집 시작", height=40,
                                      command=self._run)
        self.btn_run.pack(padx=24, pady=(0, 8), fill="x")
        return frame

    def _refresh_summary(self):
        st = self.selected_station
        st_str = f"[{st['code']}] {st['name']}" if st else "미선택"
        lines = [
            f"현장명    : {self.var_site_name.get().strip() or '미입력'}",
            f"관측소    : {st_str}",
            f"수집 기간 : {self.var_start.get().strip()} ~ {self.var_end.get().strip()}",
            f"공종 수   : {len(self.works)}개",
        ]
        for w in self.works:
            lines.append(f"  - {w['name']}: {w['start']} ~ {w['end']}")

        self.summary_box.configure(state="normal")
        self.summary_box.delete("1.0", "end")
        self.summary_box.insert("end", "\n".join(lines))
        self.summary_box.configure(state="disabled")

    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _run(self):
        if not self.selected_station:
            messagebox.showwarning("오류", "관측소를 선택해 주세요.")
            return

        self.btn_run.configure(state="disabled", text="수집 중...")
        self.progress_bar.set(0)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        def worker():
            try:
                self.after(0, lambda: self._log("DB 초기화..."))
                init_db()

                site_name = self.var_site_name.get().strip()
                site_id = (self.var_site_id.get().strip()
                           or re.sub(r"[^a-zA-Z0-9가-힣]", "", site_name)[:10]
                           or "SITE001")
                code = self.selected_station["code"]
                start_fmt = self.var_start.get().strip().replace("-", "")
                end_fmt = self.var_end.get().strip().replace("-", "")

                self.after(0, lambda: self._log(f"[{code}] {self.selected_station['name']} 데이터 수집 중..."))
                self.after(0, lambda: self.progress_bar.configure(mode="indeterminate"))
                self.after(0, lambda: self.progress_bar.start())

                raw = fetch_daily_weather(code, start_fmt, end_fmt)

                self.after(0, lambda: self.progress_bar.stop())
                self.after(0, lambda: self.progress_bar.configure(mode="determinate"))

                if not raw:
                    self.after(0, lambda: self._log("※ 수집된 데이터가 없습니다."))
                    return

                parsed = [parse_weather(r, site_id) for r in raw]
                upsert_weather(parsed)
                self.after(0, lambda: self._log(f"✓ {len(parsed)}일치 데이터 수집 완료"))
                self.after(0, lambda: self.progress_bar.set(0.7))

                if self.works:
                    self.after(0, lambda: self._log("작업불가일 산정 중..."))
                    try:
                        lat_v = float(self.var_lat.get())
                        lon_v = float(self.var_lon.get())
                    except ValueError:
                        lat_v = lon_v = 0.0

                    summarize({
                        "id": site_id, "name": site_name,
                        "lat": lat_v, "lon": lon_v,
                        "start": self.var_start.get().strip(),
                        "end": self.var_end.get().strip(),
                        "works": self.works,
                    })
                    self.after(0, lambda: self._log("✓ 엑셀 파일 저장 완료"))

                self.after(0, lambda: self.progress_bar.set(1.0))
                self.after(0, lambda: self._log("모든 작업 완료!"))
                self.after(0, lambda: messagebox.showinfo("완료", "데이터 수집 및 분석이 완료되었습니다."))

            except Exception as exc:
                self.after(0, lambda: self._log(f"[ERROR] {exc}"))
                self.after(0, lambda: messagebox.showerror("오류", str(exc)))
            finally:
                self.after(0, lambda: self.btn_run.configure(state="normal", text="▶  데이터 수집 시작"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()

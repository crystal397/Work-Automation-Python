"""
==================================================
  권역 분석 + 점묘도 시각화 GUI 앱
==================================================
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import webbrowser
from math import radians, cos, sin, asin, sqrt

# ── 런타임 의존 패키지 지연 임포트 ──────────────────────────
def _import_or_warn(module_name):
    try:
        return __import__(module_name)
    except ImportError:
        messagebox.showerror(
            "패키지 누락",
            f"'{module_name}' 패키지가 설치되어 있지 않습니다.\n"
            f"터미널에서 실행: pip install {module_name}",
        )
        return None


# ─────────────────────────────────────────────────────────────
#  색상 팔레트
# ─────────────────────────────────────────────────────────────
POINT_COLORS_CYCLE = [
    "#E63946", "#2196F3", "#4CAF50", "#FF9800",
    "#9C27B0", "#00BCD4", "#795548", "#607D8B",
]

ZONE_DOT_COLORS_CYCLE = [
    "#FF1744", "#FF9100", "#00B0FF", "#76FF03",
    "#E040FB", "#00E5FF", "#FFEA00", "#FF6D00",
]

ZONE_STYLE_CYCLE = [
    {"weight": 2.5, "dash": None,  "fill_opacity": 0.07},
    {"weight": 1.8, "dash": "6 4", "fill_opacity": 0.04},
    {"weight": 1.2, "dash": "3 6", "fill_opacity": 0.02},
    {"weight": 1.0, "dash": "2 8", "fill_opacity": 0.01},
    {"weight": 0.8, "dash": "1 8", "fill_opacity": 0.01},
]

TILE = "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png"
TILE_ATTR = "&copy; <a href='https://carto.com/'>CARTO</a>"


# ─────────────────────────────────────────────────────────────
#  핵심 분석 함수들
# ─────────────────────────────────────────────────────────────
def haversine(lat1, lng1, lat2, lng2) -> float:
    R = 6_371_000
    lat1, lng1, lat2, lng2 = map(radians, [lat1, lng1, lat2, lng2])
    dlat, dlng = lat2 - lat1, lng2 - lng1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlng / 2) ** 2
    return 2 * R * asin(sqrt(a))


def load_excel(path: str):
    pd = _import_or_warn("pandas")
    if pd is None:
        return None
    df = pd.read_excel(path, header=0)
    df = df.rename(columns={df.columns[1]: "lat", df.columns[2]: "lng"})
    df = df.dropna(subset=["lat", "lng"])
    df["lat"] = pd.to_numeric(df["lat"], errors="coerce")
    df["lng"] = pd.to_numeric(df["lng"], errors="coerce")
    df = df.dropna(subset=["lat", "lng"]).reset_index(drop=True)
    return df


def classify_zones(df, reference_points: list, radius_map: dict):
    for rp in reference_points:
        name = rp["name"]
        distances = df.apply(
            lambda row: haversine(row["lat"], row["lng"], rp["lat"], rp["lng"]),
            axis=1,
        )

        sorted_radii = sorted(radius_map.items())   # [(1, 250), (2, 500), ...]

        def to_zone(d):
            for zone_no, r in sorted_radii:
                if d <= r:
                    return f"{zone_no}권역"
            return "해당없음"

        df[f"권역_{name}"]    = distances.apply(to_zone)
        df[f"거리_{name}(m)"] = distances.round(1)
    return df


def draw_map(df, reference_points: list, radius_map: dict):
    folium = _import_or_warn("folium")
    if folium is None:
        return None

    center_lat = sum(rp["lat"] for rp in reference_points) / len(reference_points)
    center_lng = sum(rp["lng"] for rp in reference_points) / len(reference_points)
    m = folium.Map(location=[center_lat, center_lng], zoom_start=13, tiles=TILE, attr=TILE_ATTR)

    sorted_radii = sorted(radius_map.items())

    # 1. 권역 원
    for i, rp in enumerate(reference_points):
        color = POINT_COLORS_CYCLE[i % len(POINT_COLORS_CYCLE)]
        for j, (zone_no, radius) in enumerate(sorted_radii):
            s = ZONE_STYLE_CYCLE[j % len(ZONE_STYLE_CYCLE)]
            folium.Circle(
                location=[rp["lat"], rp["lng"]],
                radius=radius,
                color=color,
                weight=s["weight"],
                dash_array=s["dash"],
                fill=True,
                fill_color=color,
                fill_opacity=s["fill_opacity"],
                tooltip=f"{rp['name']} {zone_no}권역 ({radius}m)",
            ).add_to(m)

    # 2. 기준점 마커
    for i, rp in enumerate(reference_points):
        color = POINT_COLORS_CYCLE[i % len(POINT_COLORS_CYCLE)]
        folium.Marker(
            location=[rp["lat"], rp["lng"]],
            tooltip=rp["name"],
            icon=folium.Icon(color="white", icon_color=color, icon="star", prefix="fa"),
        ).add_to(m)
        label_html = (
            f'<div style="font-size:12px;font-weight:700;color:{color};'
            f'white-space:nowrap;text-shadow:1px 1px 2px #fff,-1px -1px 2px #fff;'
            f'margin-top:-28px;margin-left:10px;">&#9733; {rp["name"]}</div>'
        )
        folium.Marker(
            location=[rp["lat"], rp["lng"]],
            icon=folium.DivIcon(html=label_html, icon_size=(150, 20)),
        ).add_to(m)

    # 3. 점묘도 레이어
    from folium import FeatureGroup
    for rp in reference_points:
        zone_col = f"권역_{rp['name']}"
        dist_col = f"거리_{rp['name']}(m)"

        for j, (zone_no, radius) in enumerate(sorted_radii):
            zone_key  = f"{zone_no}권역"
            dot_color = ZONE_DOT_COLORS_CYCLE[j % len(ZONE_DOT_COLORS_CYCLE)]
            fg        = FeatureGroup(name=f"{rp['name']} · {zone_key} ({radius}m)", show=True)
            subset    = df[df[zone_col] == zone_key]

            for _, row in subset.iterrows():
                tip = f"{rp['name']} | {zone_key} | 거리: {row[dist_col]}m"
                folium.CircleMarker(
                    location=[row["lat"], row["lng"]],
                    radius=3,
                    color=dot_color,
                    fill=True,
                    fill_color=dot_color,
                    fill_opacity=0.75,
                    weight=0,
                    tooltip=tip,
                ).add_to(fg)
            fg.add_to(m)

    folium.LayerControl(collapsed=False).add_to(m)

    # 4. 범례 (해당없음 제외)
    legend_items = '<b style="font-size:14px;">&#128205; 권역 범례</b><br>'
    for j, (zone_no, radius) in enumerate(sorted_radii):
        dot_color = ZONE_DOT_COLORS_CYCLE[j % len(ZONE_DOT_COLORS_CYCLE)]
        legend_items += (
            f'<span style="color:{dot_color};">&#9679;</span> '
            f'{zone_no}권역 ({radius}m 이내)<br>'
        )
    legend_items += '<hr style="margin:6px 0">'
    for i, rp in enumerate(reference_points):
        c = POINT_COLORS_CYCLE[i % len(POINT_COLORS_CYCLE)]
        legend_items += f'<span style="color:{c};">&#9733;</span> {rp["name"]}<br>'

    legend_html = (
        '<div style="position:fixed;bottom:40px;left:40px;z-index:1000;'
        'background:white;border-radius:10px;padding:14px 18px;'
        'box-shadow:0 2px 10px rgba(0,0,0,0.25);'
        'font-family:sans-serif;font-size:13px;line-height:1.8;">'
        + legend_items + "</div>"
    )
    m.get_root().html.add_child(folium.Element(legend_html))
    return m


def save_excel(df, reference_points: list, out_dir: str):
    pd = _import_or_warn("pandas")
    if pd is None:
        return []
    paths = []
    for rp in reference_points:
        name     = rp["name"]
        zone_col = f"권역_{name}"
        dist_col = f"거리_{name}(m)"

        # 이 기준점 외 컬럼 제거
        drop_cols = []
        for other in reference_points:
            if other["name"] != name:
                drop_cols += [f"권역_{other['name']}", f"거리_{other['name']}(m)"]
        base_df = df.drop(columns=[c for c in drop_cols if c in df.columns])

        # 거리 오름차순 정렬
        base_df = base_df.sort_values(by=dist_col, ascending=True)

        safe_name = name.replace(" ", "_")
        out_path  = os.path.join(out_dir, f"zone_result_{safe_name}.xlsx")

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            base_df.to_excel(writer, sheet_name="전체데이터", index=False)

        paths.append(out_path)

    return paths


# ─────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("권역 분석 · 점묘도 시각화")
        self.resizable(True, True)
        self.minsize(780, 680)
        self.configure(bg="#F5F5F5")

        # 상태 변수
        self.excel_path   = tk.StringVar()
        self.output_dir   = tk.StringVar()
        self.ref_rows     = []   # 기준점 행 위젯 목록
        self.zone_rows    = []   # 권역 행 위젯 목록

        self._build_ui()
        self._add_ref_row()   # 기본 기준점 1개
        self._add_ref_row()
        self._add_ref_row()
        self._add_zone_row(250)   # 기본 3권역
        self._add_zone_row(500)
        self._add_zone_row(1000)

    # ── UI 빌드 ─────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 14, "pady": 6}

        # ── 상단 타이틀 ──
        title_frame = tk.Frame(self, bg="#1565C0")
        title_frame.pack(fill="x")
        tk.Label(
            title_frame,
            text="  📍 권역 분석 · 점묘도 시각화",
            font=("맑은 고딕", 15, "bold"),
            fg="white", bg="#1565C0",
            pady=10,
        ).pack(side="left")

        # ── 스크롤 가능 메인 프레임 ──
        canvas  = tk.Canvas(self, bg="#F5F5F5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg="#F5F5F5")
        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        body = self.scroll_frame

        # ── 섹션: 엑셀 파일 ──
        self._section(body, "① 엑셀 파일 선택")
        file_frame = tk.Frame(body, bg="#F5F5F5")
        file_frame.pack(fill="x", **pad)
        tk.Entry(file_frame, textvariable=self.excel_path, width=55,
                 font=("맑은 고딕", 10)).pack(side="left", padx=(0, 6))
        tk.Button(file_frame, text="파일 선택", command=self._pick_excel,
                  bg="#1565C0", fg="white", font=("맑은 고딕", 9, "bold"),
                  relief="flat", padx=10).pack(side="left")

        # ── 섹션: 출력 폴더 ──
        self._section(body, "② 결과 저장 폴더")
        out_frame = tk.Frame(body, bg="#F5F5F5")
        out_frame.pack(fill="x", **pad)
        tk.Entry(out_frame, textvariable=self.output_dir, width=55,
                 font=("맑은 고딕", 10)).pack(side="left", padx=(0, 6))
        tk.Button(out_frame, text="폴더 선택", command=self._pick_outdir,
                  bg="#1565C0", fg="white", font=("맑은 고딕", 9, "bold"),
                  relief="flat", padx=10).pack(side="left")

        # ── 섹션: 기준점 ──
        self._section(body, "③ 기준점 설정")
        ref_header = tk.Frame(body, bg="#E3F2FD")
        ref_header.pack(fill="x", padx=14, pady=(0, 2))
        for txt, w in [("이름", 12), ("위도", 16), ("경도", 16), ("", 5)]:
            tk.Label(ref_header, text=txt, bg="#E3F2FD",
                     font=("맑은 고딕", 9, "bold"), width=w, anchor="w").pack(side="left")

        self.ref_container = tk.Frame(body, bg="#F5F5F5")
        self.ref_container.pack(fill="x", padx=14)

        tk.Button(body, text="＋ 기준점 추가", command=self._add_ref_row,
                  bg="#0D47A1", fg="white", font=("맑은 고딕", 9),
                  relief="flat", padx=8, pady=3).pack(anchor="w", padx=14, pady=4)

        # ── 섹션: 권역 ──
        self._section(body, "④ 권역 설정")
        zone_header = tk.Frame(body, bg="#E8F5E9")
        zone_header.pack(fill="x", padx=14, pady=(0, 2))
        for txt, w in [("권역 번호", 10), ("반경 (m)", 12), ("", 5)]:
            tk.Label(zone_header, text=txt, bg="#E8F5E9",
                     font=("맑은 고딕", 9, "bold"), width=w, anchor="w").pack(side="left")

        self.zone_container = tk.Frame(body, bg="#F5F5F5")
        self.zone_container.pack(fill="x", padx=14)

        tk.Button(body, text="＋ 권역 추가", command=lambda: self._add_zone_row(),
                  bg="#1B5E20", fg="white", font=("맑은 고딕", 9),
                  relief="flat", padx=8, pady=3).pack(anchor="w", padx=14, pady=4)

        # ── 실행 버튼 ──
        tk.Button(
            body, text="▶  분석 실행", command=self._run,
            bg="#E53935", fg="white",
            font=("맑은 고딕", 13, "bold"),
            relief="flat", padx=20, pady=8,
        ).pack(pady=(16, 6))

        # ── 로그 창 ──
        self._section(body, "⑤ 실행 로그")
        self.log = scrolledtext.ScrolledText(
            body, height=10, font=("Consolas", 9),
            bg="#1E1E1E", fg="#D4D4D4", state="disabled",
            relief="flat",
        )
        self.log.pack(fill="x", padx=14, pady=(0, 20))

    def _section(self, parent, text):
        f = tk.Frame(parent, bg="#F5F5F5")
        f.pack(fill="x", padx=14, pady=(12, 2))
        tk.Label(f, text=text, font=("맑은 고딕", 11, "bold"),
                 fg="#1565C0", bg="#F5F5F5").pack(side="left")
        tk.Frame(f, bg="#BBDEFB", height=2).pack(side="left", fill="x", expand=True, padx=8)

    # ── 파일·폴더 선택 ──────────────────────────────────────
    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(path))

    def _pick_outdir(self):
        d = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if d:
            self.output_dir.set(d)

    # ── 기준점 행 관리 ───────────────────────────────────────
    def _add_ref_row(self, name="", lat="", lng=""):
        idx = len(self.ref_rows) + 1
        row_frame = tk.Frame(self.ref_container, bg="#FAFAFA",
                             relief="groove", bd=1)
        row_frame.pack(fill="x", pady=2)

        name_var = tk.StringVar(value=name or f"기준점 {chr(64+idx)}")
        lat_var  = tk.StringVar(value=lat)
        lng_var  = tk.StringVar(value=lng)

        for var, w in [(name_var, 12), (lat_var, 16), (lng_var, 16)]:
            tk.Entry(row_frame, textvariable=var, width=w,
                     font=("맑은 고딕", 10)).pack(side="left", padx=4, pady=3)

        color_badge = tk.Label(
            row_frame,
            bg=POINT_COLORS_CYCLE[(idx - 1) % len(POINT_COLORS_CYCLE)],
            width=3,
        )
        color_badge.pack(side="left", padx=4)

        del_btn = tk.Button(
            row_frame, text="✕", fg="red", bg="#FAFAFA",
            relief="flat", font=("맑은 고딕", 9),
            command=lambda f=row_frame, r=(name_var, lat_var, lng_var): self._del_ref_row(f, r),
        )
        del_btn.pack(side="left", padx=2)

        self.ref_rows.append((row_frame, name_var, lat_var, lng_var))

    def _del_ref_row(self, frame, row_vars):
        frame.destroy()
        self.ref_rows = [(f, n, la, lo) for (f, n, la, lo) in self.ref_rows
                         if f != frame]

    # ── 권역 행 관리 ────────────────────────────────────────
    def _add_zone_row(self, radius_default=None):
        idx = len(self.zone_rows) + 1
        row_frame = tk.Frame(self.zone_container, bg="#F9FBE7",
                             relief="groove", bd=1)
        row_frame.pack(fill="x", pady=2)

        radius_var = tk.StringVar(value=str(radius_default) if radius_default else "")
        tk.Label(row_frame, text=f"{idx}권역", width=10,
                 font=("맑은 고딕", 10, "bold"),
                 bg=ZONE_DOT_COLORS_CYCLE[(idx - 1) % len(ZONE_DOT_COLORS_CYCLE)],
                 fg="white").pack(side="left", padx=4, pady=3)
        tk.Entry(row_frame, textvariable=radius_var, width=12,
                 font=("맑은 고딕", 10)).pack(side="left", padx=4)
        tk.Label(row_frame, text="m", bg="#F9FBE7",
                 font=("맑은 고딕", 10)).pack(side="left")

        del_btn = tk.Button(
            row_frame, text="✕", fg="red", bg="#F9FBE7",
            relief="flat", font=("맑은 고딕", 9),
            command=lambda f=row_frame, rv=radius_var: self._del_zone_row(f, rv),
        )
        del_btn.pack(side="left", padx=6)

        self.zone_rows.append((row_frame, radius_var))

    def _del_zone_row(self, frame, _var):
        frame.destroy()
        self.zone_rows = [(f, v) for (f, v) in self.zone_rows if f != frame]

    # ── 로그 출력 ────────────────────────────────────────────
    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── 입력값 유효성 검사 ───────────────────────────────────
    def _validate(self):
        if not self.excel_path.get():
            messagebox.showerror("오류", "엑셀 파일을 선택하세요.")
            return False
        if not self.output_dir.get():
            messagebox.showerror("오류", "결과 저장 폴더를 선택하세요.")
            return False
        if not self.ref_rows:
            messagebox.showerror("오류", "기준점을 1개 이상 입력하세요.")
            return False
        for (_, n, la, lo) in self.ref_rows:
            if not n.get().strip():
                messagebox.showerror("오류", "기준점 이름이 비어 있습니다.")
                return False
            try:
                float(la.get()); float(lo.get())
            except ValueError:
                messagebox.showerror("오류", f"[{n.get()}] 위도·경도는 숫자여야 합니다.")
                return False
        if not self.zone_rows:
            messagebox.showerror("오류", "권역을 1개 이상 입력하세요.")
            return False
        for i, (_, rv) in enumerate(self.zone_rows):
            try:
                v = float(rv.get())
                if v <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("오류", f"{i+1}권역의 반경은 양수여야 합니다.")
                return False
        return True

    # ── 실행 ────────────────────────────────────────────────
    def _run(self):
        if not self._validate():
            return
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        try:
            # 입력값 수집
            reference_points = [
                {"name": n.get().strip(),
                 "lat":  float(la.get()),
                 "lng":  float(lo.get())}
                for (_, n, la, lo) in self.ref_rows
            ]
            radius_map = {
                i + 1: int(float(rv.get()))
                for i, (_, rv) in enumerate(self.zone_rows)
            }
            out_dir = self.output_dir.get()
            os.makedirs(out_dir, exist_ok=True)

            self._log("=" * 50)
            self._log("📂 엑셀 파일 로딩 중...")
            df = load_excel(self.excel_path.get())
            if df is None:
                return
            self._log(f"   ✅ {len(df):,}건 로드 완료")

            self._log("📐 권역 분류 중...")
            df = classify_zones(df, reference_points, radius_map)
            self._log("   ✅ 완료")

            self._log("🗺️  지도 생성 중...")
            m = draw_map(df, reference_points, radius_map)
            if m is None:
                return
            map_path = os.path.join(out_dir, "zone_map.html")
            m.save(map_path)
            self._log(f"   ✅ 지도 저장 → {map_path}")

            self._log("📊 엑셀 저장 중...")
            paths = save_excel(df, reference_points, out_dir)
            for p in paths:
                self._log(f"   ✅ {p}")

            self._log("=" * 50)
            self._log("🎉 분석 완료!")

            # 완료 후 지도 열기 여부 묻기
            self.after(
                200,
                lambda: self._ask_open_map(map_path),
            )

        except Exception as e:
            self._log(f"❌ 오류 발생: {e}")
            import traceback
            self._log(traceback.format_exc())

    def _ask_open_map(self, map_path):
        if messagebox.askyesno("완료", "분석이 완료되었습니다.\n지도를 브라우저에서 열까요?"):
            webbrowser.open(f"file:///{map_path.replace(os.sep, '/')}")


# ─────────────────────────────────────────────────────────────
#  엔트리포인트
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
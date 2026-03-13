"""
main.py
공기연장 간접비 보고서 자동화 - 메인 실행 파일

실행 방법:
    python main.py
"""
import os
import warnings
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

import win32com.client as win32

from excel_reader import (
    load_workbook,
    read_header_info,
    read_main_table_values,
    read_indirect_labor_table,
    read_severance_table,
    read_fourth_table_values,
    read_fifth_table_values,
    read_sixth_table_values,
)
from hwp_writer import (
    write_header_fields,
    write_main_table,
    write_indirect_labor_table,
    write_severance_table,
    write_fourth_table,
    write_fifth_table,
    write_sixth_table,
)

# 엑셀 Print Area 관련 경고 무시
warnings.filterwarnings("ignore", category=UserWarning)


# ─────────────────────────────────────────────
# 파일 선택 UI
# ─────────────────────────────────────────────

def select_files() -> tuple[str, str]:
    """
    파일 탐색기로 엑셀/한글 템플릿 경로를 선택받아 반환.
    취소 시 ('', '') 반환.
    """
    root = tk.Tk()
    root.withdraw()

    excel_path = filedialog.askopenfilename(
        title="변환할 엑셀 파일을 선택하세요",
        filetypes=[("Excel Files", "*.xlsx;*.xls")],
    )
    if not excel_path:
        return "", ""

    hwp_path = filedialog.askopenfilename(
        title="한글 템플릿 파일을 선택하세요",
        filetypes=[("HWP Files", "*.hwpx")],
    )
    return excel_path, hwp_path


# ─────────────────────────────────────────────
# 저장
# ─────────────────────────────────────────────

def save_outputs(hwp, output_dir: str) -> None:
    """HWP 및 PDF로 결과물 저장"""
    today = datetime.now().strftime("%Y%m%d")
    base_name = f"공기연장보고서_{today}"

    hwp_out = os.path.join(output_dir, f"{base_name}.hwp")
    pdf_out = os.path.join(output_dir, f"{base_name}.pdf")

    for path, fmt in [(hwp_out, "HWPX"), (pdf_out, "PDF")]:
        try:
            hwp.SaveAs(path, fmt, "")
            print(f"  저장 완료: {path}")
        except Exception as e:
            print(f"  [경고] {fmt} 저장 실패: {e}")


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────

def run() -> None:
    # 1. 파일 선택
    excel_path, hwp_template_path = select_files()
    if not excel_path or not hwp_template_path:
        print("파일 선택이 취소되었습니다.")
        return

    # 2. 엑셀 데이터 읽기
    print("\n[읽기] 엑셀 데이터를 불러오는 중...")
    wb = load_workbook(excel_path)

    header        = read_header_info(wb)
    main_values   = read_main_table_values(wb)
    indirect_rows = read_indirect_labor_table(wb)
    sev_rows, sev_totals = read_severance_table(wb)
    fourth_values = read_fourth_table_values(wb)
    fifth_values  = read_fifth_table_values(wb)
    sixth_values  = read_sixth_table_values(wb)

    # 3. 한글 실행 및 템플릿 열기
    print("\n[실행] 한글 프로그램을 여는 중...")
    try:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
    except Exception as e:
        print(f"[오류] 한글 실행 실패: {e}")
        return

    hwp.XHwpWindows.Item(0).Visible = True

    try:
        hwp.Open(hwp_template_path, "HWPX", "")
    except Exception as e:
        print(f"[오류] 템플릿 파일 열기 실패: {e}")
        return

    # 4. 데이터 입력
    print("\n[입력] 한글 문서에 데이터를 입력하는 중...")
    write_header_fields(hwp, header)
    write_main_table(hwp, main_values)
    write_indirect_labor_table(hwp, indirect_rows)
    write_severance_table(hwp, sev_rows, sev_totals)
    write_fourth_table(hwp, fourth_values)
    write_fifth_table(hwp, fifth_values)
    write_sixth_table(hwp, sixth_values)

    # 5. 저장
    print("\n[저장] 결과 파일을 저장하는 중...")
    save_outputs(hwp, os.getcwd())

    print("\n✅ 모든 작업이 완료되었습니다!")


if __name__ == "__main__":
    run()

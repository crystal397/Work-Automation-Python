"""
수신자료 파일 점검 도구
- 폴더 내 모든 파일 순회 → 열림 여부 + 손상 여부 점검 → 엑셀 보고서 생성
- 압축 파일(.zip/.7z/.rar) 내부 파일도 개별 점검
- PDF 스캔본/텍스트 구분
"""

import os
import sys
import zipfile
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# 파일 형식 매핑
# ─────────────────────────────────────────────
EXT_LABEL = {
    ".xlsx": "Microsoft Excel 워크시트",
    ".xls":  "Microsoft Excel 워크시트",
    ".xlsm": "Microsoft Excel 워크시트",
    ".xlsb": "Microsoft Excel 워크시트",
    ".csv":  "CSV 파일",
    ".docx": "Microsoft Word 문서",
    ".doc":  "Microsoft Word 문서",
    ".docm": "Microsoft Word 문서",
    ".pptx": "Microsoft PowerPoint 프레젠테이션",
    ".ppt":  "Microsoft PowerPoint 프레젠테이션",
    ".pptm": "Microsoft PowerPoint 프레젠테이션",
    ".pdf":  "Adobe Acrobat 문서",
    ".hwp":  "한글 문서",
    ".hwpx": "한글 문서",
    ".zip":  "압축(ZIP) 파일",
    ".7z":   "7-Zip 압축 파일",
    ".rar":  "RAR 압축 파일",
    ".tar":  "TAR 압축 파일",
    ".gz":   "GZ 압축 파일",
    ".msg":  "Outlook 항목",
    ".eml":  "이메일 파일",
    ".txt":  "텍스트 파일",
    ".rtf":  "서식 있는 텍스트 파일",
    ".png":  "PNG 이미지",
    ".jpg":  "JPEG 이미지",
    ".jpeg": "JPEG 이미지",
    ".gif":  "GIF 이미지",
    ".bmp":  "BMP 이미지",
    ".tif":  "TIFF 이미지",
    ".tiff": "TIFF 이미지",
    ".dwg":  "AutoCAD 도면",
    ".dxf":  "AutoCAD DXF 도면",
    ".dwf":  "AutoCAD DWF 파일",
    ".mp4":  "MP4 동영상",
    ".avi":  "AVI 동영상",
    ".mov":  "MOV 동영상",
    ".mp3":  "MP3 오디오",
    ".xml":  "XML 파일",
    ".json": "JSON 파일",
    ".mpp":  "Microsoft Project 파일",
}

# ─────────────────────────────────────────────
# 파일 열기 시도
# Returns: (비고, 파일형식_override)
#   비고: 정상이면 "", 문제 있으면 사유 문자열
#   파일형식_override: PDF 스캔본/텍스트 등 특수 표기 필요 시, 나머지는 ""
# ─────────────────────────────────────────────
def try_open_file(filepath: Path) -> tuple:
    ext = filepath.suffix.lower()

    try:
        size = filepath.stat().st_size
    except (FileNotFoundError, OSError):
        return ("열기 실패 (경로 오류: 특수문자 등)", "")

    if size == 0:
        return ("파일 손상 (빈 파일)", "")

    try:
        # ── Excel 계열
        if ext in (".xlsx", ".xlsm", ".xls", ".xlsb"):
            # 1차: openpyxl
            try:
                import openpyxl as ox
                wb = ox.load_workbook(filepath, read_only=True, data_only=True)
                has = any(s.max_row and s.max_row > 0 for s in wb.worksheets)
                wb.close()
                return ("", "") if has else ("파일 손상 (내용 없음)", "")
            except Exception:
                pass

            # 2차: xlrd (구형 xls)
            try:
                import xlrd
                wb = xlrd.open_workbook(filepath)
                has = any(wb.sheet_by_index(i).nrows > 0 for i in range(wb.nsheets))
                if has:
                    raw = filepath.read_bytes()
                    irm_kw = [b'\x06DataSpaces', b'EncryptionInfo', b'EncryptedPackage',
                              b'Rights Management', b'\x00I\x00R\x00M',
                              b'aadrm.com', b'Encrypted-Rights-Data',
                              b'rms.microsoft.com', b'_wmcs', b'AUTHENTICATEDDATA']
                    if any(kw in raw for kw in irm_kw):
                        return ("열기 실패 (보안 문서: IRM/DRM 보호)", "")
                return ("", "") if has else ("파일 손상 (내용 없음)", "")
            except Exception:
                pass

            # 3차: 바이너리 분석
            try:
                raw = filepath.read_bytes()
                # OLE2 시그니처 확인
                if raw[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                    irm_kw = [b'\x06DataSpaces', b'EncryptionInfo', b'EncryptedPackage',
                              b'Rights Management', b'\x00I\x00R\x00M',
                              b'aadrm.com', b'Encrypted-Rights-Data',
                              b'rms.microsoft.com', b'_wmcs', b'AUTHENTICATEDDATA']
                    if any(kw in raw for kw in irm_kw):
                        return ("열기 실패 (보안 문서: IRM/DRM 보호)", "")
                    return ("", "")
                # HTML 기반 Excel
                if b"<html" in raw[:1024].lower() or b"<table" in raw[:1024].lower():
                    return ("", "")
                # CSV 형태
                if len(raw.decode("utf-8", errors="ignore").strip()) > 0:
                    return ("", "")
            except Exception:
                pass

            return ("열기 실패 (형식 불일치 또는 손상)", "")

        # ── Word
        elif ext == ".docx":
            import docx
            doc = docx.Document(filepath)
            has = bool(doc.paragraphs or doc.tables)
            return ("", "") if has else ("파일 손상 (내용 없음)", "")

        # ── PowerPoint
        elif ext == ".pptx":
            from pptx import Presentation
            prs = Presentation(filepath)
            return ("", "") if len(prs.slides) > 0 else ("파일 손상 (내용 없음)", "")

        # ── PDF
        elif ext == ".pdf":
            import fitz
            doc = fitz.open(filepath)
            if doc.page_count == 0:
                doc.close()
                return ("파일 손상 (내용 없음)", "")
            # AIP/DRM 보호 감지
            first_text = doc[0].get_text().lower()
            aip_kw = ["azure information protection", "this is a protected document",
                      "microsoft information protection", "rights management", "irm protected"]
            if any(kw in first_text for kw in aip_kw):
                doc.close()
                return ("열기 실패 (보안 문서: AIP/DRM 보호)", "")
            # 스캔본 vs 텍스트 구분
            sample = min(doc.page_count, 5)
            total_chars = sum(len(doc[i].get_text().strip()) for i in range(sample))
            doc.close()
            ftype = "Adobe Acrobat 문서 (텍스트)" if total_chars / sample >= 50 else "Adobe Acrobat 문서 (스캔본)"
            return ("", ftype)

        # ── 이메일
        elif ext in (".msg", ".eml"):
            return ("", "") if size > 100 else ("파일 손상 (내용 없음)", "")

        # ── 그 외: 크기로만 판단
        else:
            return ("", "")

    except PermissionError:
        return ("열기 실패 (접근 권한 없음)", "")
    except zipfile.BadZipFile:
        return ("파일 손상 (압축 파일 오류)", "")
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg or "암호" in msg:
            return ("열기 실패 (암호 보호)", "")
        elif "permission" in msg or "access" in msg or "denied" in msg:
            return ("열기 실패 (접근 권한 없음)", "")
        elif "login" in msg or "auth" in msg or "credential" in msg:
            return ("열기 실패 (로그인 필요)", "")
        elif "corrupt" in msg or "invalid" in msg or "truncat" in msg:
            return ("파일 손상 (파일 오류)", "")
        else:
            return (f"열기 실패 ({type(e).__name__})", "")


# ─────────────────────────────────────────────
# 압축 해제
# ─────────────────────────────────────────────
def extract_archive(filepath: Path, dest: Path) -> tuple:
    """압축 해제. (성공여부, 오류메시지) 반환."""
    ext = filepath.suffix.lower()
    try:
        if ext == ".zip":
            with zipfile.ZipFile(filepath, "r") as z:
                for info in z.infolist():
                    if not (info.flag_bits & 0x800):
                        try:
                            info.filename = info.filename.encode("cp437").decode("cp949")
                        except Exception:
                            try:
                                info.filename = info.filename.encode("cp437").decode("utf-8")
                            except Exception:
                                pass
                    info.filename = info.filename.replace("\\", "/")
                    z.extract(info, dest)
            return (True, "")
        elif ext == ".7z":
            import py7zr
            with py7zr.SevenZipFile(filepath, mode="r") as z:
                z.extractall(dest)
            return (True, "")
        elif ext == ".rar":
            import rarfile
            with rarfile.RarFile(filepath, "r") as z:
                z.extractall(dest)
            return (True, "")
        else:
            return (False, "지원하지 않는 압축 형식")
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            return (False, "열기 실패 (암호 보호)")
        return (False, f"압축 해제 실패 ({type(e).__name__})")


# ─────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────
def to_extended_path(p: Path) -> Path:
    """Windows 긴 경로 처리 (UNC 확장 경로 적용)"""
    s = str(p)
    prefix = "\\\\?\\"
    if os.name == "nt" and not s.startswith(prefix):
        s = prefix + s
    return Path(s)

MAX_DEPTH = 7

total_cols = 1 + MAX_DEPTH + 3
# 1: 번호
# MAX_DEPTH: 폴더 depth
# 3: 파일명, 파일형식, 비고

last_col_letter = get_column_letter(total_cols)

def get_folder_hierarchy(root: Path, filepath: Path) -> list:
    """루트 기준 상대 경로에서 폴더 계층 추출 (최대 4단계)"""
    rel = filepath.relative_to(root)
    parts = list(rel.parts[:-1])
    while len(parts) < MAX_DEPTH:
        parts.append("")
    return parts[:MAX_DEPTH]


def fill_parts(base_parts: list, zip_name: str) -> list:
    """압축파일명을 분류 빈 자리에 채워넣기. 빈 자리 없으면 세부분류 덮어쓰기."""
    parts = list(base_parts)
    for i in range(MAX_DEPTH - 1, -1, -1):
        if parts[i] == "":
            parts[i] = zip_name
            return parts
    parts[MAX_DEPTH - 1] = zip_name
    return parts


def make_rec(parts, fname, ftype, note, path_str):
    rec = {
        f"{i+1}단계": parts[i] for i in range(MAX_DEPTH)
    }
    rec.update({
        "파일명": fname,
        "파일형식": ftype,
        "비고": note,
        "_path": path_str,
        "_flagged": bool(note),
    })
    return rec


def ext_to_label(ext: str) -> str:
    return EXT_LABEL.get(ext, ext[1:].upper() + " 파일" if ext else "알 수 없음")


# ─────────────────────────────────────────────
# 압축 내부 점검
# ─────────────────────────────────────────────
ARCHIVE_EXTS = {".zip", ".7z", ".rar"}

def scan_archive(filepath: Path, base_parts: list, zip_name: str) -> list:
    records = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="chk_"))
    try:
        ok, err = extract_archive(filepath, tmp_dir)
        if not ok:
            parts = fill_parts(base_parts, zip_name)
            records.append(make_rec(parts, zip_name, ext_to_label(filepath.suffix.lower()), err, str(filepath)))
            return records

        inner_files = sorted(f for f in tmp_dir.rglob("*") if f.is_file())
        if not inner_files:
            parts = fill_parts(base_parts, zip_name)
            records.append(make_rec(parts, "", "", "파일 손상 (빈 압축파일)", str(filepath)))
            return records

        for inner in inner_files:
            inner_ext = inner.suffix.lower()
            rel_name = str(inner.relative_to(tmp_dir))
            note, ftype_override = try_open_file(inner)
            ftype = ftype_override if ftype_override else ext_to_label(inner_ext)
            parts = fill_parts(base_parts, zip_name)
            records.append(make_rec(parts, rel_name, ftype, note, str(filepath) + "/" + rel_name))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return records


# ─────────────────────────────────────────────
# 폴더 순회
# ─────────────────────────────────────────────
def scan_folder(root: Path) -> list:
    records = []
    root = to_extended_path(root.resolve())

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        dir_path = Path(dirpath)

        if not filenames:
            has_child = any(files for _, _, files in os.walk(dirpath))
            if not has_child:
                parts = get_folder_hierarchy(root, dir_path / "_dummy_")
                rec = make_rec(parts, "", "", "폴더 비었음", str(dir_path))
                rec["_flagged"] = False
                records.append(rec)
            continue

        for fname in sorted(filenames):
            fpath = dir_path / fname
            ext = fpath.suffix.lower()
            parts = get_folder_hierarchy(root, fpath)

            if ext in ARCHIVE_EXTS:
                records.extend(scan_archive(fpath, parts, fname))
            else:
                note, ftype_override = try_open_file(fpath)
                ftype = ftype_override if ftype_override else ext_to_label(ext)
                records.append(make_rec(parts, fname, ftype, note, str(fpath)))

    return records


# ─────────────────────────────────────────────
# 스타일 상수
# ─────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
DAMAGE_FILL = PatternFill("solid", fgColor="FFE0E0")
EMPTY_FILL  = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL    = PatternFill("solid", fgColor="F2F7FF")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")
THIN        = Side(style="thin", color="AAAAAA")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FONT = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
NORMAL_FONT = Font(name="맑은 고딕", size=9)
BOLD_FONT   = Font(name="맑은 고딕", size=9, bold=True)
DAMAGE_FONT = Font(name="맑은 고딕", size=9, color="C00000", bold=True)
TITLE_FONT  = Font(name="맑은 고딕", size=14, bold=True, color="FFFFFF")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
COL_WIDTHS = {"A": 5, "B": 16, "C": 28, "D": 28, "E": 28, "F": 40, "G": 28, "H": 18}
HEADERS = [
    "번호",
    "1단계", "2단계", "3단계", "4단계", "5단계", "6단계", "7단계",
    "파일명", "파일형식", "비고"
]


def write_cell(ws, row, col, value, font=None, fill=None, alignment=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:      c.font      = font
    if fill:      c.fill      = fill
    if alignment: c.alignment = alignment
    if border:    c.border    = border
    return c


# ─────────────────────────────────────────────
# 엑셀 보고서 생성
# ─────────────────────────────────────────────
def build_excel(records: list, out_path: str, folder_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "수신자료 정리"

    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # 타이틀
    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A1:{last_col_letter}1")
    c = ws["A1"]
    c.value     = f"수신자료 정리 / {datetime.now().strftime('%y%m%d')}_파일 수령 현황"
    c.font      = TITLE_FONT
    c.fill      = HEADER_FILL
    c.alignment = CENTER

    # 경로
    ws.row_dimensions[2].height = 18
    ws.merge_cells(f"A2:{last_col_letter}2")
    c = ws["A2"]
    c.value     = f"대상 폴더: {folder_path}"
    c.font      = Font(name="맑은 고딕", size=8, italic=True, color="595959")
    c.fill      = PatternFill("solid", fgColor="D6E4F0")
    c.alignment = LEFT

    # 헤더
    ws.row_dimensions[3].height = 22
    for col_idx, hdr in enumerate(HEADERS, 1):
        write_cell(ws, 3, col_idx, hdr, font=HEADER_FONT, fill=HEADER_FILL,
                   alignment=CENTER, border=THIN_BORDER)

    # 데이터
    flagged_list = []
    for row_num, (idx, rec) in enumerate(enumerate(records, 1), 4):
        ws.row_dimensions[row_num].height = 16
        is_flagged = rec["_flagged"]
        bigo = rec["비고"]
        is_empty = "폴더 비었음" in bigo

        row_fill = (
            EMPTY_FILL  if is_empty  else
            DAMAGE_FILL if is_flagged else
            ALT_FILL    if row_num % 2 == 0 else
            WHITE_FILL
        )

        vals = [idx] + [rec[f"{i+1}단계"] for i in range(MAX_DEPTH)] + [rec["파일명"], rec["파일형식"], bigo]

        for col_idx, val in enumerate(vals, 1):
            font = (DAMAGE_FONT if is_flagged and col_idx in (6, 8)
                    else BOLD_FONT if col_idx == 1
                    else NORMAL_FONT)
            write_cell(ws, row_num, col_idx, val, font=font, fill=row_fill,
                       alignment=CENTER if col_idx in (1, 7, 8) else LEFT,
                       border=THIN_BORDER)

        if is_flagged and not is_empty:
            flagged_list.append(rec)

    # 집계
    mid = total_cols // 2

    summary_row = 3 + len(records) + 2
    ws.row_dimensions[summary_row].height = 20
    ws.merge_cells(f"A{summary_row}:{get_column_letter(mid)}{summary_row}")
    c = ws.cell(row=summary_row, column=1,
                value=f"총 파일 수: {len([r for r in records if r['파일명']])}건")
    c.font = BOLD_FONT
    c.fill = PatternFill("solid", fgColor="D6E4F0")
    c.alignment = CENTER
    c.border = THIN_BORDER

    start_col = mid + 1  # 병합 시작 컬럼
    ws.merge_cells(f"{get_column_letter(start_col)}{summary_row}:{last_col_letter}{summary_row}")
    c = ws.cell(row=summary_row, column=start_col, value=f"비고 기록: {len(flagged_list)}건")
    c.font = Font(name="맑은 고딕", size=9, bold=True, color="C00000")
    c.fill = DAMAGE_FILL
    c.alignment = CENTER
    c.border = THIN_BORDER

    # 비고 목록 시트
    ws2 = wb.create_sheet("비고 기록 파일 목록")
    for col_letter, width in COL_WIDTHS.items():
        ws2.column_dimensions[col_letter].width = width
    ws2.column_dimensions["F"].width = 50
    ws2.column_dimensions["H"].width = 25

    ws2.row_dimensions[1].height = 28
    ws2.merge_cells(f"A1:{last_col_letter}1")
    c = ws2["A1"]
    c.value     = "비고 기록 파일 목록"
    c.font      = TITLE_FONT
    c.fill      = PatternFill("solid", fgColor="C00000")
    c.alignment = CENTER

    ws2.row_dimensions[2].height = 22
    for col_idx, hdr in enumerate(HEADERS, 1):
        write_cell(ws2, 2, col_idx, hdr, font=HEADER_FONT,
                   fill=PatternFill("solid", fgColor="C00000"),
                   alignment=CENTER, border=THIN_BORDER)

    for r_idx, rec in enumerate(flagged_list, 3):
        ws2.row_dimensions[r_idx].height = 16
        vals = [r_idx - 2, rec["대분류"], rec["중분류"], rec["소분류"],
                rec["세부분류"], rec["파일명"], rec["파일형식"], rec["비고"]]
        for col_idx, val in enumerate(vals, 1):
            write_cell(ws2, r_idx, col_idx, val,
                       font=DAMAGE_FONT if col_idx in (6, 8) else NORMAL_FONT,
                       fill=DAMAGE_FILL,
                       alignment=CENTER if col_idx in (1, 7, 8) else LEFT,
                       border=THIN_BORDER)

    wb.save(out_path)
    return flagged_list


# ─────────────────────────────────────────────
# 경로 입력 (GUI → 콘솔 fallback)
# ─────────────────────────────────────────────
def get_folder_path_gui() -> str:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("수신자료 파일 점검", "점검할 폴더를 선택해 주세요.")
    folder = filedialog.askdirectory(title="점검 대상 폴더 선택")
    root.destroy()
    return folder


def get_folder_path_console() -> str:
    print("=" * 50)
    print("  수신자료 파일 점검 도구")
    print("=" * 50)
    print()
    while True:
        path = input("점검할 폴더 경로를 입력하세요: ").strip().strip('"')
        if path:
            return path
        print("[ERROR] 경로를 입력해 주세요.")


# ─────────────────────────────────────────────
# 실행
# ─────────────────────────────────────────────
if __name__ == "__main__":
    folder_path = ""
    try:
        folder_path = get_folder_path_gui()
        if not folder_path:
            print("[INFO] GUI 취소 → 콘솔 입력으로 전환합니다.")
            folder_path = get_folder_path_console()
    except Exception as e:
        print(f"[INFO] GUI 사용 불가 ({e}) → 콘솔 입력으로 전환합니다.")
        folder_path = get_folder_path_console()

    folder = Path(folder_path)
    if not folder.exists():
        print(f"[ERROR] 폴더를 찾을 수 없습니다: {folder_path}")
        input("엔터를 눌러 종료...")
        sys.exit(1)

    print(f"[INFO] 폴더 순회 시작: {folder_path}")
    records = scan_folder(folder)
    print(f"[INFO] 총 {len(records)}건 확인 완료")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_out = f"수신자료_파일점검결과_{timestamp}.xlsx"

    flagged = build_excel(records, excel_out, folder_path)
    print(f"[OK] 엑셀 보고서 저장: {excel_out}")
    print(f"[INFO] 비고 기록 파일: {len(flagged)}건")

    print("\n=== 완료 ===")
    print(f"  보고서: {excel_out}")
    if flagged:
        print(f"\n  비고 기록 파일 목록 ({len(flagged)}건):")
        for rec in flagged:
            cat = " > ".join(v for v in [rec["대분류"], rec["중분류"], rec["소분류"], rec["세부분류"]] if v)
            print(f"    [{cat}]  {rec['파일명']}  →  {rec['비고']}")

    input("\n엔터를 눌러 종료...")
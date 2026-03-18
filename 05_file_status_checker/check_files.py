"""
폴더 내 파일 순회 → 파일 열림 여부 + 손상 여부 점검 → 엑셀 보고서 + 회신문 생성
압축 파일(.zip/.7z/.rar) 내부 파일도 개별 점검
"""

import os
import sys
import zipfile
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────
FOLDER_PATH = r"C:\Users\CT Group\씨티그룹 Dropbox\200  연구원(공통)\분쟁지원팀\2026년\260123_대우건설_수출용 신형 연구로 건설공사 클레임\0. 수신자료\260317_1차 요청자료\1차 요청자료"

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
# ─────────────────────────────────────────────
def try_open_file(filepath: Path) -> str:
    """
    파일을 열어 상태를 반환.
    Returns: 비고 문자열 (정상이면 "", 문제 있으면 사유)
    """
    ext = filepath.suffix.lower()

    try:
        size = filepath.stat().st_size
    except (FileNotFoundError, OSError):
        return "열기 실패 (경로 오류: 특수문자 등)"

    if size == 0:
        return "파일 손상 (빈 파일)"

    try:
        if ext in (".xlsx", ".xlsm", ".xls", ".xlsb"):
            # 1차: openpyxl 시도 (xlsx/xlsm)
            try:
                import openpyxl as ox
                wb = ox.load_workbook(filepath, read_only=True, data_only=True)
                has = any(sheet.max_row and sheet.max_row > 0 for sheet in wb.worksheets)
                wb.close()
                return "" if has else "파일 손상 (내용 없음)"
            except Exception as e1:
                pass

            # 2차: xlrd 시도 (xls 구형 형식)
            try:
                import xlrd
                wb = xlrd.open_workbook(filepath)
                has = any(wb.sheet_by_index(i).nrows > 0 for i in range(wb.nsheets))
                if has:
                    # OLE2 파일에서 IRM/DRM 보호 여부 바이너리 탐색
                    try:
                        raw = filepath.read_bytes()
                        irm_keywords = [b'\x06DataSpaces', b'EncryptionInfo', b'EncryptedPackage',
                                        b'Rights Management', b'\x00I\x00R\x00M']
                        if any(kw in raw for kw in irm_keywords):
                            return "열기 실패 (보안 문서: IRM/DRM 보호)"
                    except Exception:
                        pass
                return "" if has else "파일 손상 (내용 없음)"
            except Exception as e2:
                pass

            # 3차: 텍스트/HTML로 읽기 시도 (확장자 불일치 케이스)
            try:
                raw = filepath.read_bytes()

                # OLE2 형식 시그니처 확인 (구형 xls)
                is_ole2 = raw[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
                if is_ole2:
                    # IRM/DRM 보호 키워드 탐색
                    irm_keywords = [
                        b'\x06DataSpaces', b'EncryptionInfo', b'EncryptedPackage',
                        b'Rights Management', b'\x00I\x00R\x00M',
                        b'aadrm.com', b'Encrypted-Rights-Data', b'rms.microsoft.com',
                        b'_wmcs', b'AUTHENTICATEDDATA',
                    ]
                    if any(kw in raw for kw in irm_keywords):
                        return "열기 실패 (보안 문서: IRM/DRM 보호)"
                    return ""  # OLE2지만 보호 없음

                # HTML 기반 Excel 파일 감지
                if b"<html" in raw[:1024].lower() or b"<table" in raw[:1024].lower():
                    return ""  # HTML 형식이지만 열 수 있는 파일

                # CSV 형태 감지
                text = raw.decode("utf-8", errors="ignore")
                if len(text.strip()) > 0:
                    return ""
            except Exception:
                pass

            return "열기 실패 (형식 불일치 또는 손상)"

        elif ext == ".docx":
            import docx
            doc = docx.Document(filepath)
            has = bool(doc.paragraphs or doc.tables)
            return "" if has else "파일 손상 (내용 없음)"

        elif ext == ".pptx":
            from pptx import Presentation
            prs = Presentation(filepath)
            return "" if len(prs.slides) > 0 else "파일 손상 (내용 없음)"

        elif ext == ".pdf":
            import fitz
            doc = fitz.open(filepath)
            has = doc.page_count > 0
            if has:
                # AIP / DRM 보호 파일 감지: 첫 페이지 텍스트에서 보안 키워드 확인
                first_text = doc[0].get_text().lower()
                doc.close()
                aip_keywords = [
                    "azure information protection",
                    "this is a protected document",
                    "microsoft information protection",
                    "rights management",
                    "irm protected",
                ]
                if any(kw in first_text for kw in aip_keywords):
                    return "열기 실패 (보안 문서: AIP/DRM 보호)"
                return ""
            doc.close()
            return "파일 손상 (내용 없음)"

        elif ext in (".msg", ".eml"):
            return "" if size > 100 else "파일 손상 (내용 없음)"

        else:
            # 지원하지 않는 형식: 크기로만 판단
            return ""

    except PermissionError:
        return "열기 실패 (접근 권한 없음)"
    except zipfile.BadZipFile:
        return "파일 손상 (압축 파일 오류)"
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg or "암호" in msg:
            return "열기 실패 (암호 보호)"
        elif "permission" in msg or "access" in msg or "denied" in msg:
            return "열기 실패 (접근 권한 없음)"
        elif "login" in msg or "auth" in msg or "credential" in msg:
            return "열기 실패 (로그인 필요)"
        elif "corrupt" in msg or "invalid" in msg or "truncat" in msg:
            return "파일 손상 (파일 오류)"
        else:
            return f"열기 실패 ({type(e).__name__})"

# ─────────────────────────────────────────────
# 압축 해제 및 내부 파일 점검
# ─────────────────────────────────────────────
def fix_zip_filename(name_bytes: bytes) -> str:
    """ZIP 내부 파일명 인코딩 자동 감지 (CP949 → UTF-8 순으로 시도)"""
    for enc in ("utf-8", "cp949", "euc-kr", "latin-1"):
        try:
            return name_bytes.encode("cp437").decode(enc)
        except Exception:
            continue
    return name_bytes  # fallback

def extract_archive(filepath: Path, dest: Path) -> tuple[bool, str]:
    """압축 해제. 성공 여부와 오류 메시지 반환."""
    ext = filepath.suffix.lower()
    try:
        if ext == ".zip":
            with zipfile.ZipFile(filepath, "r") as z:
                for info in z.infolist():
                    # 파일명 인코딩 처리: UTF-8 플래그가 없으면 CP949 시도
                    if not (info.flag_bits & 0x800):  # UTF-8 플래그 없음
                        try:
                            info.filename = info.filename.encode("cp437").decode("cp949")
                        except Exception:
                            try:
                                info.filename = info.filename.encode("cp437").decode("utf-8")
                            except Exception:
                                pass  # 원본 유지
                    # 경로 구분자 정규화
                    info.filename = info.filename.replace("\\", "/")
                    z.extract(info, dest)
            return True, ""
        elif ext == ".7z":
            import py7zr
            with py7zr.SevenZipFile(filepath, mode="r") as z:
                z.extractall(dest)
            return True, ""
        elif ext == ".rar":
            import rarfile
            with rarfile.RarFile(filepath, "r") as z:
                z.extractall(dest)
            return True, ""
        else:
            return False, "지원하지 않는 압축 형식"
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            return False, "열기 실패 (암호 보호)"
        return False, f"압축 해제 실패 ({type(e).__name__})"

def scan_archive(filepath: Path, base_parts: list[str], zip_name: str) -> list[dict]:
    """
    압축 파일 내부를 임시 폴더에 해제 후 개별 파일 점검.
    base_parts: 압축파일이 속한 폴더의 분류 [대,중,소,세부]
    zip_name: 압축파일명 (분류 빈 자리에 채워넣을 값)
    """
    records = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="chk_"))

    try:
        ok, err = extract_archive(filepath, tmp_dir)
        if not ok:
            # 해제 실패 → 압축파일 자체를 1행으로 기록
            parts = _fill_parts(base_parts, zip_name)
            records.append(_make_rec(parts, zip_name,
                                     EXT_LABEL.get(filepath.suffix.lower(), "압축 파일"),
                                     err, str(filepath)))
            return records

        # 해제 성공 → 내부 파일 순회
        inner_files = sorted(tmp_dir.rglob("*"))
        inner_files = [f for f in inner_files if f.is_file()]

        if not inner_files:
            parts = _fill_parts(base_parts, zip_name)
            records.append(_make_rec(parts, "",  "", "파일 손상 (빈 압축파일)", str(filepath)))
            return records

        for inner in inner_files:
            inner_ext = inner.suffix.lower()
            # 내부 파일의 상대 경로를 파일명으로 표시 (하위 폴더 있을 경우 포함)
            rel_name = str(inner.relative_to(tmp_dir))
            note = try_open_file(inner)
            parts = _fill_parts(base_parts, zip_name)
            records.append(_make_rec(
                parts,
                rel_name,
                EXT_LABEL.get(inner_ext, inner_ext[1:].upper() + " 파일" if inner_ext else "알 수 없음"),
                note,
                str(filepath) + "/" + rel_name,
            ))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return records

def _fill_parts(base_parts: list[str], zip_name: str) -> list[str]:
    """
    base_parts의 빈 자리(뒤에서부터)에 zip_name을 채워넣음.
    모두 차있으면 세부분류(인덱스 3)를 zip_name으로 교체.
    """
    parts = list(base_parts)
    for i in range(3, -1, -1):   # 세부→소→중→대 순으로 빈 자리 탐색
        if parts[i] == "":
            parts[i] = zip_name
            return parts
    # 빈 자리 없음 → 세부분류 덮어쓰기
    parts[3] = zip_name
    return parts

def _make_rec(parts, fname, ftype, note, path_str):
    return {
        "대분류":  parts[0],
        "중분류":  parts[1],
        "소분류":  parts[2],
        "세부분류": parts[3],
        "파일명":  fname,
        "파일형식": ftype,
        "비고":    note,
        "_path":   path_str,
        "_flagged": bool(note),
    }

# ─────────────────────────────────────────────
# 경로 유틸
# ─────────────────────────────────────────────
def to_extended_path(p: Path) -> Path:
    """Windows에서 긴 경로 및 특수문자 경로 처리 (UNC 확장 경로 적용)"""
    s = str(p)
    prefix = "\\\\?\\"
    if os.name == "nt" and not s.startswith(prefix):
        s = prefix + s
    return Path(s)

def get_folder_hierarchy(root: Path, filepath: Path) -> list[str]:
    """루트 기준 상대 경로에서 폴더 계층 추출 (최대 4단계)"""
    rel = filepath.relative_to(root)
    parts = list(rel.parts[:-1])
    while len(parts) < 4:
        parts.append("")
    return parts[:4]

# ─────────────────────────────────────────────
# 폴더 순회
# ─────────────────────────────────────────────
ARCHIVE_EXTS = {".zip", ".7z", ".rar"}

def scan_folder(root: Path) -> list[dict]:
    records = []
    root = to_extended_path(root.resolve())

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        dir_path = Path(dirpath)

        if not filenames:
            has_any_child_file = any(files for _, _, files in os.walk(dirpath))
            if not has_any_child_file:
                parts = get_folder_hierarchy(root, dir_path / "_dummy_")
                records.append(_make_rec(parts, "", "", "폴더 비었음", str(dir_path)))
                # 빈 폴더는 _flagged=False 로 재설정
                records[-1]["_flagged"] = False
            continue

        for fname in sorted(filenames):
            fpath = dir_path / fname
            ext = fpath.suffix.lower()
            parts = get_folder_hierarchy(root, fpath)

            if ext in ARCHIVE_EXTS:
                # 압축 파일 → 내부 점검
                inner_records = scan_archive(fpath, parts, fname)
                records.extend(inner_records)
            else:
                note = try_open_file(fpath)
                records.append(_make_rec(
                    parts, fname,
                    EXT_LABEL.get(ext, ext[1:].upper() + " 파일" if ext else "알 수 없음"),
                    note, str(fpath)
                ))

    return records

# ─────────────────────────────────────────────
# 스타일 상수
# ─────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
DAMAGE_FILL = PatternFill("solid", fgColor="FFE0E0")
EMPTY_FILL  = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL    = PatternFill("solid", fgColor="F2F7FF")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")

THIN = Side(style="thin", color="AAAAAA")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADER_FONT = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
NORMAL_FONT = Font(name="맑은 고딕", size=9)
BOLD_FONT   = Font(name="맑은 고딕", size=9, bold=True)
DAMAGE_FONT = Font(name="맑은 고딕", size=9, color="C00000", bold=True)
TITLE_FONT  = Font(name="맑은 고딕", size=14, bold=True, color="FFFFFF")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

COL_WIDTHS = {"A": 5, "B": 16, "C": 28, "D": 28, "E": 28, "F": 40, "G": 28, "H": 18}
HEADERS = ["번호", "대분류", "중분류", "소분류", "세부분류", "파일명", "파일형식", "비고"]

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
def build_excel(records: list[dict], out_path: str, folder_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "수신자료 정리"

    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # 타이틀
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value     = f"수신자료 정리 / {datetime.now().strftime('%y%m%d')}_파일 수령 현황"
    c.font      = TITLE_FONT
    c.fill      = HEADER_FILL
    c.alignment = CENTER

    # 경로
    ws.row_dimensions[2].height = 18
    ws.merge_cells("A2:H2")
    c = ws["A2"]
    c.value     = f"대상 폴더: {folder_path}"
    c.font      = Font(name="맑은 고딕", size=8, italic=True, color="595959")
    c.fill      = PatternFill("solid", fgColor="D6E4F0")
    c.alignment = LEFT

    # 헤더
    ws.row_dimensions[3].height = 22
    for col_idx, hdr in enumerate(HEADERS, 1):
        write_cell(ws, 3, col_idx, hdr,
                   font=HEADER_FONT, fill=HEADER_FILL,
                   alignment=CENTER, border=THIN_BORDER)

    # 데이터
    flagged_list = []
    for row_num, (idx, rec) in enumerate(enumerate(records, 1), 4):
        ws.row_dimensions[row_num].height = 16
        is_flagged = rec["_flagged"]
        bigo = rec["비고"]
        is_empty = "폴더 비었음" in bigo

        row_fill = (
            EMPTY_FILL  if is_empty else
            DAMAGE_FILL if is_flagged else
            ALT_FILL    if row_num % 2 == 0 else
            WHITE_FILL
        )

        vals = [idx, rec["대분류"], rec["중분류"], rec["소분류"],
                rec["세부분류"], rec["파일명"], rec["파일형식"], bigo]

        for col_idx, val in enumerate(vals, 1):
            font = (DAMAGE_FONT if is_flagged and col_idx in (6, 8)
                    else BOLD_FONT if col_idx == 1
                    else NORMAL_FONT)
            write_cell(ws, row_num, col_idx, val,
                       font=font, fill=row_fill,
                       alignment=CENTER if col_idx in (1, 7, 8) else LEFT,
                       border=THIN_BORDER)

        if is_flagged and not is_empty:
            flagged_list.append(rec)

    # 집계
    summary_row = 3 + len(records) + 2
    ws.row_dimensions[summary_row].height = 20
    ws.merge_cells(f"A{summary_row}:E{summary_row}")
    c = ws.cell(row=summary_row, column=1,
                value=f"총 파일 수: {len([r for r in records if r['파일명']])}건")
    c.font = BOLD_FONT
    c.fill = PatternFill("solid", fgColor="D6E4F0")
    c.alignment = CENTER
    c.border = THIN_BORDER

    ws.merge_cells(f"F{summary_row}:H{summary_row}")
    c = ws.cell(row=summary_row, column=6,
                value=f"비고 기록: {len(flagged_list)}건")
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
    ws2.merge_cells("A1:H1")
    c = ws2["A1"]
    c.value     = "비고 기록 파일 목록"
    c.font      = TITLE_FONT
    c.fill      = PatternFill("solid", fgColor="C00000")
    c.alignment = CENTER

    ws2.row_dimensions[2].height = 22
    for col_idx, hdr in enumerate(HEADERS, 1):
        write_cell(ws2, 2, col_idx, hdr,
                   font=HEADER_FONT,
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
# 실행
# ─────────────────────────────────────────────
def get_folder_path_gui() -> str:
    """tkinter GUI로 폴더 선택 다이얼로그 표시"""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()  # 메인 창 숨김

    messagebox.showinfo(
        "수신자료 파일 점검",
        "점검할 폴더를 선택해 주세요."
    )

    folder = filedialog.askdirectory(title="점검 대상 폴더 선택")
    root.destroy()
    return folder


def get_folder_path_console() -> str:
    """콘솔에서 폴더 경로 직접 입력"""
    print("=" * 50)
    print("  수신자료 파일 점검 도구")
    print("=" * 50)
    print()
    while True:
        path = input("점검할 폴더 경로를 입력하세요: ").strip().strip('"')
        if path:
            return path
        print("[ERROR] 경로를 입력해 주세요.")


if __name__ == "__main__":
    # ── 경로 설정: GUI → 실패 시 콘솔 fallback
    folder_path = ""
    try:
        folder_path = get_folder_path_gui()
        if not folder_path:
            print("[INFO] GUI에서 폴더 선택이 취소되었습니다. 콘솔 입력으로 전환합니다.")
            folder_path = get_folder_path_console()
    except Exception as e:
        print(f"[INFO] GUI를 사용할 수 없습니다 ({e}). 콘솔 입력으로 전환합니다.")
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
    docx_out  = f"수신자료_재송부요청_{timestamp}.docx"

    flagged = build_excel(records, excel_out, folder_path)
    print(f"[OK] 엑셀 보고서 저장: {excel_out}")
    print(f"[INFO] 비고 기록 파일: {len(flagged)}건")

    build_reply_docx(flagged, docx_out, folder_path)

    copy_dest = Path(f"비고파일_모음_{timestamp}")
    copy_flagged_files(flagged, copy_dest)

    print("\n=== 완료 ===")
    print(f"  보고서  : {excel_out}")
    print(f"  회신문  : {docx_out}")
    print(f"  복사 폴더: {copy_dest}")
    if flagged:
        print(f"\n  비고 기록 파일 목록 ({len(flagged)}건):")
        for rec in flagged:
            cat = " > ".join(
                v for v in [rec["대분류"], rec["중분류"], rec["소분류"], rec["세부분류"]] if v
            )
            print(f"    [{cat}]  {rec['파일명']}  →  {rec['비고']}")

    input("\n엔터를 눌러 종료...")
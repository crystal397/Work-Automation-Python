import os
import time
import json
import glob
from datetime import datetime
from dotenv import load_dotenv
from pypdf import PdfReader, PdfWriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# .env 파일 로드 (보안 설정)
load_dotenv()

# ─────────────────────────────────────────────
#  ① 설정 및 경로 (환경변수 활용)
# ─────────────────────────────────────────────
LOGIN_URL = os.getenv("TARGET_URL", "https://automation-target-site.com/login")
USERNAME  = os.getenv("SITE_USERNAME")
PASSWORD  = os.getenv("SITE_PASSWORD")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(BASE_DIR, "data", "inputs")
PROGRESS_FILE = os.path.join(BASE_DIR, "logs", "upload_progress.json")
MAX_FILE_SIZE = 9 * 1024 * 1024  # 9MB 제한

# 9가지 카테고리 정보 (메뉴 ID와 카테고리명)
CATEGORIES = [
    ("menu_energy", "energy"),
    ("menu_tax", "tax"),
    ("menu_labor", "labor"),
    ("menu_printing", "printing"),
    ("menu_supplies", "supplies"),
    ("menu_travel", "travel"),
    ("menu_welfare", "welfare"),
    ("menu_rent", "rent"),
    ("menu_fee", "fee"),
]

# ─────────────────────────────────────────────
#  ② 유틸리티 함수
# ─────────────────────────────────────────────

def save_progress(filename, chunk_idx):
    os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
    progress = {}
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            progress = json.load(f)
    progress[filename] = chunk_idx
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(progress, f, indent=4)

def convert_excel_to_pdf(excel_path):
    """Excel을 PDF로 변환 (win32com 활용)"""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    abs_path = os.path.abspath(excel_path)
    pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(abs_path)
        for ws in wb.Worksheets:
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        return pdf_path
    except Exception as e:
        print(f"변환 실패: {e}")
        return None

def get_pdf_chunks(file_path):
    """9MB 초과 시 PDF 분할"""
    if os.path.getsize(file_path) <= MAX_FILE_SIZE:
        return [file_path]
    
    reader = PdfReader(file_path)
    chunks = []
    for i in range(0, len(reader.pages), 5): # 5페이지씩 분할
        writer = PdfWriter()
        for page in reader.pages[i:i+5]:
            writer.add_page(page)
        chunk_path = f"{os.path.splitext(file_path)[0]}_p{len(chunks)+1}.pdf"
        with open(chunk_path, "wb") as f:
            writer.write(f)
        chunks.append(chunk_path)
    return chunks

# ─────────────────────────────────────────────
#  ③ 핵심 자동화 로직
# ─────────────────────────────────────────────

def run_automation():
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 15)

    try:
        # 로그인
        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.ID, "user_id"))).send_keys(USERNAME)
        driver.find_element(By.ID, "user_pw").send_keys(PASSWORD)
        driver.find_element(By.ID, "login_btn").click()

        for menu_id, cat_name in CATEGORIES:
            print(f"▶ 카테고리 시작: {cat_name}")
            
            # 메뉴 이동
            wait.until(EC.element_to_be_clickable((By.ID, menu_id))).click()
            time.sleep(1)

            files = glob.glob(os.path.join(INPUT_FOLDER, f"*{cat_name}*.*"))
            for file_path in files:
                # 엑셀 처리
                if file_path.endswith(('.xlsx', '.xls')):
                    file_path = convert_excel_to_pdf(file_path)
                
                # 분할 처리 및 업로드
                chunks = get_pdf_chunks(file_path)
                for idx, chunk in enumerate(chunks, 1):
                    # 파일 선택 및 전송
                    upload_input = wait.until(EC.presence_of_element_located((By.NAME, "upload_file")))
                    upload_input.send_keys(os.path.abspath(chunk))
                    
                    # 변환/업로드 버튼 클릭
                    driver.find_element(By.ID, "btn_convert").click()
                    wait.until(EC.alert_is_present()).accept() # 알림창 확인
                    
                    save_progress(os.path.basename(file_path), idx)
                    print(f"  ✓ 완료: {os.path.basename(chunk)}")

    finally:
        driver.quit()

if __name__ == "__main__":
    run_automation()

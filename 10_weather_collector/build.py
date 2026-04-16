import subprocess
import sys
from pathlib import Path

args = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name", "기상데이터수집기",
    "--collect-all", "customtkinter",
    "--hidden-import", "sqlalchemy.dialects.sqlite",
    "--hidden-import", "openpyxl",
    "--hidden-import", "pandas",
    "--hidden-import", "dotenv",
    "gui.py",
]

print("빌드 시작...")
result = subprocess.run(args)
if result.returncode != 0:
    print("\n[ERROR] 빌드 실패.")
    sys.exit(1)

print("\n빌드 완료!")
print("위치: dist\\기상데이터수집기.exe")

# ── 빌드 후처리 ──────────────────────────────────────
dist_dir = Path("dist")

# 1) .env — 실제 API 키 제거, 빈 템플릿 생성 (보안)
env_dest = dist_dir / ".env"
env_dest.write_text("KMA_API_KEY=\n", encoding="utf-8")
print("\n[보안] dist/.env: API 키 빈 템플릿 생성")
print("       배포 수신자가 프로그램 최초 실행 시 직접 입력합니다.")

# 2) 사용안내.txt 복사
guide_src = Path("dist") / "사용안내.txt"
if not guide_src.exists():
    local_guide = Path(__file__).parent / "사용안내.txt"
    if local_guide.exists():
        import shutil
        shutil.copy(local_guide, guide_src)
        print("[완료] 사용안내.txt dist/에 복사됨.")

print("\n배포 준비 완료: dist/ 폴더를 ZIP으로 압축하여 전달하세요.")

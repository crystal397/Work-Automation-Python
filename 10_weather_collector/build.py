import subprocess
import sys

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
print("사용안내.txt 를 dist\\ 폴더에 직접 복사해 주세요.")

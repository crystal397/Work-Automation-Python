"""진입점 — 로깅 초기화 후 GUI 시작"""
import logging
from pathlib import Path

import config

# ── 로그 디렉토리 생성 ────────────────────────────────────────────────────────
config.LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[
        logging.FileHandler(config.LOG_FILE, encoding="utf-8"),
        # GUI 핸들러는 App.__init__ 내에서 추가됨
    ],
)

if __name__ == "__main__":
    from gui import App

    app = App()
    app.mainloop()

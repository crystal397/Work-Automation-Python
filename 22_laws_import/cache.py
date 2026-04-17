"""SQLite 기반 로컬 캐시 레이어

테이블 구조:
  cache (key TEXT PK, value TEXT, cached_at TEXT)

key 네이밍 규칙:
  search:{target}:{query}        → MST 번호
  history:{target}:{mst}         → 연혁 raw list
  text:{target}:{law_id}         → 법령 본문 dict
"""
import json
import logging
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

import config

logger = logging.getLogger(__name__)


class LawCache:
    """API 응답 로컬 캐시 — SQLite"""

    def __init__(self, db_path: Path = config.CACHE_DB_PATH):
        db_path.parent.mkdir(parents=True, exist_ok=True)
        self._db_path = db_path
        self.conn = sqlite3.connect(str(db_path), check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self._init_schema()

    def _init_schema(self) -> None:
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS cache (
                key       TEXT PRIMARY KEY,
                value     TEXT NOT NULL,
                cached_at TEXT NOT NULL
            )
        """)
        self.conn.commit()

    def _is_valid(self, cached_at_str: str) -> bool:
        try:
            cached_at = datetime.fromisoformat(cached_at_str)
            return datetime.now() - cached_at < timedelta(days=config.CACHE_TTL_DAYS)
        except ValueError:
            return False

    # ── 공개 메서드 ───────────────────────────────────────────────────────────

    def get(self, key: str) -> Optional[Any]:
        """캐시에서 값 조회. 만료됐거나 없으면 None 반환."""
        cur = self.conn.execute(
            "SELECT value, cached_at FROM cache WHERE key = ?", (key,)
        )
        row = cur.fetchone()
        if row and self._is_valid(row[1]):
            logger.debug("캐시 히트: %s", key)
            return json.loads(row[0])
        return None

    def set(self, key: str, value: Any) -> None:
        """캐시에 값 저장 (upsert)."""
        self.conn.execute(
            "INSERT OR REPLACE INTO cache (key, value, cached_at) VALUES (?, ?, ?)",
            (key, json.dumps(value, ensure_ascii=False), datetime.now().isoformat()),
        )
        self.conn.commit()
        logger.debug("캐시 저장: %s", key)

    def clear(self) -> int:
        """전체 캐시 삭제 → 삭제된 행 수 반환."""
        cur = self.conn.execute("SELECT COUNT(*) FROM cache")
        count = cur.fetchone()[0]
        self.conn.execute("DELETE FROM cache")
        self.conn.commit()
        logger.info("캐시 초기화: %d건 삭제", count)
        return count

    def stats(self) -> dict:
        """캐시 현황 (건수, DB 파일 크기)."""
        cur = self.conn.execute("SELECT COUNT(*), MIN(cached_at), MAX(cached_at) FROM cache")
        row = cur.fetchone()
        size_kb = self._db_path.stat().st_size // 1024 if self._db_path.exists() else 0
        return {
            "count": row[0],
            "oldest": row[1] or "-",
            "newest": row[2] or "-",
            "size_kb": size_kb,
        }

    def close(self) -> None:
        self.conn.close()

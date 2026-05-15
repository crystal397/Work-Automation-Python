"""레거시 도구(22_laws_import) 브릿지.

22_laws_import 의 LawMatcher + WordReportGenerator 를 36 패키지에서 동적 로드.
디렉토리명에 숫자가 있어 일반 ``import`` 가 불가능하므로 sys.path 동적 추가.

사용 예::

    from contract_meta import bridges

    docx_path = bridges.run_laws_import(
        bid_date=date(2023, 8, 1),
        out_path=Path("out/<프로젝트>/5_적용법령.docx"),
        oc="user@example.com",   # 선택 — 미지정 시 22의 .env LAW_API_OC 사용
    )
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path
from types import ModuleType


def _project_root() -> Path:
    # 36_contract_meta/src/contract_meta/bridges.py 에서 3단계 위
    return Path(__file__).resolve().parents[3]


def _laws_import_dir() -> Path:
    return _project_root() / "22_laws_import"


def _ensure_laws_import_on_path() -> Path:
    """22_laws_import 디렉토리를 sys.path 에 추가 (이미 있으면 skip)."""
    laws_dir = _laws_import_dir()
    if not laws_dir.is_dir():
        raise FileNotFoundError(f"22_laws_import 미설치: {laws_dir}")
    p = str(laws_dir)
    if p not in sys.path:
        sys.path.insert(0, p)
    return laws_dir


def load_laws_modules() -> tuple[ModuleType, ModuleType, ModuleType]:
    """22 의 (config, engine, report) 3개 모듈을 import."""
    _ensure_laws_import_on_path()
    try:
        import config         # noqa: E402
        import engine         # noqa: E402
        import report         # noqa: E402
    except ModuleNotFoundError as e:
        missing = e.name
        raise ModuleNotFoundError(
            f"22_laws_import 의 의존성 '{missing}' 가 설치되어 있지 않습니다. "
            f"다음 명령으로 설치하세요:\n"
            f"    pip install -r 22_laws_import/requirements.txt\n"
            f"필요 패키지: requests, xmltodict, python-docx, python-dotenv, beautifulsoup4"
        ) from e
    return config, engine, report


def run_laws_import(
    bid_date: date,
    out_path: Path,
    *,
    oc: str | None = None,
    laws: list[tuple[str, str, str]] | None = None,
    progress_callback=None,
) -> Path:
    """22_laws_import 의 매칭 + 보고서 생성 한 번에 실행.

    Args:
        bid_date:          입찰공고일 (보통 contract_meta.project.bid_announcement_date).
        out_path:          생성될 .docx 경로.
        oc:                법제처 API OC (이메일 ID). None 이면 22 의 .env / 환경변수 사용.
        laws:              매칭 대상 법령 리스트. None 이면 22 의 config.TARGET_LAWS 전체.
        progress_callback: (i, total, name) 콜백 (선택).

    Returns:
        생성된 docx 경로.
    """
    config, engine, report = load_laws_modules()
    matcher_oc = oc if oc is not None else config.LAW_API_OC
    laws_list = laws if laws is not None else config.TARGET_LAWS

    matcher = engine.LawMatcher(oc=matcher_oc)
    results = matcher.match_all(bid_date, laws_list, progress_callback=progress_callback)

    generator = report.WordReportGenerator()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    return generator.generate(bid_date, results, out_path)

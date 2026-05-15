"""레거시 도구(12_manhour_aggregation, 24_crash_construction) 브릿지.

기존에 검증된 12 회사별 어댑터(4개 양식)와 24 v12 정교한 출력일보 코드를
다시 작성하지 않고, 38 패키지에서 일관된 인터페이스로 호출할 수 있도록 한다.
12·24 디렉토리명에 숫자가 있어 일반 ``import`` 가 불가능하므로 ``importlib`` 동적 로드.

지원하는 양식 / 도구
- 12_manhour_aggregation.readers.common: 금풍건설·기장원자로·㈜대형건설사1 xlsx
- 12_manhour_aggregation.readers.pdf_reader: 스마트에스지 PDF + 급여명세서 PDF
- 24_crash_construction.mandays_report_automation_v12: 정교한 출력일보
  (셀 스냅샷·머지셀·열너비·행높이·인쇄설정·페이지나눔·휴일 색상)

사용 예::

    from crash_aggregation import bridges

    # 12: 회사별 xlsx
    sheets = bridges.load_legacy_xlsx_reader("common").parse_workbook(path)

    # 24 v12: 정교한 출력일보 생성
    bridges.run_legacy_detailed_report(
        source_dir=Path("..."),
        output_dir=Path("..."),
    )
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType


# ── 경로 (프로젝트 루트 기준) ──────────────────────────────────────────────
def _project_root() -> Path:
    # 38_crash_aggregation/src/crash_aggregation/bridges.py 에서 3 단계 위
    return Path(__file__).resolve().parents[3]


def _load_module_from_path(name: str, file_path: Path) -> ModuleType:
    """importlib 으로 임의 경로의 .py 파일을 동적 import."""
    spec = importlib.util.spec_from_file_location(name, str(file_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load module from {file_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


# ── 12 회사별 어댑터 ─────────────────────────────────────────────────────
_LEGACY_12_READERS: dict[str, str] = {
    "common": "12_manhour_aggregation/readers/common.py",
    "pdf": "12_manhour_aggregation/readers/pdf_reader.py",
}


def load_legacy_xlsx_reader(variant: str = "common") -> ModuleType:
    """12_manhour_aggregation.readers.<variant> 모듈을 동적 로드.

    Args:
        variant: 'common' (xlsx: 금풍·기장·대우) | 'pdf' (스마트에스지·급여명세)

    Returns:
        로드된 모듈 객체. 기존의 detect_yearmonth(), find_date_header() 등 모두 호출 가능.
    """
    if variant not in _LEGACY_12_READERS:
        raise ValueError(f"variant 는 {list(_LEGACY_12_READERS)} 중 하나")
    path = _project_root() / _LEGACY_12_READERS[variant]
    if not path.exists():
        raise FileNotFoundError(f"레거시 12_manhour_aggregation 미설치: {path}")
    return _load_module_from_path(f"legacy_12_{variant}", path)


# ── 24 v12 정교한 출력일보 ─────────────────────────────────────────────────
_LEGACY_24_V12 = "24_crash_construction/mandays_report_automation_v12.py"


def load_legacy_detailed_report_module() -> ModuleType:
    """24_crash_construction v12 (1200줄) 모듈을 동적 로드.

    Returns:
        v12 모듈. ``main()``, ``build_section()``, ``snapshot_cells()`` 등 직접 호출.
    """
    path = _project_root() / _LEGACY_24_V12
    if not path.exists():
        raise FileNotFoundError(f"레거시 24_crash_construction 미설치: {path}")
    return _load_module_from_path("legacy_24_v12", path)


def run_legacy_detailed_report(
    source_dir: Path | None = None,
    output_dir: Path | None = None,
) -> Path:
    """24 v12 의 정교한 출력일보 (인쇄영역·머지셀·48행 페이지·휴일 표시) 생성.

    24 v12 의 main() 은 환경변수(SOURCE_DIR / OUTPUT_DIR)로 경로를 받음.
    인자가 주어지면 환경변수에 설정하고 main() 호출.

    Args:
        source_dir: 월별 xlsx 가 들어있는 폴더 (예: '24년06월.xlsx' 등)
        output_dir: 생성될 통합 출력일보 xlsx 가 저장될 폴더

    Returns:
        생성된 출력 xlsx 경로 (24 v12 의 main() 결과 추정).
    """
    import os

    if source_dir is not None:
        os.environ["SOURCE_DIR"] = str(source_dir.resolve())
    if output_dir is not None:
        os.environ["OUTPUT_DIR"] = str(output_dir.resolve())

    mod = load_legacy_detailed_report_module()
    if not hasattr(mod, "main"):
        raise AttributeError("24 v12 모듈에 main() 함수가 없습니다")
    return mod.main()

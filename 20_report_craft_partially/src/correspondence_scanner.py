"""
업체 수신자료 폴더 3-pass 공문 필터링

Pass 1 — 폴더명 기반 분류 (공문 관련 폴더 / 비관련 폴더)
Pass 2 — 파일 내용 확인 (공문 형식인지 첫 1~2페이지 검사)
Pass 3 — 귀책분석 관련성 필터 (키워드 매칭)

결과 저장:
  output/scan_folder_tree.md   — 전체 폴더 트리 및 분류 결과
  output/scan_candidates.json  — Pass2 통과 공문 후보 목록
  output/scan_result.json      — Pass3 확정 공문 목록 (편집 가능)
  output/correspondence_texts.md — 확정 공문 전문(全文)
"""

from __future__ import annotations

import json
import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import NamedTuple

from tqdm import tqdm

SCAN_WORKERS = 6  # Pass 2/3 병렬 처리 스레드 수

from .text_extractor import extract_file, extract_first_pages, ExtractResult


# ── 키워드 설정 (config에서 import) ─────────────────────────────────────────────

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))
import config


# ── 데이터 구조 ─────────────────────────────────────────────────────────────────

class CorrespondenceItem:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.folder_score: int = 0        # 폴더명 관련성 점수
        self.is_correspondence: bool = False
        self.is_relevant: bool = False

        # 공문 메타데이터 (Pass 2에서 추출)
        self.date: str = ""               # 발신일자 (파일명 날짜 우선, 없으면 OCR)
        self.filename_date: str = ""      # 파일명에서 추출한 날짜 (YYMMDD/YYYYMMDD 접두사)
        self.ocr_date: str = ""           # OCR 추출 원본 날짜 (불일치 경고용)
        self.doc_number: str = ""         # 공문번호
        self.sender: str = ""             # 발신처
        self.receiver: str = ""           # 수신처
        self.subject: str = ""            # 제목
        self.summary: str = ""            # 핵심내용(1줄)

        # 전문(全文) — Pass 3 + generate 단계에서 사용
        self.full_text: str = ""

        # Pass 3 부가 정보
        self.matched_keywords: list[str] = []   # 매칭된 키워드
        self.borderline: bool = False            # True = 탈락했지만 검토 필요
        self.ocr_quality: str = "OK"            # OK / WARN — 텍스트 추출 품질

    def to_dict(self) -> dict:
        return {
            "file_path": str(self.file_path),
            "mtime": _mtime(self.file_path),
            "date": self.date,
            "filename_date": self.filename_date,
            "ocr_date": self.ocr_date,
            "doc_number": self.doc_number,
            "sender": self.sender,
            "receiver": self.receiver,
            "subject": self.subject,
            "summary": self.summary,
            "folder_score": self.folder_score,
            "is_correspondence": self.is_correspondence,
            "is_relevant": self.is_relevant,
            "full_text": self.full_text,
            "matched_keywords": self.matched_keywords,
            "borderline": self.borderline,
            "ocr_quality": self.ocr_quality,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "CorrespondenceItem":
        item = cls(Path(d["file_path"]))
        item.folder_score     = d.get("folder_score", 0)
        item.is_correspondence = d.get("is_correspondence", False)
        item.is_relevant      = d.get("is_relevant", False)
        item.date             = d.get("date", "")
        item.filename_date    = d.get("filename_date", "")
        item.ocr_date         = d.get("ocr_date", "")
        item.doc_number       = d.get("doc_number", "")
        item.sender           = d.get("sender", "")
        item.receiver         = d.get("receiver", "")
        item.subject          = d.get("subject", "")
        item.summary           = d.get("summary", "")
        item.full_text         = d.get("full_text", "")
        item.matched_keywords  = d.get("matched_keywords", [])
        item.borderline        = d.get("borderline", False)
        item.ocr_quality       = d.get("ocr_quality", "OK")
        return item


# ── 공문 메타데이터 추출 (첫 페이지 텍스트 → 정규식) ─────────────────────────────

_DATE_RE = re.compile(
    r"(\d{4}[\.\-/]\d{1,2}[\.\-/]\d{1,2})|"   # 2024.03.15
    r"(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)|"      # 2024년 3월 15일
    r"'(\d{2}[\.\-/]\d{1,2}[\.\-/]\d{1,2})"   # '22.08.11 (약식)
)
_DOCNUM_RE = re.compile(
    r"[가-힣A-Za-z]{1,10}[-\s]?\d{2,4}[-\s]?\d{2,6}|"  # 한글-숫자-숫자
    r"[A-Z]{2,6}-\d{4}-\d{3,6}|"                         # 영문-숫자-숫자
    r"제\s*\d+\s*호"                                        # 제N호
)
_RECEIVER_RE = re.compile(r"수\s*신\s*[:：]?\s*(.{2,30})")
_SENDER_RE   = re.compile(r"발\s*신\s*[:：]?\s*(.{2,30})|"
                           r"기\s*안\s*[:：]?\s*(.{2,30})")
_SUBJECT_RE  = re.compile(r"(?:제\s*목|제\s*:)\s*[:：]?\s*(.{3,80})")


_FILENAME_DATE_RE = re.compile(r"(\d{4}[\.\-/]\d{1,2}[\.\-/]\d{1,2})|(\d{6,8})")
_FILENAME_DOCNUM_RE = re.compile(r"제(\d{4}-\d{1,3})호|[A-Za-z가-힣]+-\d{3,6}|\[공문-\d+\]")


def _parse_meta(text: str, filename: str = "") -> dict:
    """첫 페이지 텍스트에서 공문 메타데이터 추출. 파일명 날짜를 우선 적용.
    반환 dict에 'filename_date' 키 포함 (파일명에서 추출한 날짜, 없으면 빈 문자열).
    """
    meta = {"date": "", "filename_date": "", "ocr_date": "", "doc_number": "", "sender": "", "receiver": "", "subject": ""}

    # ── 파일명에서 날짜 미리 추출 (OCR보다 신뢰도 높음) ──────────────────────────
    # 예) "220728 공기연장..." → 2022.07.28  /  "20220728_공문" → 2022.07.28
    filename_date = ""
    if filename:
        stem = Path(filename).stem
        m = _FILENAME_DATE_RE.search(stem)
        if m:
            raw = m.group(0)
            if raw.isdigit() and len(raw) == 8:
                filename_date = f"{raw[:4]}.{raw[4:6]}.{raw[6:]}"
            elif raw.isdigit() and len(raw) == 6:
                filename_date = f"20{raw[:2]}.{raw[2:4]}.{raw[4:]}"
            else:
                filename_date = raw
    meta["filename_date"] = filename_date

    # ── OCR 추출 ──────────────────────────────────────────────────────────────
    m = _DATE_RE.search(text)
    if m:
        ocr_date = m.group(0).strip()
        meta["date"] = ocr_date
        meta["ocr_date"] = ocr_date  # 원본 OCR 날짜 보존 (불일치 경고용)

    # 문서번호: 공문 헤더(앞 600자)에서만 검색 → 본문에 인용된 타 공문번호 오인식 방지
    header_text = text[:600]
    m = _DOCNUM_RE.search(header_text)
    if m:
        meta["doc_number"] = m.group(0).strip()

    m = _RECEIVER_RE.search(text)
    if m:
        meta["receiver"] = m.group(1).strip()[:30]

    m = _SENDER_RE.search(text)
    if m:
        meta["sender"] = (m.group(1) or m.group(2) or "").strip()[:30]

    m = _SUBJECT_RE.search(text)
    if m:
        meta["subject"] = m.group(1).strip()[:60]

    # ── 파일명으로 교정 / 보완 ────────────────────────────────────────────────
    if filename:
        stem = Path(filename).stem

        # 날짜: 파일명 날짜가 있으면 항상 우선 사용
        # (OCR은 본문 안에 인용된 다른 공문 날짜를 잡는 경우가 많음)
        if filename_date:
            meta["date"] = filename_date

        if not meta["doc_number"]:
            m = _FILENAME_DOCNUM_RE.search(stem)
            if m:
                meta["doc_number"] = m.group(0).strip()

        if not meta["subject"]:
            # 파일명에서 번호/날짜 등 제거하고 제목으로 사용
            subject = re.sub(r"^\d+\.", "", stem).strip()
            subject = re.sub(r"\(.*?\)|\[.*?\]", "", subject).strip()
            meta["subject"] = subject[:60]

    return meta


def _is_correspondence_text(text: str, min_score: int = 2) -> bool:
    """첫 페이지 텍스트가 공문 형식인지 판단"""
    required = 0
    if _RECEIVER_RE.search(text):
        required += 1
    if _SENDER_RE.search(text):
        required += 1
    if _SUBJECT_RE.search(text):
        required += 1
    if _DATE_RE.search(text):
        required += 1
    return required >= min_score


# ── 캐시 (파일별 처리 결과 저장 → 재실행 시 건너뜀) ──────────────────────────────

def _mtime(path: Path) -> float:
    """파일 수정 시각 (변경 여부 판단용)"""
    try:
        return path.stat().st_mtime
    except OSError:
        return 0.0


# Pass 3 로직이 변경될 때마다 이 값을 올린다 → 이전 캐시 자동 무효화
_CACHE_VERSION = 2  # v2: Pass 3에서 subject+파일명 검색 추가 (2026-04-16)


def _load_cache(cache_path: Path) -> dict[str, dict]:
    """
    캐시 파일 로드. _CACHE_VERSION 불일치 시 빈 캐시 반환 (전체 재스캔).
    키: 절대 파일 경로 문자열
    값: to_dict() 결과 + mtime
    """
    if not cache_path.exists():
        return {}
    try:
        data = json.loads(cache_path.read_text(encoding="utf-8"))
        if data.get("__version__") != _CACHE_VERSION:
            print(f"  [캐시] 버전 불일치 (저장={data.get('__version__')} / 현재={_CACHE_VERSION}) → 전체 재스캔")
            return {}
        return {k: v for k, v in data.items() if k != "__version__"}
    except Exception:
        return {}


def _save_cache(cache_path: Path, cache: dict[str, dict]):
    data = {"__version__": _CACHE_VERSION}
    data.update(cache)
    cache_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _cache_hit(cache: dict[str, dict], file_path: Path) -> CorrespondenceItem | None:
    """
    캐시에 해당 파일이 있고 mtime이 동일하면 캐시 결과 반환.
    파일이 수정됐거나 캐시에 없으면 None 반환.
    """
    key = str(file_path.resolve())
    entry = cache.get(key)
    if entry is None:
        return None
    if abs(entry.get("mtime", -1) - _mtime(file_path)) > 1:  # 1초 허용 오차
        return None
    return CorrespondenceItem.from_dict(entry)


# ── Pass 1: 폴더명 기반 분류 ──────────────────────────────────────────────────────

# 모듈 로드 시 1회만 소문자화 (매 폴더·공문마다 반복 연산 방지)
_FOLDER_KW_LOWER    = [kw.lower() for kw in config.CORRESPONDENCE_FOLDER_KEYWORDS]
_SKIP_KW_LOWER      = [kw.lower() for kw in config.SKIP_FOLDER_KEYWORDS]
_RELEVANCE_KW_LOWER = [kw.lower() for kw in config.RELEVANCE_KEYWORDS]
_BORDERLINE_NEG_KW_LOWER = [kw.lower() for kw in config.BORDERLINE_NEGATIVE_KEYWORDS]

_SCAN_SUPPORTED_EXTS = frozenset({
    ".docx", ".pdf", ".hwp", ".hwpx",
    ".jpg", ".jpeg", ".png", ".tif", ".tiff"
})


def _score_folder(folder: Path) -> int:
    """폴더명의 공문 관련도 점수 반환 (양수=포함, 음수=제외)"""
    name = folder.name.lower()
    for kw in _FOLDER_KW_LOWER:
        if kw in name:
            return 1   # 공문 관련 폴더
    for kw in _SKIP_KW_LOWER:
        if kw in name:
            return -1  # 제외 폴더
    return 0   # 판단 불가 (중립)


def _scan_tree(root: Path) -> tuple[list[tuple[Path, int]], list[tuple[Path, int]]]:
    """
    한 번의 rglob으로 폴더 점수 계산 + 파일 목록 수집.
    (기존 _build_folder_tree + _collect_candidate_files 통합 → I/O 절반)

    반환: (folder_scores, candidate_files)
      folder_scores  : [(폴더경로, 점수), ...]
      candidate_files: [(파일경로, 폴더점수), ...] — 점수 내림차순
    """
    folder_scores: list[tuple[Path, int]] = []
    score_map: dict[Path, int] = {}
    raw_files: list[Path] = []

    # 정렬된 rglob: 부모 디렉터리가 항상 자식보다 먼저 나오므로
    # 파일을 만날 때 score_map 이 이미 완성돼 있다.
    for p in sorted(root.rglob("*")):
        if p.is_dir():
            score = _score_folder(p)
            folder_scores.append((p, score))
            score_map[p] = score
        elif p.is_file() and p.suffix.lower() in _SCAN_SUPPORTED_EXTS:
            raw_files.append(p)

    def _folder_score_for(file_path: Path) -> int:
        """파일의 부모 폴더 중 가장 관련성 높은 점수 반환.

        가까운 조상 폴더가 양성(+1)이면 더 먼 조상의 음성(-1)을 무시한다.
        예) '공기연장 실정보고자료'(+1) 안의 파일이
            '250613_간접비 산출 증빙(…노무비, 경비…)'(-1) 안에 있어도 포함.
        """
        best = 0
        for parent in file_path.parents:
            if parent == root:
                break
            s = score_map.get(parent, 0)
            if s > best:
                best = s
            elif s < 0:
                if best > 0:
                    # 더 가까운 조상이 이미 양성 → 먼 음성 조상 무시
                    continue
                return -1  # 양성 신호 없이 음성 폴더 → 제외
        return best

    candidate_files = sorted(
        [(p, _folder_score_for(p)) for p in raw_files],
        key=lambda x: -x[1],
    )
    return folder_scores, candidate_files


# ── Pass 2: 공문 여부 확인 ─────────────────────────────────────────────────────

_LENIENT_ROOT_KW = re.compile(r"공기연장|공사기간|분쟁|귀책|간접비|클레임", re.IGNORECASE)


def _pass2_check(
    file_path: Path,
    folder_score: int,
    root: Path | None = None,
) -> CorrespondenceItem | None:
    """
    첫 1~2페이지를 읽어 공문인지 확인.
    폴더점수 < 0 이면 건너뜀.
    root 폴더명이 공기연장·간접비·귀책 관련이면 판단 기준 완화 (2→1).
    """
    if folder_score < 0:
        return None

    item = CorrespondenceItem(file_path)
    item.folder_score = folder_score

    result = extract_first_pages(file_path, max_chars=2000)
    if result.quality == "FAIL" or not result.text.strip():
        return None

    # 스캔 루트 폴더명이 관련 키워드를 포함하면 형식 검사 생략
    lenient = root is not None and bool(_LENIENT_ROOT_KW.search(root.name))
    # 명시적 공문 폴더(folder_score>=1)에 있는 파일은 판별 기준을 1로 완화
    # (OCR 추출 시 레이블이 잘려도 날짜 하나만 있어도 통과)
    min_corr_score = 1 if folder_score >= 1 else 2
    if not lenient and not _is_correspondence_text(result.text, min_score=min_corr_score):
        return None

    item.is_correspondence = True
    meta = _parse_meta(result.text, filename=file_path.name)
    item.date = meta["date"]
    item.filename_date = meta["filename_date"]
    item.ocr_date = meta["ocr_date"]
    item.doc_number = meta["doc_number"]
    item.sender = meta["sender"]
    item.receiver = meta["receiver"]
    item.subject = meta["subject"]

    # 요약: 제목이 없으면 첫 줄 사용
    item.summary = item.subject or result.text.split("\n")[0][:60]

    return item


# ── 공통 유틸 ────────────────────────────────────────────────────────────────────

def _is_under(file_path: Path, directory: Path) -> bool:
    """file_path 가 directory 하위에 있는지 확인."""
    try:
        file_path.relative_to(directory)
        return True
    except ValueError:
        return False


def _display_path(file_path: Path, vendor_dirs: list[Path]) -> str:
    """파일 경로를 vendor_dirs 중 하나를 기준으로 상대경로로 반환."""
    for vd in vendor_dirs:
        try:
            return str(file_path.relative_to(vd))
        except ValueError:
            pass
    return str(file_path)


# ── Pass 3: 귀책분석 관련성 필터 ─────────────────────────────────────────────────

def _pass3_relevance(item: CorrespondenceItem, full_text: str) -> tuple[bool, list[str]]:
    """
    공문 전문 + 제목 + 파일명에서 귀책분석 관련 키워드 확인.
    전문 OCR 추출이 실패해도 제목/파일명으로 관련성 판단 가능.
    반환: (관련여부, 매칭된 키워드 목록)
    """
    # 전문(全文) 검색
    text_lower = full_text.lower()
    matched_set: set[str] = {
        config.RELEVANCE_KEYWORDS[i]
        for i, kw in enumerate(_RELEVANCE_KW_LOWER)
        if kw in text_lower
    }

    # 제목 + 파일명 보조 검색 (OCR 품질 불량 시 보완)
    title_lower = (item.subject + " " + item.file_path.stem).lower()
    for kw_orig, kw_lo in zip(config.RELEVANCE_KEYWORDS, _RELEVANCE_KW_LOWER):
        if kw_lo in title_lower:
            matched_set.add(kw_orig)

    matched = sorted(matched_set, key=lambda k: config.RELEVANCE_KEYWORDS.index(k))
    return bool(matched), matched


def _reclassify_borderline(
    items: list[CorrespondenceItem],
) -> tuple[list[CorrespondenceItem], list[CorrespondenceItem], list[CorrespondenceItem]]:
    """
    borderline 항목을 제목·파일명 기준으로 재분류.

    - promoted      : RELEVANCE_KEYWORDS 제목 매칭 → confirmed 승격
    - remaining     : 판단 불가 → borderline 유지
    - auto_excluded : BORDERLINE_NEGATIVE_KEYWORDS 매칭 → scan_result.json 제외

    이 함수는 _pass3_relevance 에서 이미 전문·제목 검사를 모두 통과하지 못한
    항목만 인수로 받아야 합니다. (promoted 케이스 대부분은 Pass 3 개선으로 흡수됨)
    BORDERLINE_NEGATIVE_KEYWORDS 를 이용한 명확한 무관 항목만 자동 제외합니다.
    """
    promoted: list[CorrespondenceItem] = []
    remaining: list[CorrespondenceItem] = []
    auto_excluded: list[CorrespondenceItem] = []

    for it in items:
        # 제목·파일명 소문자화 (폴더 경로 일부도 포함)
        check = (it.subject + " " + it.file_path.stem + " "
                 + " ".join(it.file_path.parts[-3:])).lower()

        # ── 긍정 재검사: RELEVANCE_KEYWORDS 가 제목에 있으면 confirmed 승격 ────
        rel_matched = [
            kw_orig for kw_orig, kw_lo in zip(config.RELEVANCE_KEYWORDS, _RELEVANCE_KW_LOWER)
            if kw_lo in check
        ]
        if rel_matched:
            it.matched_keywords = rel_matched
            it.is_relevant = True
            it.borderline = False
            promoted.append(it)
            continue

        # ── 부정 재검사: BORDERLINE_NEGATIVE_KEYWORDS 매칭 → 자동 제외 ─────────
        if any(nk in check for nk in _BORDERLINE_NEG_KW_LOWER):
            auto_excluded.append(it)
            continue

        remaining.append(it)

    return promoted, remaining, auto_excluded


# ── 단일 폴더 스캔 (내부용) ────────────────────────────────────────────────────

def _scan_one(
    vendor_dir: Path,
    cache: dict[str, dict],
) -> tuple[list[CorrespondenceItem], list[CorrespondenceItem], list[CorrespondenceItem], list[str], int, list]:
    """
    단일 폴더에 대해 Pass 1~3 실행.
    cache: 이미 처리한 파일 결과 (mtime 기반 — 변경 없으면 재처리 생략)
    반환: (correspondence_items, relevant_items, failed_reads, total_files, folder_scores)
    """
    print(f"\n  [Pass 1] 폴더 트리: {vendor_dir}")
    folder_scores, candidate_files = _scan_tree(vendor_dir)

    pos = sum(1 for _, s in folder_scores if s > 0)
    neg = sum(1 for _, s in folder_scores if s < 0)
    neu = sum(1 for _, s in folder_scores if s == 0)
    print(f"    공문 관련: {pos}개 | 중립: {neu}개 | 제외: {neg}개")

    total_files = len(candidate_files)
    print(f"    전체 파일: {total_files}개")

    # Pass 2 — 캐시 히트 시 재처리 생략 (병렬 처리)
    correspondence_items: list[CorrespondenceItem] = []
    cache_lock = threading.Lock()
    skipped_neg = skipped_cache_p2 = 0

    # 제외 폴더(-score) 먼저 분리
    to_process = []
    for file_path, folder_score in candidate_files:
        if folder_score < 0:
            skipped_neg += 1
        else:
            to_process.append((file_path, folder_score))

    def _p2_worker(file_path: Path, folder_score: int):
        with cache_lock:
            cached = _cache_hit(cache, file_path)
        if cached is not None:
            return "cache", cached
        item = _pass2_check(file_path, folder_score, root=vendor_dir)
        key = str(file_path.resolve())
        if item:
            with cache_lock:
                cache[key] = item.to_dict()
            return "hit", item
        else:
            dummy = CorrespondenceItem(file_path)
            dummy.folder_score = folder_score
            with cache_lock:
                cache[key] = dummy.to_dict()
            return "miss", None

    with tqdm(total=len(to_process), desc="  Pass 2 공문확인", unit="파일", leave=True) as bar:
        with ThreadPoolExecutor(max_workers=SCAN_WORKERS) as ex:
            futures = {ex.submit(_p2_worker, fp, fs): (fp, fs) for fp, fs in to_process}
            for fut in as_completed(futures):
                result, item = fut.result()
                if result == "cache":
                    skipped_cache_p2 += 1
                    if item and item.is_correspondence:
                        correspondence_items.append(item)
                elif result == "hit":
                    correspondence_items.append(item)
                bar.set_postfix(공문=len(correspondence_items), 캐시=skipped_cache_p2)
                bar.update(1)

    if skipped_cache_p2:
        tqdm.write(f"    캐시 재사용: {skipped_cache_p2}개 건너뜀")
    print(f"    공문 확인: {len(correspondence_items)}개 (제외 폴더 {skipped_neg}개 건너뜀)")

    # Pass 3 — 캐시에 full_text 있으면 재처리 생략 (병렬 처리)
    relevant_items: list[CorrespondenceItem] = []
    borderline_items: list[CorrespondenceItem] = []
    failed_reads: list[str] = []
    skipped_cache_p3 = 0

    def _p3_worker(item: CorrespondenceItem):
        with cache_lock:
            cached = _cache_hit(cache, item.file_path)
        if cached is not None and cached.full_text:
            return "cache", cached, None
        full = extract_file(item.file_path)
        if full.quality == "FAIL" or not full.text.strip():
            return "fail", item, None
        is_relevant, matched = _pass3_relevance(item, full.text)
        item.full_text = full.text
        item.matched_keywords = matched
        item.is_relevant = is_relevant
        item.borderline = not is_relevant
        item.ocr_quality = full.quality  # OK / WARN
        with cache_lock:
            cache[str(item.file_path.resolve())] = item.to_dict()
        return "done", item, is_relevant

    with tqdm(total=len(correspondence_items), desc="  Pass 3 관련성", unit="공문", leave=True) as bar:
        with ThreadPoolExecutor(max_workers=SCAN_WORKERS) as ex:
            futures = {ex.submit(_p3_worker, item): item for item in correspondence_items}
            for fut in as_completed(futures):
                status, item, is_relevant = fut.result()
                if status == "cache":
                    skipped_cache_p3 += 1
                    if item.is_relevant:
                        relevant_items.append(item)
                    elif item.borderline:
                        borderline_items.append(item)
                elif status == "fail":
                    failed_reads.append(item.file_path.name)
                elif status == "done":
                    if is_relevant:
                        relevant_items.append(item)
                    else:
                        borderline_items.append(item)
                bar.set_postfix(관련=len(relevant_items), 경계=len(borderline_items), 캐시=skipped_cache_p3)
                bar.update(1)

    if skipped_cache_p3:
        tqdm.write(f"    캐시 재사용: {skipped_cache_p3}개 건너뜀")
    print(f"    관련 공문: {len(relevant_items)}개 | 경계선 검토 필요: {len(borderline_items)}개")

    return correspondence_items, relevant_items, borderline_items, failed_reads, total_files, folder_scores


# ── 전체 스캔 실행 ────────────────────────────────────────────────────────────────

def scan(vendor_dirs: list[Path], output_dir: Path) -> Path:
    """
    3-pass 스캔 실행 후 결과 파일 저장.
    vendor_dirs: 스캔할 폴더 경로 목록 (1개 이상)
    반환: 편집 가능한 scan_result.json 경로
    """
    # GUI 모드(동일 프로세스 반복 실행) 시 config 변경이 반영되도록 키워드 갱신
    global _FOLDER_KW_LOWER, _SKIP_KW_LOWER, _RELEVANCE_KW_LOWER, _BORDERLINE_NEG_KW_LOWER
    _FOLDER_KW_LOWER       = [kw.lower() for kw in config.CORRESPONDENCE_FOLDER_KEYWORDS]
    _SKIP_KW_LOWER         = [kw.lower() for kw in config.SKIP_FOLDER_KEYWORDS]
    _RELEVANCE_KW_LOWER    = [kw.lower() for kw in config.RELEVANCE_KEYWORDS]
    _BORDERLINE_NEG_KW_LOWER = [kw.lower() for kw in config.BORDERLINE_NEGATIVE_KEYWORDS]

    output_dir.mkdir(parents=True, exist_ok=True)
    vendor_dirs = [d.resolve() for d in vendor_dirs]

    # ── 캐시 로드 ──────────────────────────────────────────────────────────────
    cache_path = output_dir / "scan_cache.json"
    cache = _load_cache(cache_path)
    cache_size = len(cache)
    if cache_size:
        print(f"\n[캐시] {cache_path.name} 에서 {cache_size}개 항목 로드 (변경된 파일만 재스캔)")
    else:
        print(f"\n[캐시] 없음 — 전체 스캔")

    print(f"\n[스캔 대상 경로 — {len(vendor_dirs)}개]")
    for i, d in enumerate(vendor_dirs, 1):
        print(f"  [{i}] {d}")
    print("─" * 60)

    # ── 각 경로별 스캔 ──────────────────────────────────────────────────────────
    all_correspondence: list[CorrespondenceItem] = []
    all_relevant: list[CorrespondenceItem] = []
    all_borderline: list[CorrespondenceItem] = []
    all_failed: list[str] = []
    total_files_all = 0
    per_dir_stats: list[dict] = []
    tree_lines_all = ["# 폴더 트리 분류 결과", ""]

    seen_paths: set[Path] = set()

    for dir_idx, vendor_dir in enumerate(vendor_dirs, 1):
        print(f"\n━━ 경로 [{dir_idx}/{len(vendor_dirs)}]: {vendor_dir.name}")
        corr, rel, borderline, failed, total, folder_scores = _scan_one(vendor_dir, cache)

        tree_lines_all += [f"## [{dir_idx}] {vendor_dir}", "", "| 분류 | 폴더 경로 |", "|------|-----------|"]
        for folder, score in folder_scores:
            rel_path = folder.relative_to(vendor_dir)
            label = "✅ 공문 관련" if score > 0 else ("❌ 제외" if score < 0 else "⬜ 중립")
            tree_lines_all.append(f"| {label} | {rel_path} |")
        tree_lines_all.append("")

        new_corr       = [i for i in corr       if i.file_path not in seen_paths]
        new_rel        = [i for i in rel        if i.file_path not in seen_paths]
        new_borderline = [i for i in borderline if i.file_path not in seen_paths]
        for i in new_rel + new_borderline:
            seen_paths.add(i.file_path)

        all_correspondence.extend(new_corr)
        all_relevant.extend(new_rel)
        all_borderline.extend(new_borderline)
        all_failed.extend(failed)
        total_files_all += total

        per_dir_stats.append({
            "dir": str(vendor_dir),
            "total_files": total,
            "correspondence_found": len(new_corr),
            "relevant_confirmed": len(new_rel),
            "borderline": len(new_borderline),
            "failed_reads": failed,
        })

    # 날짜 기준 정렬
    def _sort_key(x: CorrespondenceItem):
        d = re.sub(r"[^\d]", "", x.date)
        return d if len(d) >= 8 else "99999999"

    all_relevant.sort(key=_sort_key)
    all_borderline.sort(key=_sort_key)

    # ── borderline 자동 재분류 ─────────────────────────────────────────────────
    promoted, all_borderline, auto_excluded = _reclassify_borderline(all_borderline)
    all_relevant.extend(promoted)
    all_relevant.sort(key=_sort_key)

    if promoted:
        print(f"  [재분류] borderline → confirmed 승격: {len(promoted)}개 (제목 키워드 매칭)")
    if auto_excluded:
        print(f"  [재분류] 자동 제외: {len(auto_excluded)}개 (BORDERLINE_NEGATIVE_KEYWORDS 매칭)")

    # per_dir_stats 소급 보정: 재분류 결과 반영
    # (per_dir_stats는 각 디렉터리 스캔 직후 수집되어 promoted/auto_excluded 미반영)
    promoted_set = {i.file_path for i in promoted}
    excluded_set = {i.file_path for i in auto_excluded}
    for stat in per_dir_stats:
        stat_dir = Path(stat["dir"])
        stat_promoted = sum(
            1 for fp in promoted_set
            if _is_under(fp, stat_dir)
        )
        stat_excluded = sum(
            1 for fp in excluded_set
            if _is_under(fp, stat_dir)
        )
        stat["relevant_confirmed"] += stat_promoted
        stat["borderline"] -= (stat_promoted + stat_excluded)
        stat["auto_excluded"] = stat_excluded

    pct = len(all_relevant) / max(total_files_all, 1) * 100
    print(f"\n[합산 결과]")
    print(f"  전체 파일:      {total_files_all}개")
    print(f"  공문 확인:      {len(all_correspondence)}개")
    print(f"  포함 확정:      {len(all_relevant)}개 (전체 대비 {pct:.1f}%)")
    print(f"    (승격 포함:   {len(promoted)}개 제목 매칭으로 확정)")
    print(f"  수동 검토 필요: {len(all_borderline)}개  ← scan_borderline.md 확인")
    print(f"  자동 제외:      {len(auto_excluded)}개  ← scan_borderline.md 하단 참조")
    print(f"  명확히 제외:    {len(all_correspondence) - len(all_relevant) - len(all_borderline) - len(auto_excluded)}개")

    # 폴더 트리 저장
    (output_dir / "scan_folder_tree.md").write_text(
        "\n".join(tree_lines_all), encoding="utf-8"
    )

    # 캐시 저장 (다음 실행 시 재사용)
    _save_cache(cache_path, cache)
    print(f"  캐시 저장: {len(cache)}개 항목 → {cache_path.name}")

    # ── 결과 저장 ──────────────────────────────────────────────────────────────
    # confirmed + borderline 을 scan_result.json 에 포함
    # relevance 필드: "confirmed" | "confirmed-by-title" | "borderline"
    def _with_relevance(item: CorrespondenceItem, relevance: str) -> dict:
        d = item.to_dict()
        d["relevance"] = relevance
        return d

    # promoted 항목에는 "confirmed-by-title" 태그 부여 (추적용)
    promoted_paths = {i.file_path for i in promoted}

    def _relevance_tag(item: CorrespondenceItem) -> str:
        if item.file_path in promoted_paths:
            return "confirmed-by-title"
        return "confirmed"

    all_items_combined = (
        [_with_relevance(i, _relevance_tag(i)) for i in all_relevant] +
        [_with_relevance(i, "borderline") for i in all_borderline]
    )
    # 날짜 기준 재정렬 (confirmed/borderline 혼합 후 정렬)
    def _sort_key_d(d: dict) -> str:
        raw = re.sub(r"[^\d]", "", d.get("date", ""))
        return raw if len(raw) >= 8 else "99999999"
    all_items_combined.sort(key=_sort_key_d)

    # scan_no 부여: Claude가 data.json items에 복사할 번호 (1-based)
    for idx, item_dict in enumerate(all_items_combined, 1):
        item_dict["no"] = idx

    result_data = {
        "vendor_dirs": [str(d) for d in vendor_dirs],
        "stats": {
            "total_files": total_files_all,
            "correspondence_found": len(all_correspondence),
            "relevant_confirmed": len(all_relevant),
            "promoted_by_title": len(promoted),
            "borderline_remaining": len(all_borderline),
            "auto_excluded": len(auto_excluded),
            "relevant_pct": round(pct, 1),
            "failed_reads": all_failed,
            "per_dir": per_dir_stats,
        },
        "items": all_items_combined,
    }
    result_path = output_dir / "scan_result.json"
    result_path.write_text(
        json.dumps(result_data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # 사람이 읽기 쉬운 마크다운 요약
    summary_lines = [
        "# 스캔 결과 요약",
        "",
        "## 통계",
        "",
        f"- 스캔 경로 수: {len(vendor_dirs)}개",
        f"- 전체 파일 수: {total_files_all}개",
        f"- 공문 확인 파일: {len(all_correspondence)}개",
        f"- 포함 확정(귀책 키워드 매칭): **{len(all_relevant)}개** (전체 대비 {pct:.1f}%)",
        f"  - 전문(全文) 매칭: {len(all_relevant) - len(promoted)}개",
        f"  - 제목·파일명 매칭(승격): {len(promoted)}개",
        f"- 수동 검토 필요(키워드 미매칭): **{len(all_borderline)}개** → scan_borderline.md 참조",
        f"  ※ 수동 검토 항목도 scan_result.json에 포함됩니다. 불필요한 항목은 삭제하세요.",
        f"- 자동 제외(무관 키워드 매칭): **{len(auto_excluded)}개** → scan_borderline.md 하단 참조",
        f"- 명확히 제외: {len(all_correspondence) - len(all_relevant) - len(all_borderline) - len(auto_excluded)}개",
        "",
    ]
    # 경로별 소계
    if len(vendor_dirs) > 1:
        summary_lines += ["### 경로별 소계", ""]
        for stat in per_dir_stats:
            summary_lines.append(
                f"- `{stat['dir']}`  "
                f"파일 {stat['total_files']}개 / 공문 {stat['correspondence_found']}개 "
                f"/ 포함 {stat['relevant_confirmed']}개 / 경계선 {stat['borderline']}개"
            )
        summary_lines.append("")

    # OCR 품질 WARN 항목 수집
    all_ocr_warn = [i for i in all_relevant + all_borderline if i.ocr_quality == "WARN"]
    if all_ocr_warn:
        summary_lines += [
            "### ⚠️ OCR 품질 주의 항목",
            "",
            "아래 파일은 텍스트 추출 품질이 낮습니다. 공문번호·날짜·발신처를 직접 확인하세요.",
            "",
            *[f"- `{_display_path(i.file_path, vendor_dirs)}` — {i.subject or i.file_path.name}" for i in all_ocr_warn],
            "",
        ]

    if all_failed:
        summary_lines += [
            "### [읽기 실패]",
            "",
            *[f"- {f}" for f in all_failed],
            "",
        ]
    # file_path → scan_no 맵: scan_result.json의 실제 번호와 일치하도록
    _path_to_scanno: dict[str, int] = {d["file_path"]: d["no"] for d in all_items_combined}

    summary_lines += [
        "## 관련 공문 목록 — 발신일자순",
        "",
        "※ No 컬럼은 scan_result.json의 scan_no와 동일합니다. Claude에게 전달할 때 그대로 사용하세요.",
        "※ 발신일자 뒤 `⚠️(OCR:날짜)` 표시는 파일명 날짜와 OCR 추출 날짜가 다름을 의미합니다.",
        "  → 파일명 날짜가 적용되었습니다. 공문번호·발신처도 확인이 필요할 수 있습니다.",
        "",
        "| No | 발신일자 | 공문번호 | 발신처 | 수신처 | 제목 | 핵심내용 | 파일경로 |",
        "|-----|---------|---------|--------|--------|------|---------|---------|",
    ]
    for item in all_relevant:
        no = _path_to_scanno.get(str(item.file_path), "?")
        display_path = _display_path(item.file_path, vendor_dirs)
        # 파일명 날짜와 OCR 날짜가 다르면 경고 표시 (공문번호·발신처도 확인 필요)
        date_display = item.date
        if item.filename_date and item.ocr_date and item.filename_date != item.ocr_date:
            date_display = f"{item.date} ⚠️OCR={item.ocr_date}"
        summary_lines.append(
            f"| {no} | {date_display} | {item.doc_number} | {item.sender} | "
            f"{item.receiver} | {item.subject} | {item.summary} | {display_path} |"
        )
    summary_lines += [
        "",
        "---",
        "",
        "**다음 단계**:",
        "1. (선택) `scan_borderline.md` 를 열어 참고 항목을 확인하세요.",
        "   - 참고 항목은 이미 scan_result.json에 포함되어 있습니다.",
        "   - 귀책분석와 무관한 항목은 scan_result.json에서 해당 항목을 삭제하세요.",
        "2. `python main.py prepare` 를 실행하세요. (scan_result.json 편집 없이 바로 실행 가능)",
    ]
    (output_dir / "scan_summary.md").write_text(
        "\n".join(summary_lines), encoding="utf-8"
    )

    # scan_borderline.md — 수동 검토 필요 항목 + 자동 제외 항목 목록
    borderline_lines = [
        "# 경계선 공문 검토 목록",
        "",
        "## 수동 검토 필요 항목 (scan_result.json에 포함됨)",
        "",
        "귀책 키워드 미매칭이지만 공문 형식을 갖춘 파일 목록입니다.",
        "**이 항목들은 scan_result.json에 자동 포함되어 있습니다.**",
        "귀책분석와 명백히 무관한 항목만 scan_result.json에서 직접 삭제하세요.",
        "",
        f"총 {len(all_borderline)}개",
        "",
    ]
    if all_borderline:
        borderline_lines += [
            "| No | 발신일자 | 공문번호 | 발신처 | 수신처 | 제목 | 파일경로 |",
            "|----|---------|---------|--------|--------|------|---------|",
        ]
        for no, item in enumerate(all_borderline, 1):
            display_path = _display_path(item.file_path, vendor_dirs)
            borderline_lines.append(
                f"| {no} | {item.date} | {item.doc_number} | {item.sender} | "
                f"{item.receiver} | {item.subject} | {display_path} |"
            )
    else:
        borderline_lines.append("_(없음)_")

    # 자동 제외 항목 섹션
    borderline_lines += [
        "",
        "---",
        "",
        "## 자동 제외 항목 (scan_result.json에서 제외됨)",
        "",
        "BORDERLINE_NEGATIVE_KEYWORDS(식대·요금고지서 등) 매칭으로 자동 제외된 항목입니다.",
        "scan_result.json에 포함되지 않습니다.",
        "잘못 제외된 항목이 있으면 scan_result.json에 수동으로 추가하거나",
        "config.py의 BORDERLINE_NEGATIVE_KEYWORDS 에서 해당 키워드를 제거하세요.",
        "",
        f"총 {len(auto_excluded)}개",
        "",
    ]
    if auto_excluded:
        borderline_lines += [
            "| No | 발신일자 | 공문번호 | 발신처 | 수신처 | 제목 | 파일경로 |",
            "|----|---------|---------|--------|--------|------|---------|",
        ]
        for no, item in enumerate(auto_excluded, 1):
            display_path = _display_path(item.file_path, vendor_dirs)
            borderline_lines.append(
                f"| {no} | {item.date} | {item.doc_number} | {item.sender} | "
                f"{item.receiver} | {item.subject} | {display_path} |"
            )
    else:
        borderline_lines.append("_(없음)_")

    (output_dir / "scan_borderline.md").write_text(
        "\n".join(borderline_lines), encoding="utf-8"
    )

    # correspondence_texts.md — 전문(全文) 저장 (prepare 단계 입력)
    # confirmed + borderline 모두 날짜순으로 포함. borderline은 (참고) 표시.
    texts_lines = [
        "# 관련 공문 전문",
        "",
        "아래는 귀책분석에 활용할 공문 전문입니다.",
        "이 파일은 `python main.py prepare` 명령이 자동으로 사용합니다.",
        "",
        f"- 귀책 키워드 확정: {len(all_relevant)}개",
        f"- 참고 항목(자동 포함): {len(all_borderline)}개  ← **(참고)** 표시 항목",
        "",
        "---",
        "",
    ]
    all_for_texts = (
        [(item, "confirmed") for item in all_relevant] +
        [(item, "borderline") for item in all_borderline]
    )
    all_for_texts.sort(key=lambda x: re.sub(r"[^\d]", "", x[0].date) if x[0].date else "99999999")

    for no, (item, relevance) in enumerate(all_for_texts, 1):
        display_path = _display_path(item.file_path, vendor_dirs)
        date_note = ""
        if item.filename_date and item.ocr_date and item.filename_date != item.ocr_date:
            date_note = f" ⚠️ OCR 원본={item.ocr_date} (파일명 날짜로 교정됨 — 공문번호·발신처 직접 확인 요)"
        rel_label = "" if relevance == "confirmed" else " **(참고 — 귀책 키워드 미매칭)**"
        ocr_label = " ⚠️OCR품질주의" if item.ocr_quality == "WARN" else ""
        texts_lines += [
            f"## [{no}] {item.subject or item.file_path.name}{rel_label}{ocr_label}",
            "",
            f"- **발신일자**: {item.date}{date_note}",
            f"- **공문번호**: {item.doc_number}",
            f"- **발신처**: {item.sender}",
            f"- **수신처**: {item.receiver}",
            f"- **파일**: {display_path}",
            "",
            "### 전문",
            "",
            item.full_text,
            "",
            "---",
            "",
        ]
    (output_dir / "correspondence_texts.md").write_text(
        "\n".join(texts_lines), encoding="utf-8"
    )

    print(f"\n저장 완료:")
    print(f"  - {output_dir / 'scan_summary.md'}")
    print(f"  - {output_dir / 'scan_borderline.md'}  ← 수동 검토 {len(all_borderline)}개 / 자동 제외 {len(auto_excluded)}개")
    print(f"  - {output_dir / 'scan_result.json'}  ← 불필요 항목 삭제 후 prepare 실행 (편집 생략 가능)")
    print(f"  - {output_dir / 'scan_folder_tree.md'}")
    print(f"  - {output_dir / 'correspondence_texts.md'}")


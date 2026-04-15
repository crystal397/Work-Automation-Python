"""
파일 자동 분류기 — 수신자료를 한 폴더에 넣으면 카테고리별로 자동 분류
파일명 패턴 + 내용 키워드(첫 수백 바이트) 기준

분류 결과:
  INCLUDE  : 구조 추출에 포함
  SKIP     : 금액 데이터 (급여/영수증) — Step 2.7에서 수동 입력
  ※ 판단 불가 파일은 INCLUDE(LOW)로 처리 — 빠뜨리는 것보다 더 넣는 게 안전
"""

from __future__ import annotations

import re
from pathlib import Path


# ── SKIP 패턴 (명확한 금액 자료만 — 이것만 제외) ─────────────────────────────────

# 강제 SKIP — 어느 폴더에 있든 무조건 제외 (명백한 지출·경비 자료)
_FORCE_SKIP_PATTERNS = [
    # 파일명 끝에 금액이 붙은 영수증 (예: 봉달이설렁탕_24,000.pdf)
    # 쉼표가 포함된 숫자만 매칭 (날짜형 _20250315.pdf는 제외)
    r"_[\d,]*,[\d,]*\.(pdf|jpg|jpeg|png|tif|tiff)$",
    # 파일명 끝에 쉼표 없는 소액 금액 (3~5자리: 1,000~99,999원 범위)
    # 6자리 이상은 날짜(YYMMDD=250315)와 구분 필요 — 아래 패턴에서 별도 처리
    r"_\d{3,5}\.(pdf|jpg|jpeg|png|tif|tiff)$",
    # 6자리 숫자 중 YYMMDD 날짜가 아닌 것만 금액으로 판단
    # YYMMDD: 3~4번째 자리 = 월(01~12). 월이 00 또는 13~99이면 날짜가 아님 → 금액
    # 예) _136400 → 3~4번째=64 (월 불가) → 금액 / _250315 → 3~4번째=03 (월 가능) → 날짜
    r"_\d{2}(?:0{2}|1[3-9]|[2-9]\d)\d{2}\.(pdf|jpg|jpeg|png|tif|tiff)$",
    # 카드·전표·회계
    r"법인카드", r"카드사용내역", r"카드내역", r"카드사용",
    r"전표", r"계정별전표", r"총계정",
    r"계정별원장", r"계정별합계",
    r"재정집행", r"전도금",
    # 급여·노무
    r"급여대장", r"임금대장", r"급여명세", r"임금명세", r"급여명부",
    r"급여현황", r"임금현황", r"월급여", r"급여총괄",
    r"payroll", r"salary",
    r"급여",        # 개인별 급여 파일 포함
    r"노무비",
    r"일용직",
    # 영수증·세금계산서·거래명세
    r"세금계산서", r"영수증", r"invoice", r"receipt",
    r"거래명세서", r"거래명세표",
    r"보험료\s*영수증",
    # 복합기·사무기기 사용내역
    r"복합기",
    r"이용내역", r"사용\s*내역",
    # 건강검진·의료
    r"건강검진", r"건강진단", r"검진비", r"특수검진",
    # 출장·교통비
    r"출장비", r"출장신청", r"출장정산", r"출장여비",
    r"귀향여비", r"귀향비",
    r"교통비", r"시내교통비",
    r"여비",
    r"여비교통비\s*정산",
    # 피복·복지
    r"피복비", r"피복\s*지급", r"근무복",
    r"복리후생비",
    # 통신비
    r"통신비", r"통신요금",
    r"수신료",      # TV수신료 등
    # 당직·야간
    r"당직비", r"당직수당", r"야간수당",
    # 식대
    r"석식", r"중식", r"조식", r"식대",
    r"아워홈",      # 식당업체 청구서
    r"과세\s*비과세",
    # 차량·주유
    r"주유비", r"주유",
    r"차량수리", r"차량정비", r"차량유지", r"차량임차",
    r"수리비",
    # 소모성 경비
    r"탕비",        # 숙소탕비용품
    r"소모품",
    r"정수기",
    r"무인경비", r"월정료",
    r"구입대", r"구입비",
    r"퀵서비스", r"퀵\s*이용",
    # 공과금·광열비
    r"상수도", r"수도요금",
    r"전기요금", r"전기료", r"전력요금",
    r"도시가스",
    # 숙소
    r"숙소",
    # 세금·공과금
    r"주민세", r"재산세",
    r"세금과공과",
    # 유류비
    r"유류비", r"가스유류비",
    # 환경·점검비
    r"오수처리",
    r"수질측정", r"수질분석",
    r"안전점검",
    r"안전화",
    # 인쇄·도서비
    r"도서인쇄비", r"제본대",
    r"환경일보",    # 신문구독료
    # 수수료
    r"수수료", r"지급수수료", r"보증수수료",
    # 간접비 금액 자료
    r"간접비\s*산정자료", r"간접비\s*자료",
    r"명함",
]

# 일반 SKIP — 폴더 미분류 파일 대상 (INCLUDE 폴더 안 파일에는 적용 안 됨)
_SKIP_NAME_PATTERNS = [
    # 공정·일정 자료
    r"공정표", r"예정공정", r"공정계획", r"시공계획",
    r"3개월\s*공정", r"월간공정", r"주간공정",
    r"dpm\s*report", r"\d+weeks",
    # 회의 자료
    r"회의자료", r"회의록", r"월간회의", r"주간회의", r"공정회의",
    # 입찰·낙찰 관련 문서
    r"입찰공고", r"입찰참가", r"낙찰기준", r"심사기준",
    r"낙찰적격", r"종합심사", r"사전심사", r"관리지침", r"관리수칙",
    r"복수예비가격", r"이용약관", r"전자입찰",
    # 공사 관리·시방·도면
    r"시방서", r"공사시방", r"특기시방",
    r"설계도", r"도면",
    r"품질관리", r"안전관리", r"환경관리",
    r"구조계산", r"물량산출",
    # 행정 서류
    r"각서", r"서약서", r"청렴계약", r"담합",
    r"착공계", r"착공신고", r"준공계", r"준공신고",
    r"건설공사대장", r"공사대장", r"현장대장",
    r"현장사진", r"준공사진",
    r"시험성적", r"검사결과",
]

_SKIP_CONTENT_KEYWORDS = [
    "급여대장", "임금대장", "급여명세서", "임금명세서",
    "세금계산서", "공급가액", "세액합계", "부가세액",
]

# ── SKIP 폴더명 패턴 (상위 폴더명 기준) ──────────────────────────────────────────
# 파일 자체가 INCLUDE 패턴에 매칭되면 폴더 패턴은 무시됨

_SKIP_FOLDER_PATTERNS = [
    # 급여·노무 금액 자료
    r"노무비",
    r"급여명세", r"급여대장", r"임금대장", r"임금\s*증빙",
    # 경비 증빙 (폴더명이 경비 항목명인 경우)
    r"경비\s*증빙", r"경비\s*자료", r"^경비$",
    r"도서\s*인쇄", r"도서인쇄비",
    r"보증\s*수수료", r"보증수수료",
    r"복리\s*후생", r"복리후생비",
    r"세금과\s*공과", r"세금과공과",
    r"소모\s*용품", r"소모용품비", r"소모품",
    r"여비\s*교통", r"여비교통통신비", r"여비교통비",
    r"전력\s*수도", r"전력수도광열비", r"광열비",
    r"지급\s*임차", r"임차료", r"지급임차수수료",
    r"지급\s*수수료",
    r"출장비", r"교통비",
    # 영수증·전표
    r"영수증", r"전표", r"세금계산서",
    # 재정·회계
    r"재정", r"명함",
]

# ── INCLUDE 파일명 패턴 ────────────────────────────────────────────────────────────

_INCLUDE_NAME_PATTERNS = [
    # 계약 관련 — "계약" 포함 파일명 전반 (계약서, 변경계약, 계약현황, 하도급계약 등)
    r"계약",
    r"공동수급", r"협약서", r"약정서", r"낙찰통보", r"낙찰확인",
    r"일반조건", r"특수조건",
    # 공문·문서 (공기연장 귀책사유 확인용)
    r"공문", r"협조문", r"통보문", r"요청문", r"회신문", r"답변서",
    r"공지", r"통지", r"지시서", r"요청서",
    # 공기연장 관련
    r"공기연장", r"준공연장", r"기간연장", r"이행기간",
    r"실정보고", r"설계변경", r"변경설계",
    # 현장 조직·인원 (간접노무비 인원 확인용)
    r"조직도", r"현장조직", r"경력증명", r"재직증명", r"재직확인",
    r"인원현황", r"직원명단", r"인원배치", r"배치계획",
    r"신분증", r"자격증", r"면허증",
    # 산출내역·원가 (요율 확인용) — 다양한 축약형 포함
    r"산출내역", r"원가계산", r"공사원가", r"내역서", r"설계내역",
    r"기성내역", r"기성계",
    r"수량산출", r"단가산출", r"산출서",   # 수량산출서, 단가산출서
    r"단가표",                              # 단가표 (요율 포함 가능)
    # 하도급 (하도급사 목록 확인용)
    r"하도급", r"하청", r"외주", r"협력업체",
    # 보험·요율
    r"보험료", r"산재보험", r"고용보험", r"보험요율", r"요율확인",
    # 간접비 보고서·사감정
    r"간접비", r"추가공사비", r"사감정", r"감정보고서",
    # 인허가 (귀책사유 근거)
    r"인허가", r"민원",
    r"검토보고", r"검토의견",
]

_INCLUDE_CONTENT_KEYWORDS = [
    "도급계약서", "변경계약서", "공사도급", "하도급계약",
    "계약금액", "공사기간", "준공기한",
    "공기연장", "이행기간 변경", "설계변경",
    "현장소장", "공동수급협정", "산출내역서",
    "산재보험료", "고용보험요율", "일반관리비율",
]

# ── INCLUDE 폴더명 패턴 (폴더 이름이 일치하면 그 안의 파일 전체 포함) ──────────────
# 파일명 패턴과 무관하게, 상위 폴더명만으로 INCLUDE 판정

_INCLUDE_FOLDER_PATTERNS = [
    # 계약 관련 — "계약" 포함 폴더명 전반
    # (계약서, 계약현황, 변경계약, 하도급계약, 계약자료, 계약문서 등 모두 포함)
    r"계약",
    r"공동수급", r"협약", r"낙찰",
    # 공문·공기연장
    r"공문", r"협조문", r"통보문",
    r"공기연장", r"기간연장", r"준공연장", r"이행기간",
    r"설계변경", r"실정보고",
    # 현장 인원 (조직도·경력증명)
    r"조직도", r"현장조직", r"경력증명", r"재직증명",
    r"인원현황", r"직원명단", r"인원자료",
    # 산출내역·원가 — 다양한 명칭 포함
    r"산출내역", r"원가계산", r"공사원가", r"내역서",
    r"설계내역", r"기성내역",
    r"수량산출", r"단가산출", r"산출서",
    # 하도급
    r"하도급", r"하청", r"외주", r"협력업체",
    # 보험·요율
    r"보험료", r"산재보험", r"고용보험", r"보험요율",
    # 간접비·사감정
    r"간접비", r"추가공사비", r"사감정",
]


# ── 분류 로직 ─────────────────────────────────────────────────────────────────────

def _folder_is_include(path: Path) -> tuple[bool, str]:
    """상위 폴더명 중 INCLUDE 패턴에 매칭되는 것이 있으면 (True, 매칭 폴더명) 반환"""
    for part in path.parts[:-1]:
        if any(re.search(fp, part, re.IGNORECASE) for fp in _INCLUDE_FOLDER_PATTERNS):
            return True, part
    return False, ""


def _folder_is_skip(path: Path) -> bool:
    """상위 폴더명 중 SKIP 패턴에 매칭되는 것이 있으면 True"""
    return any(
        re.search(fp, part, re.IGNORECASE)
        for part in path.parts[:-1]
        for fp in _SKIP_FOLDER_PATTERNS
    )


def classify_file(path: Path) -> dict:
    """단일 파일 분류 → {action, label, confidence}

    우선순위:
      1. INCLUDE 폴더 매칭  → INCLUDE HIGH (사용자가 직접 관리하는 폴더 — 최우선)
      2. SKIP 폴더 매칭     → SKIP HIGH   (경비·급여 폴더 — 폴더 판단 신뢰)
      3. 강제 SKIP 파일명   → SKIP HIGH   (법인카드·금액포함 파일명)
      4. INCLUDE 파일명/내용 → INCLUDE
      5. SKIP 파일명/내용   → SKIP
      6. 어느 쪽도 아님     → INCLUDE LOW (빠뜨리는 것보다 더 넣는 게 안전)
    """
    name = path.name.lower()

    # 내용 첫 2KB (텍스트 파일만)
    content = ""
    try:
        if path.suffix.lower() in {".txt", ".md", ".csv"}:
            content = path.read_text(encoding="utf-8", errors="ignore")[:2000]
    except Exception:
        pass

    # 1) 강제 SKIP 파일명 — 폴더보다 우선 (파일명에 금액·법인카드 등 명백한 지출 신호)
    #    예: 250712_봉달이설렁탕_24,000.pdf → 어느 폴더에 있든 제외
    if any(re.search(p, name, re.IGNORECASE) for p in _FORCE_SKIP_PATTERNS):
        return {"action": "SKIP", "label": "금액자료(제외)", "confidence": "HIGH"}

    # 2) INCLUDE 폴더 — 사용자가 직접 관리하는 폴더명 (강제 SKIP 통과한 파일만 적용)
    folder_inc, matched_folder = _folder_is_include(path)
    if folder_inc:
        label = _guess_label(name) if _guess_label(name) != "기타" else _guess_label(matched_folder)
        return {"action": "INCLUDE", "label": label, "confidence": "HIGH",
                "folder_match": matched_folder}

    # 3) SKIP 폴더 — 경비·급여 폴더 하위 파일은 무조건 제외
    if _folder_is_skip(path):
        return {"action": "SKIP", "label": "금액자료(제외)", "confidence": "HIGH"}

    # 4) INCLUDE 파일명/내용 매칭
    inc_hits = sum(1 for p in _INCLUDE_NAME_PATTERNS if re.search(p, name, re.IGNORECASE))
    inc_hits += sum(1 for k in _INCLUDE_CONTENT_KEYWORDS if k in content)
    if inc_hits >= 1:
        confidence = "HIGH" if inc_hits >= 3 else "MED" if inc_hits >= 2 else "LOW"
        label = _guess_label(name)
        return {"action": "INCLUDE", "label": label, "confidence": confidence}

    # 5) SKIP 파일명/내용 매칭
    skip_hits = sum(1 for p in _SKIP_NAME_PATTERNS if re.search(p, name, re.IGNORECASE))
    skip_hits += sum(1 for k in _SKIP_CONTENT_KEYWORDS if k in content)
    if skip_hits >= 1:
        confidence = "HIGH" if skip_hits >= 2 else "MED"
        return {"action": "SKIP", "label": "금액자료(제외)", "confidence": confidence}

    # 6) 패턴 미매칭 → INCLUDE LOW (빠뜨리는 것보다 더 포함하는 게 안전)
    return {"action": "INCLUDE", "label": "기타(포함)", "confidence": "LOW"}


def _guess_label(name: str) -> str:
    """파일명으로 세부 레이블 추정"""
    if re.search(r"계약서|도급|하도급|협약|약정", name, re.I):
        return "계약서"
    if re.search(r"공문|협조문|통보문|공기연장|설계변경|실정보고", name, re.I):
        return "공문"
    if re.search(r"조직도|경력증명|재직증명|인원", name, re.I):
        return "조직도/경력"
    if re.search(r"산출내역|원가계산|내역서|기성내역|단가", name, re.I):
        return "산출내역서"
    if re.search(r"간접비|추가공사비|사감정|정산", name, re.I):
        return "간접비자료"
    if re.search(r"보험|요율|산재|고용", name, re.I):
        return "보험요율"
    if re.search(r"보증서|이행보증|계약보증", name, re.I):
        return "보증서"
    return "기타"


def classify_folder(folder: str | Path) -> list[dict]:
    """폴더 내 모든 파일 분류 → 결과 목록"""
    folder = Path(folder)
    if not folder.exists():
        return []

    supported = {".pdf", ".xlsx", ".xls", ".docx", ".hwp",
                 ".txt", ".csv", ".jpg", ".png", ".jpeg"}
    results = []

    for path in sorted(folder.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix.lower() not in supported:
            continue

        result = classify_file(path)
        results.append({
            "path":     str(path.relative_to(folder)),
            "filename": path.name,
            **result,
        })

    return results


def print_classify_report(results: list[dict], folder: str) -> str:
    """분류 결과 출력"""
    include = [r for r in results if r["action"] == "INCLUDE"]
    skip    = [r for r in results if r["action"] == "SKIP"]

    # 신뢰도별 분리
    inc_high = [r for r in include if r["confidence"] == "HIGH"]
    inc_med  = [r for r in include if r["confidence"] == "MED"]
    inc_low  = [r for r in include if r["confidence"] == "LOW"]

    lines = [
        "",
        "=" * 62,
        "  파일 자동 분류 결과",
        "=" * 62,
        f"  대상 폴더: {folder}",
        f"  전체 파일: {len(results)}개",
        f"  INCLUDE  : {len(include)}개  (확실 {len(inc_high)} / 추정 {len(inc_med)} / 기타 {len(inc_low)})",
        f"  SKIP     : {len(skip)}개  → 금액 수동 입력 예정",
        "=" * 62,
    ]

    if inc_high or inc_med:
        lines.append("\n[INCLUDE — 구조 자료 (확실/추정)]")
        for r in inc_high + inc_med:
            folder_note = f"  ← 폴더: {r['folder_match']}" if r.get("folder_match") else ""
            lines.append(f"  ★ [{r['label']:14s}]  {r['filename']}{folder_note}")

    if inc_low:
        lines.append("\n[INCLUDE — 기타 (패턴 미매칭, 안전하게 포함)]")
        for r in inc_low:
            lines.append(f"  ☆ [{r['label']:14s}]  {r['filename']}")

    if skip:
        lines.append("\n[SKIP — 금액 자료 (Step 2.7에서 수동 입력)]")
        for r in skip:
            conf = "★" if r["confidence"] == "HIGH" else "☆"
            lines.append(f"  {conf} [{r['label']:14s}]  {r['filename']}")

    lines += [
        "",
        "[다음 단계]  INCLUDE 파일만 자동 복사 후 추출:",
        "    python main.py classify --copy",
        "    python main.py extract --filtered",
        "",
        "  SKIP 파일 금액은 Step 2.7에서 입력:",
        "    python main.py amounts",
        "=" * 62,
    ]

    report = "\n".join(lines)
    print(report)
    return report


def copy_includes(results: list[dict],
                  src_folder: str | Path,
                  dst_folder: str | Path,
                  dry_run: bool = False) -> list[str]:
    """INCLUDE 파일을 dst_folder에 복사 (기존 폴더 내용 먼저 삭제)"""
    import shutil
    src_folder = Path(src_folder)
    dst_folder = Path(dst_folder)

    if not dry_run:
        # 기존 폴더를 비우고 새로 생성 (이전 분류 결과 잔존 방지)
        if dst_folder.exists():
            shutil.rmtree(dst_folder)
        dst_folder.mkdir(parents=True, exist_ok=True)

    copied = []
    for r in results:
        if r["action"] != "INCLUDE":
            continue
        src = src_folder / r["path"]
        dst = dst_folder / r["filename"]
        if not dry_run:
            shutil.copy2(src, dst)
        copied.append(r["filename"])

    return copied

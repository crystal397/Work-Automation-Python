"""
귀책분석_data.json 재구성 스크립트
- 감리단 포워딩 공문 4건을 각각 별도 행으로 분리 (4개 → 8행)
- show_in_table 플래그 추가 (공문 목록 표 대상 여부)
- sender / receiver 약칭 필드 추가 (공문 목록 표 표시용)
- background_paragraphs / detail_narratives / accountability_diagram /
  conclusion_paragraphs 필드 추가 (report_generator.py 가 읽어 서술 생성)
"""

import json
from pathlib import Path

BASE = Path(__file__).parent
SRC  = BASE / "output" / "귀책분석_data.json"
DST  = BASE / "output" / "귀책분석_data.json"   # 덮어쓰기

with open(SRC, encoding="utf-8") as f:
    old = json.load(f)

old_by_no = {it["no"]: it for it in old["items"]}
def o(no):
    return old_by_no[no]


# ── 새 items 정의 ─────────────────────────────────────────────────────────────
# show_in_table=True 인 16개가 공문 목록 표에 출력된다.
# sender / receiver  : 공문 목록 표 표시용 약칭
# sender_full / receiver_full : 서술 단락용 전체 명칭

items = [
    # ─ 1. 도로관리심의 신청 1/4분기
    {
        "no": 1,
        "date": "2024.01.19.",
        "doc_number": "DL인동선5 제2024-0008호",
        "subject": "도로관리심의회 심의 신청 요청(인덕원-동탄 제5공구) — 1/4분기",
        "sender": "계약상대자",
        "sender_full": o(1)["sender"],
        "receiver": "발주자",
        "receiver_full": o(1)["receiver"],
        "show_in_table": True,
        "event_type": o(1)["event_type"],
        "causal_party": o(1)["causal_party"],
        "causal_description": o(1)["causal_description"],
        "delay_days": o(1)["delay_days"],
        "note": o(1)["note"],
    },
    # ─ 2. 한전주 이설 요청
    {
        "no": 2,
        "date": "2024.01.29.",
        "doc_number": "DL인동선5 제2024-0010호",
        "subject": "한전주 이설 요청(인덕원-동탄 5공구)",
        "sender": "계약상대자",
        "sender_full": o(2)["sender"],
        "receiver": "발주자",
        "receiver_full": o(2)["receiver"],
        "show_in_table": True,
        "event_type": o(2)["event_type"],
        "causal_party": o(2)["causal_party"],
        "causal_description": o(2)["causal_description"],
        "delay_days": o(2)["delay_days"],
        "note": o(2)["note"],
    },
    # ─ 3. 노송지대 문화재 현상변경 허가
    {
        "no": 3,
        "date": "2024.03.15.",
        "doc_number": "문화예술과-5780",
        "subject": "경기도 지정문화재(노송지대) 현상변경 허가 알림(인덕원-동탄 복선전철 5공구 공사)",
        "sender": "수원시",
        "sender_full": o(3)["sender"],
        "receiver": "계약상대자",
        "receiver_full": o(3)["receiver"],
        "show_in_table": True,
        "event_type": o(3)["event_type"],
        "causal_party": o(3)["causal_party"],
        "causal_description": o(3)["causal_description"],
        "delay_days": o(3)["delay_days"],
        "note": o(3)["note"],
    },
    # ─ 4. 1/4분기 심의결과 알림 (발주자→건설사업관리단) — 구 no.4 분리
    {
        "no": 4,
        "date": "2024.03.19.",
        "doc_number": "수도권사업단-4703",
        "subject": "1분기 도로관리심의 결과 알림",
        "sender": "발주자",
        "sender_full": "국가철도공단 수도권본부장",
        "receiver": "건설사업관리단",
        "receiver_full": "인덕원~동탄 복선전철 제5,6공구 감리단",
        "show_in_table": True,
        "event_type": "인허가",
        "causal_party": "발주처",
        "causal_description": o(4)["causal_description"],
        "delay_days": 0,
        "note": o(4)["note"],
    },
    # ─ 5. 1/4분기 심의결과 알림 (건설사업관리단→계약상대자) — 구 no.4 포워딩
    {
        "no": 5,
        "date": "2024.03.19.",
        "doc_number": "인동선56감 제2024-0062호",
        "subject": "1분기 도로관리심의 결과 알림(인덕원~동탄 5,6공구)",
        "sender": "건설사업관리단",
        "sender_full": "인덕원~동탄 복선전철 제5,6공구 감리단장",
        "receiver": "계약상대자",
        "receiver_full": "인덕원~동탄 복선전철 제5공구 현장대리인",
        "show_in_table": True,
        "event_type": "인허가",
        "causal_party": "발주처",
        "causal_description": "인동선56감 제2024-0062호(2024.03.19.)에 의거, 건설사업관리단이 수도권사업단-4703호(2024.03.19.)에 따른 2024년 1/4분기 도로관리심의회 심의 결과를 계약상대자에게 공식 전달하였다.",
        "delay_days": 0,
        "note": "수도권사업단-4703호 전달",
    },
    # ─ 6. 도로관리심의 신청 2/4분기 — 구 no.5
    {
        "no": 6,
        "date": "2024.03.19.",
        "doc_number": "DL인동선5 제2024-0028호",
        "subject": "도로관리심의회 심의 신청 요청(인덕원-동탄 제5공구) — 2/4분기",
        "sender": "계약상대자",
        "sender_full": o(5)["sender"],
        "receiver": "건설사업관리단",
        "receiver_full": o(5)["receiver"],
        "show_in_table": True,
        "event_type": o(5)["event_type"],
        "causal_party": o(5)["causal_party"],
        "causal_description": o(5)["causal_description"],
        "delay_days": o(5)["delay_days"],
        "note": o(5)["note"],
    },
    # ─ 7. 수원광주이씨고택 문화재 허가 (발주자→건설사업관리단) — 구 no.6 분리
    {
        "no": 7,
        "date": "2024.04.18.",
        "doc_number": "신안산선사업단-192",
        "subject": "문화재 현상변경 허가 알림(인덕원~동탄 5공구, 수원광주이씨고택)",
        "sender": "발주자",
        "sender_full": "국가철도공단 수도권본부장",
        "receiver": "건설사업관리단",
        "receiver_full": "인덕원~동탄 복선전철 제5,6공구 감리단",
        "show_in_table": True,
        "event_type": "인허가",
        "causal_party": "발주처",
        "causal_description": o(6)["causal_description"],
        "delay_days": 0,
        "note": o(6)["note"],
    },
    # ─ 8. 수원광주이씨고택 문화재 허가 (건설사업관리단→계약상대자) — 구 no.6 포워딩
    {
        "no": 8,
        "date": "2024.04.19.",
        "doc_number": "인동선56감 제2024-0167호",
        "subject": "문화재 현상변경 허가 알림(인덕원~동탄 5공구, 수원광주이씨고택)",
        "sender": "건설사업관리단",
        "sender_full": "인덕원~동탄 복선전철 제5,6공구 감리단장",
        "receiver": "계약상대자",
        "receiver_full": "인덕원~동탄 복선전철 제5공구 현장대리인",
        "show_in_table": True,
        "event_type": "인허가",
        "causal_party": "발주처",
        "causal_description": (
            "인동선56감 제2024-0167호(2024.04.19.)에 의거, 건설사업관리단이 "
            "신안산선사업단-192호(2024.04.18.)에 따른 수원 광주이씨고택 문화재 "
            "현상변경 허가 결과를 계약상대자에게 공식 전달하였다. 이로써 계약상대자는 "
            "해당 구간 공사착수 전 문화재 허가조건 이행 의무를 부과받았다."
        ),
        "delay_days": 0,
        "note": "신안산선사업단-192 전달 / 허가번호 제2024-0551호",
    },
    # ─ 9. 도로점용허가 신청 — 구 no.7
    {
        "no": 9,
        "date": "2024.04.26.",
        "doc_number": "DL인동선5 제2024-0053호",
        "subject": "106정거장 및 환기구7번 도로점용허가 신청 요청(인덕원-동탄 제5공구)",
        "sender": "계약상대자",
        "sender_full": o(7)["sender"],
        "receiver": "건설사업관리단",
        "receiver_full": o(7)["receiver"],
        "show_in_table": True,
        "event_type": o(7)["event_type"],
        "causal_party": o(7)["causal_party"],
        "causal_description": o(7)["causal_description"],
        "delay_days": o(7)["delay_days"],
        "note": o(7)["note"],
    },
    # ─ 10. 도로점용허가 완료 (발주자→건설사업관리단) — 구 no.8 분리
    {
        "no": 10,
        "date": "2024.05.30.",
        "doc_number": "신안산선사업단-3743",
        "subject": "도로점용허가 완료알림(인덕원~동탄 5공구, 환기구 7번/106정거장)",
        "sender": "발주자",
        "sender_full": "국가철도공단 수도권본부장",
        "receiver": "건설사업관리단",
        "receiver_full": "인덕원~동탄 복선전철 제5,6공구 감리단",
        "show_in_table": True,
        "event_type": "인허가",
        "causal_party": "발주처",
        "causal_description": o(8)["causal_description"],
        "delay_days": 93,
        "note": o(8)["note"],
    },
    # ─ 11. 도로점용허가 완료 (건설사업관리단→계약상대자) — 구 no.8 포워딩
    {
        "no": 11,
        "date": "2024.05.31.",
        "doc_number": "인동선56감 제2024-0296호",
        "subject": "도로점용허가 완료알림(인덕원~동탄 5공구, 환기구 7번/106정거장)",
        "sender": "건설사업관리단",
        "sender_full": "인덕원~동탄 복선전철 제5,6공구 감리단장",
        "receiver": "계약상대자",
        "receiver_full": "인덕원~동탄 복선전철 제5공구 현장대리인",
        "show_in_table": True,
        "event_type": "인허가",
        "causal_party": "발주처",
        "causal_description": (
            "인동선56감 제2024-0296호(2024.05.31.)에 의거, 건설사업관리단이 "
            "신안산선사업단-3743호(2024.05.30.)에 따른 환기구 7번 및 106정거장 "
            "도로점용허가 완료 사실을 계약상대자에게 공식 전달하였다."
        ),
        "delay_days": 0,
        "note": "신안산선사업단-3743 전달 / 허가번호 제2024-48호",
    },
    # ─ 12. 한전 시설부담금 청구 — 구 no.9
    {
        "no": 12,
        "date": "2024.06.27.",
        "doc_number": "경기본(고객)-2304",
        "subject": "배전선로 이설공사 시설부담금 청구 [인덕원-동탄 제5공구 106정거장 공사]",
        "sender": "한국전력공사",
        "sender_full": o(9)["sender"],
        "receiver": "발주자",
        "receiver_full": o(9)["receiver"],
        "show_in_table": True,
        "event_type": o(9)["event_type"],
        "causal_party": o(9)["causal_party"],
        "causal_description": o(9)["causal_description"],
        "delay_days": o(9)["delay_days"],
        "note": o(9)["note"],
    },
    # ─ 13. 이설공사비 지급 완료 (발주자→건설사업관리단) — 구 no.11 분리
    {
        "no": 13,
        "date": "2024.10.04.",
        "doc_number": "인덕원동탄사업단TF-2851",
        "subject": "지장물 이설 공사비 지급 완료 알림(인덕원~동탄 5공구, 한국전력공사)",
        "sender": "발주자",
        "sender_full": "국가철도공단 수도권본부장",
        "receiver": "건설사업관리단",
        "receiver_full": "인덕원~동탄 복선전철 제5,6공구 감리단",
        "show_in_table": True,
        "event_type": "지연요인",
        "causal_party": "발주처",
        "causal_description": o(11)["causal_description"],
        "delay_days": 0,
        "note": o(11)["note"],
    },
    # ─ 14. 이설공사비 지급 완료 (건설사업관리단→계약상대자) — 구 no.11 포워딩
    {
        "no": 14,
        "date": "2024.10.04.",
        "doc_number": "인동선56감 제2024-0904호",
        "subject": "지장물 이설 공사비 지급 완료 알림(인덕원~동탄 5공구, 한국전력공사)",
        "sender": "건설사업관리단",
        "sender_full": "인덕원~동탄 복선전철 제5,6공구 감리단장",
        "receiver": "계약상대자",
        "receiver_full": "인덕원~동탄 복선전철 제5공구 현장대리인",
        "show_in_table": True,
        "event_type": "지연요인",
        "causal_party": "발주처",
        "causal_description": (
            "인동선56감 제2024-0904호(2024.10.04.)에 의거, 건설사업관리단이 "
            "인덕원동탄사업단TF-2851호(2024.10.04.)에 따른 한국전력공사 지장물 "
            "이설공사비 지급 완료 사실을 계약상대자에게 공식 전달하였다."
        ),
        "delay_days": 0,
        "note": "인덕원동탄사업단TF-2851 전달",
    },
    # ─ 15. 제1차 1회 준공기한 연장 계약변경 요청 — 구 no.14
    {
        "no": 15,
        "date": "2024.11.27.",
        "doc_number": "DL인동선5 제2024-0351호",
        "subject": "[인덕원~동탄 5공구] 제1차 1회 준공기한 연장 계약변경 요청",
        "sender": "계약상대자",
        "sender_full": o(14)["sender"],
        "receiver": "건설사업관리단",
        "receiver_full": o(14)["receiver"],
        "show_in_table": True,
        "event_type": "지연요인",
        "causal_party": "발주처",
        "causal_description": o(14)["causal_description"],
        "delay_days": 366,
        "note": o(14)["note"],
    },
    # ─ 16. 제1차 3회 계약변경 요청 — 구 no.20
    {
        "no": 16,
        "date": "2025.11.26.",
        "doc_number": "DL인동선5 제2025-0805호",
        "subject": "[인덕원~동탄 5공구] 설계변경(제1차 3회)에 따른 계약변경 요청",
        "sender": "계약상대자",
        "sender_full": o(20)["sender"],
        "receiver": "건설사업관리단",
        "receiver_full": o(20)["receiver"],
        "show_in_table": True,
        "event_type": "설계변경",
        "causal_party": "발주처",
        "causal_description": o(20)["causal_description"],
        "delay_days": 121,
        "note": o(20)["note"],
    },
    # ── 비표 항목 (show_in_table: false) ─────────────────────────────────────
    {
        "no": 17, "date": "2024.09.02.", "doc_number": "",
        "subject": o(10)["subject"],
        "sender": "계약상대자", "sender_full": o(10)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(10)["receiver"],
        "show_in_table": False,
        "event_type": o(10)["event_type"], "causal_party": o(10)["causal_party"],
        "causal_description": o(10)["causal_description"],
        "delay_days": 0, "note": o(10)["note"],
    },
    {
        "no": 18, "date": "2024.11.25.", "doc_number": "DL인동선5 제2024-0345호",
        "subject": o(12)["subject"],
        "sender": "계약상대자", "sender_full": o(12)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(12)["receiver"],
        "show_in_table": False,
        "event_type": o(12)["event_type"], "causal_party": o(12)["causal_party"],
        "causal_description": o(12)["causal_description"],
        "delay_days": o(12)["delay_days"], "note": o(12)["note"],
    },
    {
        "no": 19, "date": "2024.11.25.", "doc_number": "DL인동선5 제2024-0346호",
        "subject": o(13)["subject"],
        "sender": "계약상대자", "sender_full": o(13)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(13)["receiver"],
        "show_in_table": False,
        "event_type": o(13)["event_type"], "causal_party": o(13)["causal_party"],
        "causal_description": o(13)["causal_description"],
        "delay_days": o(13)["delay_days"], "note": o(13)["note"],
    },
    {
        "no": 20, "date": "2025.09.15.", "doc_number": "",
        "subject": o(15)["subject"],
        "sender": "계약상대자", "sender_full": o(15)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(15)["receiver"],
        "show_in_table": False,
        "event_type": o(15)["event_type"], "causal_party": o(15)["causal_party"],
        "causal_description": o(15)["causal_description"],
        "delay_days": 0, "note": o(15)["note"],
    },
    {
        "no": 21, "date": "2025.10.13.", "doc_number": "",
        "subject": o(16)["subject"],
        "sender": "계약상대자", "sender_full": o(16)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(16)["receiver"],
        "show_in_table": False,
        "event_type": o(16)["event_type"], "causal_party": o(16)["causal_party"],
        "causal_description": o(16)["causal_description"],
        "delay_days": 0, "note": o(16)["note"],
    },
    {
        "no": 22, "date": "2025.11.17.", "doc_number": "",
        "subject": o(17)["subject"],
        "sender": "계약상대자", "sender_full": o(17)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(17)["receiver"],
        "show_in_table": False,
        "event_type": o(17)["event_type"], "causal_party": o(17)["causal_party"],
        "causal_description": o(17)["causal_description"],
        "delay_days": 0, "note": o(17)["note"],
    },
    {
        "no": 23, "date": "2025.11.24.", "doc_number": "",
        "subject": o(18)["subject"],
        "sender": "계약상대자", "sender_full": o(18)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(18)["receiver"],
        "show_in_table": False,
        "event_type": o(18)["event_type"], "causal_party": o(18)["causal_party"],
        "causal_description": o(18)["causal_description"],
        "delay_days": 0, "note": o(18)["note"],
    },
    {
        "no": 24, "date": "2025.11.03.", "doc_number": "DL인동선5 제2025-0736호",
        "subject": o(19)["subject"],
        "sender": "계약상대자", "sender_full": o(19)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(19)["receiver"],
        "show_in_table": False,
        "event_type": o(19)["event_type"], "causal_party": o(19)["causal_party"],
        "causal_description": o(19)["causal_description"],
        "delay_days": o(19)["delay_days"], "note": o(19)["note"],
    },
    {
        "no": 25, "date": "2024.11.27.",
        "doc_number": "DL인동선5 제2024-0351호 첨부 공기연장 검토서",
        "subject": o(21)["subject"],
        "sender": "계약상대자", "sender_full": o(21)["sender"],
        "receiver": "건설사업관리단", "receiver_full": o(21)["receiver"],
        "show_in_table": False,
        "event_type": o(21)["event_type"], "causal_party": o(21)["causal_party"],
        "causal_description": o(21)["causal_description"],
        "delay_days": o(21)["delay_days"], "note": o(21)["note"],
    },
]


# ── 새 JSON 구성 ──────────────────────────────────────────────────────────────
new_data = {
    "project_name": old["project_name"],

    # ── 선택적 섹션 제목 (없으면 report_generator.py 가 기본값 사용) ──────────
    # "chapter_heading": "제2장  귀책분석",
    # "section_3_1_heading": "3.1. 공기연장 귀책 분석",
    # "section_3_1_1_heading": "3.1.1. 지연사유 발생 경위",
    # "section_3_1_2_heading": "3.1.2. 공기지연 귀책 사유 검토",
    # "table_intro_sentence": "이와 관련해 ...",

    # ── 3.1.1 도입부 배경 서술 단락 ─────────────────────────────────────────
    "background_paragraphs": [
        (
            "계약상대자는 본건 공사의 1차수 공사는 당초 준공일인 2024. 11. 29. 이전인 "
            "2024. 11. 27.에 건설사업관리단에게 준공기한 연장 계약변경을 요청하여 "
            "2024. 12. 2.에 1회차 공사 변경계약을 체결하였고, 이후 변경 준공일인 "
            "2025. 11. 30. 이전인 2025. 11. 26.에 건설사업관리단에게 준공기한 연장 "
            "계약변경을 요청하여 2025. 12. 1.에 3회차 공사 변경 계약을 체결하였습니다."
        ),
        (
            "1회차 공사 변경 계약의 공사기간 연장 사유는 '인허가 및 지장물 이설 "
            "소요기간 필요에 따른 지연'이며, 3회차 공사 변경 계약의 공사기간 연장 "
            "사유는 '지장물 이설 지연 및 현장여건 변경 등에 따른 공사추진계획 변경'으로 "
            "앞서 살펴본 관계 법령 및 조건에 따라 계약상대자의 책임 없는 사유에 해당합니다."
        ),
    ],

    # ── items (25개) ────────────────────────────────────────────────────────
    "items": items,

    # ── 공문 목록 표 이후 상세 서술 단락 (순서대로 출력) ─────────────────────
    "detail_narratives": [
        {
            "paragraphs": [
                (
                    "계약상대자는 2024. 11. 27. 'DL인동선5 제2024-0351호' 공문을 통해 "
                    "본건 공사 중 제1차 공사의 준공기한 연장을 위한 제1회 계약변경 요청 "
                    "서류를 제출하며 계약기간 연장을 신청하였습니다. 동 공문 제출에 앞서 "
                    "계약상대자는 2024. 11. 25. DL인동선5 제2024-0345호 및 제2024-0346호를 "
                    "통해 준공기한 연장을 사전 요청한 바 있습니다."
                ),
                (
                    "당시 첨부된 '제1차 공기연장 검토서'에 따르면, 공기 지연 사유로 "
                    "각종 인허가 절차 지연이 확인됩니다. 당초 인허가 소요기간을 3개월로 "
                    "계획하였으나, 실제로는 2023. 12. 22. 공사계약 체결 이후 관계기관 협의 "
                    "및 심의 절차를 거쳐 2024. 5. 23. 도로점용허가가 완료되는 등 약 6개월이 "
                    "소요되었습니다. 이 과정에는 도로굴착허가를 위한 도로관리심의회 2회 심의 "
                    "(1/4분기: 2024.03.12. 완료, 2/4분기: 2024.04.16. 완료), 경기도 지정문화재 "
                    "(노송지대) 현상변경 허가(2024.03.15.), 국가지정문화재(수원 광주이씨고택) "
                    "현상변경 허가(2024.04.12.) 등이 포함되었습니다."
                ),
                (
                    "또한 지장물 이설과 관련하여, 한전주, 광역상수도 및 상수도 관로, "
                    "통신관로(KT, SKT, LGU+, SKB, 드림라인 등) 등 다수의 지장물에 대해 "
                    "각 관리주체와의 협의, 설계, 분담금 납부 및 협약 체결 등의 절차가 "
                    "순차적으로 진행되었으며, 이에 따른 전체 지장물 이설 완료 시점은 "
                    "2025년 3월 말로 예상되어 약 15개월의 기간이 소요되는 것으로 확인되었습니다."
                ),
                (
                    "이와 같은 인허가 지연 및 지장물 이설 소요기간을 반영하여 계약상대자는 "
                    "본 공사의 원활한 추진을 위해 준공기한을 2025. 11. 30.까지 연장할 "
                    "필요가 있음을 통지하였으며, 지장물 이설 일정에 따라 향후 공사 일정이 "
                    "추가로 변동될 수 있음을 함께 알렸습니다."
                ),
            ]
        },
        {
            "paragraphs": [
                (
                    "이에 따라 계약당사자들은 2024. 12. 2.에 변경계약을 체결하여 "
                    "제1차 공사의 준공일을 당초 2024. 11. 29.에서 2025. 11. 30.으로 변경하였습니다."
                ),
            ]
        },
        {
            "paragraphs": [
                (
                    "이후 계약상대자는 2025. 11. 26. 'DL인동선5 제2025-0805호' 공문을 통하여 "
                    "본건 공사 중 제1차 공사의 준공기한 연장을 위한 제3회 계약변경 요청 서류를 "
                    "제출하며 계약기간 연장을 신청하였습니다."
                ),
                (
                    "첨부된 '제1차 공사 공기연장 검토서(2회차)'에 따르면, 추가적인 공기 "
                    "지연 사유로는 지장물 이설 지연, 지반조건 변경에 따른 추가 시공, "
                    "인허가 조건에 따른 작업 제한, 통신 지장물과의 작업 간섭 등이 확인됩니다."
                ),
                (
                    "우선 ①광역상수도 및 한전주 이설이 당초 계획 대비 지연되면서 수직구 "
                    "H파일 항타 작업이 약 2개월의 착수 지연이 발생하였고, ②수직구 구간의 "
                    "암선 변경, 연결터널 상부 지반 불량, 연결터널부터 본선터널까지 접속부 "
                    "보완설계 등 지반조건 변화에 따른 추가 시공으로 약 1.7개월의 지연이 "
                    "발생한 것으로 검토되었으며, ③화약류 양수허가 관련 인허가 조건에 따른 "
                    "작업 제한, ④통신 지장물과의 작업 간섭으로 인한 공정 영향이 확인되었습니다."
                ),
                (
                    "또한 106정거장(북수원역) 구간은 당초 계획 대비 지장물 이설 지연이 "
                    "지속되어 단계별 복공 완료 시점이 지연됨에 따라 후속 공정에도 영향이 "
                    "발생할 수 있는 것으로 분석되었으며, 이에 따라 정거장 구간 공정 또한 "
                    "2026년 3월에 복공 6단계가 완료될 예정입니다."
                ),
                (
                    "이와 같은 사유들을 종합한 결과, 기존 1회 변경계약에서 반영된 일정 대비 "
                    "추가적인 공정 지연이 불가피한 것으로 검토되었고, 계약상대자는 본 공사의 "
                    "원활한 추진을 위해 준공기한을 2026. 3. 31.까지 연장할 필요가 있음을 "
                    "통지하였으며, 지장물 이설 일정 및 현장 여건 변화에 따라 공사 일정이 "
                    "추가로 변동될 수 있음을 함께 알렸습니다."
                ),
            ]
        },
        {
            "paragraphs": [
                (
                    "이에 따라 계약상대자들은 2025. 12. 1.에 변경계약을 체결하여 "
                    "제1차 공사의 준공일을 당초 2025. 11. 30.에서 2026. 3. 31.로 변경하였습니다."
                ),
            ]
        },
    ],

    # ── 3.1.2 도식표 전 서술 단락 ───────────────────────────────────────────
    "pre_diagram_paragraphs": [
        (
            "앞서 서술된 바와 같이, 본건 공사 중 1차수 공사는 '인허가 및 지장물 이설 "
            "소요기간 필요에 따른 지연', '지장물 이설 지연 및 현장여건 변경 등에 따른 "
            "공사추진계획 변경'으로 인하여 최초 계약체결 당시 준공일 2024. 11. 29.에서 "
            "1회 변경 계약시 준공일이 2025. 11. 30.으로, 3회 변경 계약시 준공일이 "
            "2026. 3. 31.로 총 487일 연장되었습니다."
        ),
        (
            "계약상대자는 공사계약 일반조건 제26조 제1항에 따라 공기지연 사유가 발생하는 "
            "경우 계약기간 종료 전에 지체없이 계약기간의 연장신청을 하여야 하며, "
            "공기지연 사유가 공사계약 일반조건 제25조 제3항에 따라 계약상대자의 책임 있는 "
            "사유가 아닌 경우 지체상금을 부과하여서는 아니됩니다."
        ),
        (
            "본건 공사의 공기지연은 공사계약 일반조건 제11조(공사용지)와 "
            "제25조 제3항 제3호 '발주기관의 책임으로 착공이 지연되거나 시공이 중단되었을 "
            "경우'에 해당하는 계약상대자의 책임 없는 사유입니다. 이를 도식화하여 "
            "정리하면 아래와 같습니다."
        ),
    ],

    # ── 귀책사유 도식표 (공기지연 사유 | 관련 근거 | 비용부담자) ──────────────
    "accountability_diagram": [
        {
            "cause": (
                "① 발주기관은 계약문서에 따로 정한 경우를 제외하고는 계약상대자가 "
                "공사의 수행에 필요로 하는 날까지 공사용지를 확보하여 계약상대자에게 "
                "인도하여야 한다.\n"
                "③ 발주기관은 공사용지 확보 및 민원대응 등 공사용지 확보와 직접 관련되는 "
                "업무를 계약상대자에게 전가하여서는 아니된다."
            ),
            "basis": "공사계약 일반조건\n제11조(공사용지)\n①, ③항",
            "responsible_party": "발주자",
        },
        {
            "cause": (
                "③ 계약담당공무원은 다음 각호의 어느 하나에 해당되어 공사가 지체되었다고 "
                "인정할 때에는 그 해당일수를 제1항의 지체일수에 산입하지 아니한다.\n"
                "1. 제32조에서 정한 불가항력의 사유에 의한 경우\n"
                "2. 계약상대자가 대체 사용할 수 없는 중요 사급자재 등의 공급이 지연되어 "
                "공사 진행이 불가능하였을 경우\n"
                "3. 발주기관의 책임으로 착공이 지연되거나 시공이 중단되었을 경우\n"
                "5. 제19조에 의한 설계변경(계약상대자의 책임없는 사유인 경우에 한한다)으로 "
                "인하여 준공기한내에 계약을 이행할 수 없을 경우\n"
                "7. 기타 계약상대자의 책임에 속하지 아니하는 사유로 인하여 지체된 경우"
            ),
            "basis": "공사계약 일반조건\n제25조(지체상금)\n제3항 각호",
            "responsible_party": "발주자",
        },
    ],

    # ── 3.1.2 결론 서술 단락 ────────────────────────────────────────────────
    "conclusion_paragraphs": [
        (
            "앞서 서술된 본건 공사지연 사유인 '인허가 및 지장물 이설 소요기간 필요에 따른 "
            "지연', '지장물 이설 지연 및 현장여건 변경 등에 따른 공사추진계획 변경'은 "
            "공사계약 일반조건 제11조 제1항 및 제3항과 공사계약 일반조건 제25조 제3항의 "
            "'발주기관의 책임으로 착공이 지연되거나 시공이 중단되었을 경우'에 해당하며, "
            "이는 발주자의 귀책 사유에 해당합니다."
        ),
        (
            "따라서 공사계약 일반조건 제26조 제1항에 의거 계약상대자는 계약기간에 대한 "
            "연장을 청구할 수 있으며, 계약기간을 연장한 경우 공사계약 일반조건 제23조 "
            "제1항 및 제26조 제4항에 의하여 그 변경된 내용에 따라 실비를 초과하지 "
            "아니하는 범위 안에서 계약금액을 조정하여야 합니다."
        ),
        (
            "상기 표에서 정리한 본건 공사의 공기연장, 그에 따른 계약금액조정 신청과 "
            "관련하여 계약문서, 수발신 문서 등의 관련 근거자료를 토대로 귀책 사유 및 "
            "적정성을 검토하였습니다."
        ),
    ],

    # ── 종합 요약 ────────────────────────────────────────────────────────────
    "summary": old["summary"],
}

with open(DST, "w", encoding="utf-8") as f:
    json.dump(new_data, f, ensure_ascii=False, indent=2)

# 검증
with open(DST, encoding="utf-8") as f:
    verify = json.load(f)

total = len(verify["items"])
in_table = sum(1 for i in verify["items"] if i.get("show_in_table"))
print(f"items 총계: {total}")
print(f"show_in_table=True: {in_table}")
print(f"show_in_table=False: {total - in_table}")
print(f"detail_narratives: {len(verify['detail_narratives'])}개 블록")
print(f"accountability_diagram: {len(verify['accountability_diagram'])}행")
print("JSON 저장 완료:", DST)

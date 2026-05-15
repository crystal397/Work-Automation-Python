"""33_claim_extract 연동 브릿지.

contract_meta.json 의 검증된 메타를 33 의 귀책분석_data.json 에 누락 필드 한정으로 머지한다.
33 코드는 손대지 않는다. 33이 contract_meta 를 직접 import 하지도 않는다.

연동 흐름
1) 36_contract_meta: contract-meta build → contract_meta.json (메타 + _source 검증 완료)
2) 33_claim_extract: claim-extractor → 귀책분석_data.json (사건별 분석)
3) 36_contract_meta: contract-meta link-claim → 귀책분석_data.json 에 메타 머지
4) 33_claim_extract: claim-report → docx
"""

from __future__ import annotations

import json
from pathlib import Path

from contract_meta.models import ContractMeta


def to_claim_analysis_context(meta: ContractMeta) -> dict:
    """contract_meta → 33 의 data.json 에 추가될 컨텍스트 dict.

    33 의 model 이 무시하는 키는 안전하게 누적된다 (Phase2H generator 는
    알 수 없는 키를 단순 무시).
    """
    fy = meta.first_year_contract
    extended_days = sum(r.duration_diff_days.value for r in fy.revisions)
    last_end = fy.revisions[-1].period_end.value if fy.revisions else fy.initial.period_end.value

    ctx: dict = {
        "project_name": meta.project.name.value,
        "contractor": meta.contractor.name_legal.value,
        "client": meta.owner.name.value,
        "client_type": meta.project.client_type.value,
        "contract_form": meta.project.contract_form.value,
        "bidding_method": meta.project.bidding_method.value,
        "first_year_period": {
            "start": fy.initial.period_start.value.isoformat(),
            "end_initial": fy.initial.period_end.value.isoformat(),
            "end_final": last_end.isoformat(),
            "extended_days": extended_days,
        },
        "_contract_meta_source": "36_contract_meta",
        "_contract_meta_schema_version": meta.schema_version,
    }
    if meta.calculation_target is not None:
        ct = meta.calculation_target
        ctx["calculation_target"] = {
            "start": ct.period_start.value.isoformat(),
            "end": ct.period_end.value.isoformat(),
            "days": ct.days.value,
        }
    return ctx


def merge_into_claim_data(
    contract_meta_path: Path,
    claim_data_path: Path,
    *,
    overwrite: bool = False,
) -> dict:
    """contract_meta.json 을 읽어 claim_data.json 에 메타 컨텍스트를 머지.

    overwrite=False (기본): claim_data 에 이미 값이 있는 키는 보존.
    overwrite=True: contract_meta 값으로 덮어쓰기.
    """
    with open(contract_meta_path, encoding="utf-8") as f:
        meta = ContractMeta.model_validate(json.load(f))
    with open(claim_data_path, encoding="utf-8") as f:
        data = json.load(f)

    ctx = to_claim_analysis_context(meta)
    merged_count = 0
    for k, v in ctx.items():
        if overwrite or k not in data or data[k] in (None, "", [], {}):
            data[k] = v
            merged_count += 1

    with open(claim_data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return {"merged_keys": merged_count, "context": ctx}

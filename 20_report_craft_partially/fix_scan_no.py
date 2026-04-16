"""
data.json items의 scan_no null 값을 scan_result.json의 no 값으로 자동 매칭.
파일명 기준 매칭 시도 → 부분 일치 매칭 fallback.
"""
import json
from pathlib import Path

output_dir = Path(__file__).parent / "output"

for proj_dir in sorted(d for d in output_dir.iterdir() if d.is_dir()):
    scan_json = proj_dir / "scan_result.json"
    data_json = proj_dir / "귀책분석_data.json"
    if not scan_json.exists() or not data_json.exists():
        continue

    scan = json.loads(scan_json.read_text(encoding="utf-8"))
    data = json.loads(data_json.read_text(encoding="utf-8"))

    # filename → scan_no map
    fname_map: dict[str, int] = {}
    for item in scan.get("items", []):
        fp = Path(item["file_path"])
        no = item.get("no")
        if no is not None:
            fname_map[fp.name] = no

    updated = 0
    for item in data.get("items", []):
        if item.get("scan_no") is not None:
            continue
        sf = item.get("source_file", "")
        sf_name = Path(sf).name

        # 1) 정확한 파일명 매칭
        if sf_name in fname_map:
            item["scan_no"] = fname_map[sf_name]
            updated += 1
            continue

        # 2) 부분 일치 매칭 (짧은 쪽이 긴 쪽에 포함)
        for fname, no in fname_map.items():
            if sf_name and (sf_name in fname or fname in sf_name):
                item["scan_no"] = no
                updated += 1
                break

    if updated > 0:
        data_json.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"[OK] {proj_dir.name}: scan_no {updated}개 업데이트")
    else:
        null_c = sum(1 for i in data.get("items", []) if i.get("scan_no") is None)
        print(f"[--] {proj_dir.name}: 변경 없음 (null {null_c}개 남음)")

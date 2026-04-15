"""기존 CSV 파일을 xlsx로 변환"""
import pandas as pd
from pathlib import Path

for csv_path in [Path("cak_business_list.csv"), Path("cak_business_detail.csv")]:
    if not csv_path.exists():
        print(f"없음: {csv_path}")
        continue
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    xlsx_path = csv_path.with_suffix(".xlsx")
    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    print(f"변환 완료: {xlsx_path}  ({len(df):,}행, {len(df.columns)}열)")

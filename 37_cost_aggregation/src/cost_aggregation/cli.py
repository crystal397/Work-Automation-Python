"""cost-agg CLI — build / personnel / schema."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import typer
import yaml

from cost_aggregation.builder import build_cost_result
from cost_aggregation.extractors.expense_classifier import classify_directory
from cost_aggregation.extractors.personnel import extract_personnel
from cost_aggregation.models import CostResult
from cost_aggregation.report import build_cost_report

app = typer.Typer(no_args_is_help=True, add_completion=False)


@app.command()
def build(
    input_yaml: Path = typer.Argument(..., help="cost_input.yaml"),
    meta: Path = typer.Option(..., "--meta", help="contract_meta.json"),
) -> None:
    """cost_input.yaml + contract_meta.json → cost_result.json + cost_report.md."""
    result = build_cost_result(input_yaml, meta)
    out_json = input_yaml.with_name("cost_result.json")
    out_json.write_text(result.model_dump_json(by_alias=True, indent=2), encoding="utf-8")
    out_md = input_yaml.with_name("cost_report.md")
    out_md.write_text(build_cost_report(result), encoding="utf-8")
    typer.echo(f"JSON   : {out_json}")
    typer.echo(f"Report : {out_md}")
    typer.echo(
        f"총액(원도급+하도급) : {result.aggregate.grand_total.value:,}원"
    )


@app.command()
def personnel(
    xlsx: Path = typer.Argument(..., help="인원투입현황 xlsx"),
    affiliation: str = typer.Option(..., "--affiliation", help="회사명"),
    start: str = typer.Option(..., "--start", help="산정 시작일 YYYY-MM-DD"),
    end: str = typer.Option(..., "--end", help="산정 종료일 YYYY-MM-DD"),
) -> None:
    """인원투입현황 xlsx → Personnel[] (4.3.2.1 대상인원 표)."""
    ps, ws = extract_personnel(
        xlsx,
        affiliation=affiliation,
        period_start=date.fromisoformat(start),
        period_end=date.fromisoformat(end),
    )
    typer.echo(f"인원 수: {len(ps)}")
    for p in ps:
        typer.echo(f"  {p.name.value:8s} | {p.role.value:6s} | {p.period_start.value} ~ {p.period_end.value}")
    if ws:
        for w in ws:
            typer.echo(f"  ! {w}")


@app.command(name="classify-expenses")
def classify_expenses_cmd(
    root: Path = typer.Argument(..., help="영수증·전표 PDF 들이 있는 폴더"),
    out: Path = typer.Option(None, "--out", help="결과 YAML 출력 경로 (생략 시 stdout 요약만)"),
    project_root: Path = typer.Option(None, "--project-root", help="path 를 상대경로로 표기할 기준 폴더"),
) -> None:
    """경비 11비목 자동 분류 (4.3.3.1 직접계상비목)."""
    by_item = classify_directory(root.resolve())
    proot = (project_root or root).resolve()

    out_spec = {"items": []}
    for item in sorted(by_item.keys()):
        files = []
        for p in sorted(by_item[item]):
            try:
                rel = p.relative_to(proot)
            except ValueError:
                rel = p
            files.append(str(rel).replace("\\", "/"))
        out_spec["items"].append({"label": item, "files": files})

    total = sum(len(v) for v in by_item.values())
    typer.echo(f"비목 분류 결과: {total}건 / {len(by_item)}비목")
    for item, files in sorted(by_item.items(), key=lambda kv: -len(kv[1])):
        typer.echo(f"  {item}: {len(files)}건")

    if out is not None:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(yaml.safe_dump(out_spec, allow_unicode=True, sort_keys=False), encoding="utf-8")
        typer.echo(f"\nYAML: {out}")


@app.command()
def schema(
    out: Path = typer.Option(Path("schemas/cost_result.schema.json"), "--out"),
) -> None:
    """CostResult pydantic 모델 → JSON Schema."""
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(CostResult.model_json_schema(by_alias=True), ensure_ascii=False, indent=2), encoding="utf-8")
    typer.echo(f"Schema: {out}")


if __name__ == "__main__":
    app()

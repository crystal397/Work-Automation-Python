"""crash-agg CLI — extract / build / schema."""

from __future__ import annotations

import json
from pathlib import Path

import typer

from crash_aggregation.builder import build_crash_result
from crash_aggregation.builders.monthly_matrix import write_monthly_matrix
from crash_aggregation.extractors.crew_xlsx import extract_workers
from crash_aggregation.models import CrashResult, MonthlyCrash
from crash_aggregation.report import build_report

app = typer.Typer(no_args_is_help=True, add_completion=False)


@app.command()
def extract(
    xlsx: Path = typer.Argument(..., help="월별 노무비 xlsx 1개"),
    year: int = typer.Option(..., "--year"),
    month: int = typer.Option(..., "--month"),
) -> None:
    """단일 월별 xlsx → 작업자 데이터 요약."""
    workers, ws = extract_workers(xlsx, year=year, month=month)
    typer.echo(f"파일       : {xlsx.name}")
    typer.echo(f"작업자     : {len(workers)}명")
    typer.echo(f"총 공수    : {sum(w.total_manday for w in workers):.1f}일")
    typer.echo(f"총 노무비  : {sum(w.total_krw for w in workers):,}원")
    if ws:
        for w in ws:
            typer.echo(f"  ! {w}")


@app.command()
def build(
    source_dir: Path = typer.Argument(..., help="월별 xlsx 자료 루트"),
    project: str = typer.Option(..., "--project", help="프로젝트명"),
    out_dir: Path = typer.Option(Path("out"), "--out"),
    pattern: str = typer.Option("*.xlsx", "--pattern"),
    meta_ref: Path = typer.Option(None, "--meta-ref", help="(선택) contract_meta.json 경로 — JSON 에 ref 기록"),
) -> None:
    """디렉터리 내 월별 xlsx 일괄 처리 → CrashResult JSON + crash_report.md."""
    files = sorted(source_dir.glob(pattern))
    if not files:
        typer.echo(f"파일 없음: {source_dir}/{pattern}", err=True)
        raise typer.Exit(1)

    result = build_crash_result(
        list(files),
        project_name=project,
        contract_meta_ref=str(meta_ref) if meta_ref else None,
    )

    project_out = out_dir / project
    project_out.mkdir(parents=True, exist_ok=True)
    json_path = project_out / "crash_result.json"
    json_path.write_text(result.model_dump_json(by_alias=True, indent=2), encoding="utf-8")
    md_path = project_out / "crash_report.md"
    md_path.write_text(build_report(result), encoding="utf-8")

    typer.echo(f"JSON   : {json_path}")
    typer.echo(f"Report : {md_path}")
    typer.echo(f"월      : {len(result.months)} | 인원: {result.total_workers} | 공수: {result.total_manday:,.1f} | 노무비: {result.total_krw:,}원")


@app.command(name="daily-report")
def daily_report_cmd(
    xlsx: Path = typer.Argument(..., help="월별 노무비 xlsx 1개"),
    year: int = typer.Option(..., "--year"),
    month: int = typer.Option(..., "--month"),
    out: Path = typer.Option(None, "--out", help="출력 xlsx (기본: 같은 이름_매트릭스.xlsx)"),
) -> None:
    """월별 노무비 xlsx → 월별 매트릭스 xlsx (공종별 시트 + 요약)."""
    workers, _ = extract_workers(xlsx, year=year, month=month)
    monthly = MonthlyCrash(year=year, month=month, source_file=str(xlsx), workers=workers)
    target = out or xlsx.with_name(f"{xlsx.stem}_매트릭스.xlsx")
    path = write_monthly_matrix(monthly, target)
    typer.echo(f"xlsx : {path}")
    typer.echo(f"작업자 {len(workers)}명, 공종 {len({w.gongjong.value for w in workers})}개")


@app.command(name="daily-report-v12")
def daily_report_v12_cmd(
    source: Path = typer.Argument(..., help="월별 xlsx 들이 있는 폴더"),
    out: Path = typer.Option(..., "--out", help="통합 출력일보 xlsx 가 저장될 폴더"),
) -> None:
    """레거시 24 v12 의 정교한 출력일보 호출 (인쇄영역·머지셀·48행 페이지·휴일 표시).

    24_crash_construction/mandays_report_automation_v12.py (1200줄) 의 main() 을
    동적 로드해 호출한다. 38 패키지에서 24 의 검증된 출력을 그대로 활용.
    """
    from crash_aggregation.bridges import run_legacy_detailed_report

    result_path = run_legacy_detailed_report(source_dir=source, output_dir=out)
    typer.echo(f"생성: {result_path}")


@app.command(name="company-readers")
def company_readers_cmd(
    variant: str = typer.Argument(..., help="'common' (금풍·기장·대우 xlsx) | 'pdf' (스마트에스지·급여명세 PDF)"),
) -> None:
    """레거시 12_manhour_aggregation 회사별 어댑터 호출 가능 여부 점검.

    실제 호출은 Python 에서:
        from crash_aggregation import bridges
        m = bridges.load_legacy_xlsx_reader('common')
        m.detect_yearmonth(ws); m.find_date_header(ws)
    """
    from crash_aggregation.bridges import load_legacy_xlsx_reader

    mod = load_legacy_xlsx_reader(variant)
    typer.echo(f"로드 완료: {mod.__name__}")
    typer.echo(f"  파일: {mod.__file__}")
    public = [n for n in dir(mod) if not n.startswith("_")]
    typer.echo(f"  공개 식별자 {len(public)}개: {', '.join(public[:10])}{'...' if len(public) > 10 else ''}")


@app.command()
def schema(out: Path = typer.Option(Path("schemas/crash_result.schema.json"), "--out")) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(CrashResult.model_json_schema(by_alias=True), ensure_ascii=False, indent=2), encoding="utf-8")
    typer.echo(f"Schema: {out}")


if __name__ == "__main__":
    app()

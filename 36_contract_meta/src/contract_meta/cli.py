"""contract-meta CLI.

명령
- init   : 빈 YAML 양식을 out/<project>/contract_meta.yaml 로 생성
- build  : YAML → JSON 변환 + 자동 검증 + extraction_report.md 출력
- schema : models.py 에서 JSON Schema 자동 생성 (schemas/contract_meta.schema.json)
"""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from importlib import resources
from pathlib import Path

import typer
import yaml

from contract_meta.attach_scan import scan_to_appendix
from contract_meta.attachments import build_appendix, write_toc
from contract_meta.audit import sha256_of
from contract_meta.body.renderer import render_chapter
from contract_meta.bridge import merge_into_claim_data
from contract_meta.extractors import estimate_sheet, excerpts, report_rates
from contract_meta.findings import build_findings
from contract_meta.models import ContractMeta
from contract_meta.report import build_report
from contract_meta.validators.consistency import validate

app = typer.Typer(no_args_is_help=True, add_completion=False)


@app.command()
def init(
    project: str = typer.Argument(..., help="프로젝트명 (출력 폴더명)"),
    out_dir: Path = typer.Option(Path("out"), "--out", help="출력 루트"),
) -> None:
    """빈 contract_meta.yaml 양식을 생성한다."""
    project_dir = out_dir / project
    project_dir.mkdir(parents=True, exist_ok=True)

    template_path = resources.files("contract_meta.templates").joinpath("contract_meta.template.yaml")
    target = project_dir / "contract_meta.yaml"
    if target.exists():
        typer.echo(f"이미 존재: {target} (덮어쓰기 방지)")
        raise typer.Exit(1)
    target.write_text(template_path.read_text(encoding="utf-8"), encoding="utf-8")
    (project_dir / "source_excerpts").mkdir(exist_ok=True)
    typer.echo(f"양식 생성: {target}")
    typer.echo(f"채운 뒤 실행:  contract-meta build {target}")


@app.command()
def build(
    yaml_path: Path = typer.Argument(..., help="채워진 contract_meta.yaml 경로"),
) -> None:
    """YAML 입력을 검증·변환해 ContractMeta JSON + 리포트를 생성한다."""
    data = yaml.safe_load(yaml_path.read_text(encoding="utf-8"))

    # extracted_at 자동 채움
    data.setdefault("extraction", {})
    data["extraction"]["extracted_at"] = datetime.now().isoformat(timespec="seconds")

    # input_files 의 sha256 자동 채움 (빈 경우만)
    for f in data["extraction"].get("input_files", []):
        if not f.get("sha256") and f.get("path") and Path(f["path"]).exists():
            f["sha256"] = sha256_of(f["path"])

    meta = ContractMeta.model_validate(data)

    out_json = yaml_path.with_name("contract_meta.json")
    out_json.write_text(
        meta.model_dump_json(by_alias=True, indent=2),
        encoding="utf-8",
    )

    val = validate(meta)
    report_md = build_report(meta, val)
    out_report = yaml_path.with_name("extraction_report.md")
    out_report.write_text(report_md, encoding="utf-8")

    findings_md = build_findings(meta, val)
    out_findings = yaml_path.with_name("findings.md")
    if findings_md is not None:
        out_findings.write_text(findings_md, encoding="utf-8")
    elif out_findings.exists():
        out_findings.unlink()

    excerpt_dir = yaml_path.parent / "source_excerpts"
    captured, excerpt_warnings = excerpts.emit_excerpts(meta, excerpt_dir)

    typer.echo(f"JSON     : {out_json}")
    typer.echo(f"Report   : {out_report}")
    typer.echo(f"Findings : {out_findings if findings_md else '(검출된 이슈 없음)'}")
    typer.echo(f"Excerpts : {len(captured)}건 캡쳐 ({excerpt_dir})")
    if excerpt_warnings:
        for w in excerpt_warnings:
            typer.echo(f"  ! {w}")
    typer.echo(
        f"검증     : 통과 {len(val.passed)}건 / 실패 {len(val.failed)}건 / 경고 {len(val.warnings)}건"
    )
    if val.failed:
        raise typer.Exit(1)


@app.command(name="extract-rates")
def extract_rates_cmd(
    source: Path = typer.Argument(..., help="보고서 PDF 또는 산출내역서 xlsx"),
    out: Path = typer.Option(None, "--out", help="결과 YAML 출력 경로 (생략 시 stdout)"),
) -> None:
    """보고서 PDF / 산출내역서 xlsx 에서 rates 4종 자동 추출."""
    suffix = source.suffix.lower()
    if suffix == ".pdf":
        rates, warnings = report_rates.extract_rates(source)
    elif suffix in (".xlsx", ".xlsm"):
        rates, warnings = estimate_sheet.extract_rates_auto(source)
    else:
        typer.echo(f"지원하지 않는 확장자: {suffix}", err=True)
        raise typer.Exit(1)

    block = {"rates": {}}
    for r in rates:
        block["rates"][r.field] = {
            "value": r.value,
            "_source": r.source.model_dump(),
        }

    yaml_text = yaml.safe_dump(block, allow_unicode=True, sort_keys=False)
    if warnings:
        yaml_text = "# warnings:\n" + "".join(f"#   - {w}\n" for w in warnings) + yaml_text

    if out:
        out.write_text(yaml_text, encoding="utf-8")
        typer.echo(f"Rates: {out}")
    else:
        typer.echo(yaml_text)


@app.command(name="render")
def render_cmd(
    meta_path: Path = typer.Argument(..., help="contract_meta.json"),
    chapter: str = typer.Option("cover,chapter_2,chapter_3_2", "--chapters", help="쉼표 구분 챕터명"),
    out: Path = typer.Option(None, "--out", help="출력 .md (기본: meta_path 디렉터리에 body.md)"),
    docx_out: Path = typer.Option(None, "--docx", help="(선택) docx 동시 출력 경로"),
) -> None:
    """contract_meta.json → 본문 마크다운 (Jinja2 템플릿) + 선택적 docx 변환."""
    meta = ContractMeta.model_validate(json.loads(meta_path.read_text(encoding="utf-8")))
    parts = []
    for c in [x.strip() for x in chapter.split(",") if x.strip()]:
        parts.append(render_chapter(meta, c))
    body = "\n\n---\n\n".join(parts)
    target = out or meta_path.with_name("body.md")
    target.write_text(body, encoding="utf-8")
    typer.echo(f"Body : {target}")

    if docx_out is not None:
        from contract_meta.body.to_docx_kicm import md_to_docx_kicm
        docx_path = md_to_docx_kicm(
            body, docx_out,
            project_name=meta.project.name.value,
        )
        typer.echo(f"Docx : {docx_path}  (KICM 양식: 맑은고딕 + 머리말 + 페이지번호 + 인용박스 + 표스타일)")


@app.command(name="scan-attach")
def scan_attach_cmd(
    root: Path = typer.Argument(..., help="수신자료 루트 폴더 (PDF 들이 들어있는)"),
    out: Path = typer.Option(..., "--out", help="생성될 appendix.yaml 경로"),
    project_root: Path = typer.Option(None, "--project-root", help="path 를 상대경로로 표기할 기준 폴더"),
    absolute: bool = typer.Option(True, "--absolute/--relative", help="path 를 절대경로로 저장 (기본). --relative 로 상대경로 사용"),
) -> None:
    """수신자료 폴더 스캔 → appendix.yaml 자동 생성 (5.1~5.5 휴리스틱 분류).

    --absolute (기본): yaml 의 path 를 절대경로로 저장 → attach 호출 시 cwd 무관.
    --relative: project_root 기준 상대경로 (이식성 우선, 다른 PC 로 yaml 옮길 때).
    """
    spec = scan_to_appendix(root, project_root=project_root, absolute_paths=absolute)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(yaml.safe_dump(spec, allow_unicode=True, sort_keys=False), encoding="utf-8")
    n_files = sum(len(s["files"]) for s in spec["sections"])
    typer.echo(f"appendix.yaml : {out}")
    typer.echo(f"  섹션 {len(spec['sections'])}개 / 파일 {n_files}개")
    for s in spec["sections"]:
        typer.echo(f"    {s['title']}: {len(s['files'])}건")


@app.command(name="attach")
def attach_cmd(
    spec: Path = typer.Argument(..., help="첨부 yaml 명세 (sections/files)"),
    out: Path = typer.Option(None, "--out", help="출력 5_appendix.pdf 경로"),
) -> None:
    """첨부 yaml → 합쳐진 5_appendix.pdf + appendix_toc.md."""
    pdf_out = out or spec.with_name("5_appendix.pdf")
    pdf_path, sections = build_appendix(spec, pdf_out)
    toc_path = pdf_out.with_name("appendix_toc.md")
    write_toc(sections, toc_path)
    typer.echo(f"PDF  : {pdf_path}  ({sum(1 for _ in sections)} 섹션)")
    typer.echo(f"TOC  : {toc_path}")
    for s in sections:
        typer.echo(f"  {s.title}: p.{s.start_page}~{s.end_page}")
    skipped_path = pdf_out.with_name("appendix_skipped.txt")
    if skipped_path.exists():
        n = sum(1 for _ in skipped_path.read_text(encoding="utf-8").splitlines() if _)
        typer.echo(f"\n⚠️  합본 제외 {n}건: {skipped_path}")
        typer.echo("    (.docx 등 비-PDF 는 별도 변환 후 재실행하거나, 그대로 별첨)")


@app.command(name="laws-import")
def laws_import_cmd(
    meta_path: Path = typer.Argument(..., help="contract_meta.json — project.bid_announcement_date 사용"),
    out: Path = typer.Option(None, "--out", help="출력 docx (기본: 같은 폴더의 5_4_적용법령.docx)"),
    oc: str = typer.Option(None, "--oc", help="법제처 API OC (이메일 ID). 미지정 시 22 의 .env 사용"),
    bid_date: str = typer.Option(None, "--bid-date", help="입찰공고일 YYYY-MM-DD (meta 의 값을 덮어쓰기)"),
) -> None:
    """입찰공고일 기준 최근 법령 검색 + Word 출력 (보고서 5.4 적용 법령 자동).

    22_laws_import 의 LawMatcher + WordReportGenerator 를 동적 로드해 호출.
    출력 docx 는 36의 5장 첨부 워크플로우(`scan-attach` + `attach`)가 그대로 흡수
    가능 — 파일명에 '법령' 포함되어 attach_scan.py 가 5.4 로 자동 분류.
    """
    from contract_meta.bridges import run_laws_import
    from datetime import date

    # bid_date 결정 우선순위: CLI 인자 > contract_meta.project.bid_announcement_date
    if bid_date:
        bd = date.fromisoformat(bid_date)
    else:
        meta = ContractMeta.model_validate(json.loads(meta_path.read_text(encoding="utf-8")))
        if meta.project.bid_announcement_date is None:
            typer.echo(
                "오류: contract_meta.project.bid_announcement_date 가 없습니다. "
                "yaml 에 채우거나 --bid-date YYYY-MM-DD 로 전달하세요.",
                err=True,
            )
            raise typer.Exit(1)
        bd = meta.project.bid_announcement_date.value

    target = out or meta_path.with_name("5_4_적용법령.docx")

    def progress(i, total, name):
        typer.echo(f"  [{i+1}/{total}] {name}")

    typer.echo(f"입찰공고일: {bd}")
    typer.echo(f"법령 매칭 시작...")
    result_path = run_laws_import(
        bid_date=bd,
        out_path=target,
        oc=oc,
        progress_callback=progress,
    )
    typer.echo(f"\nDocx : {result_path}")
    typer.echo(f"5장 첨부 자동 흡수 — 다음 실행 시 `scan-attach` 가 5.4 로 분류:")
    typer.echo(f"  contract-meta scan-attach <{result_path.parent}> --out appendix.yaml")


@app.command(name="link-claim")
def link_claim_cmd(
    contract_meta_path: Path = typer.Argument(..., help="contract_meta.json (build 산출물)"),
    claim_data_path: Path = typer.Argument(..., help="33_claim_extract 의 귀책분석_data.json"),
    overwrite: bool = typer.Option(False, "--overwrite", help="기존 값 덮어쓰기"),
) -> None:
    """contract_meta.json → 33_claim_extract 의 data.json 에 메타 컨텍스트 머지."""
    result = merge_into_claim_data(contract_meta_path, claim_data_path, overwrite=overwrite)
    typer.echo(f"머지 키 수: {result['merged_keys']}")
    typer.echo(f"대상 파일 : {claim_data_path}")


@app.command()
def schema(
    out: Path = typer.Option(Path("schemas/contract_meta.schema.json"), "--out", help="출력 경로"),
) -> None:
    """ContractMeta pydantic 모델 → JSON Schema 파일."""
    out.parent.mkdir(parents=True, exist_ok=True)
    schema_obj = ContractMeta.model_json_schema(by_alias=True)
    out.write_text(json.dumps(schema_obj, ensure_ascii=False, indent=2), encoding="utf-8")
    typer.echo(f"Schema: {out}")


if __name__ == "__main__":
    app()

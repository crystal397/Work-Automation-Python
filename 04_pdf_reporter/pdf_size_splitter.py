"""
PDF 파일을 지정한 크기(기본 20MB) 미만으로 분할하는 스크립트
사용법: python pdf_splitter.py input.pdf
"""

import os
import sys
import io
from pypdf import PdfReader, PdfWriter


def split_pdf_by_size(input_path: str, max_size_mb: float = 2.0, output_dir: str = None):
    """
    PDF를 max_size_mb 미만의 크기로 분할하여 저장합니다.

    Args:
        input_path: 입력 PDF 파일 경로
        max_size_mb: 분할 기준 크기 (MB), 기본값 20MB
        output_dir: 출력 디렉토리 (None이면 입력 파일과 같은 위치)
    """
    max_size_bytes = max_size_mb * 1024 * 1024

    if not os.path.exists(input_path):
        print(f"❌ 파일을 찾을 수 없습니다: {input_path}")
        return

    file_size = os.path.getsize(input_path)
    print(f"📄 원본 파일: {input_path} ({file_size / 1024 / 1024:.2f} MB)")

    if file_size < max_size_bytes:
        print(f"✅ 파일 크기가 이미 {max_size_mb}MB 미만입니다. 분할이 필요하지 않습니다.")
        return

    # 출력 디렉토리 설정
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(input_path))
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(input_path))[0]

    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    print(f"📑 총 페이지 수: {total_pages}")

    part_num = 1
    start_page = 0
    saved_files = []

    while start_page < total_pages:
        writer = PdfWriter()
        current_page = start_page

        # 페이지를 하나씩 추가하면서 크기 체크
        while current_page < total_pages:
            writer.add_page(reader.pages[current_page])

            # 메모리에 써서 현재 크기 측정
            buffer = io.BytesIO()
            writer.write(buffer)
            current_size = buffer.tell()

            if current_size >= max_size_bytes:
                # 마지막으로 추가한 페이지가 초과 원인 → 제거
                if len(writer.pages) > 1:
                    # 마지막 페이지를 제외하고 다시 저장
                    writer2 = PdfWriter()
                    for i in range(len(writer.pages) - 1):
                        writer2.add_page(reader.pages[start_page + i])

                    output_path = os.path.join(output_dir, f"{base_name}_part{part_num}.pdf")
                    with open(output_path, "wb") as f:
                        writer2.write(f)

                    actual_size = os.path.getsize(output_path)
                    print(f"  ✂️  Part {part_num} 저장: {output_path} ({actual_size / 1024 / 1024:.2f} MB, {len(writer2.pages)} 페이지)")
                    saved_files.append(output_path)

                    start_page = current_page  # 현재 페이지부터 다음 파트 시작
                    part_num += 1
                    break
                else:
                    # 단일 페이지만으로도 크기 초과 → 어쩔 수 없이 저장
                    print(f"  ⚠️  Page {current_page + 1}은 단독으로 {current_size / 1024 / 1024:.2f} MB입니다. 그대로 저장합니다.")
                    output_path = os.path.join(output_dir, f"{base_name}_part{part_num}.pdf")
                    with open(output_path, "wb") as f:
                        writer.write(f)

                    saved_files.append(output_path)
                    print(f"  ✂️  Part {part_num} 저장: {output_path} ({os.path.getsize(output_path) / 1024 / 1024:.2f} MB, 1 페이지)")
                    start_page = current_page + 1
                    part_num += 1
                    break
            else:
                current_page += 1
        else:
            # while 루프가 break 없이 끝남 → 남은 페이지들 저장
            if len(writer.pages) > 0:
                output_path = os.path.join(output_dir, f"{base_name}_part{part_num}.pdf")
                with open(output_path, "wb") as f:
                    writer.write(f)

                actual_size = os.path.getsize(output_path)
                print(f"  ✂️  Part {part_num} 저장: {output_path} ({actual_size / 1024 / 1024:.2f} MB, {len(writer.pages)} 페이지)")
                saved_files.append(output_path)
            break

    print(f"\n🎉 완료! 총 {len(saved_files)}개 파일로 분할되었습니다.")
    print(f"📁 저장 위치: {output_dir}")
    return saved_files


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python pdf_splitter.py <PDF파일경로> [최대크기MB] [출력디렉토리]")
        print("예시:  python pdf_splitter.py my_document.pdf 20 ./output")
        sys.exit(1)

    input_pdf = sys.argv[1]
    max_mb = float(sys.argv[2]) if len(sys.argv) > 2 else 20.0
    out_dir = sys.argv[3] if len(sys.argv) > 3 else None

    split_pdf_by_size(input_pdf, max_size_mb=max_mb, output_dir=out_dir)

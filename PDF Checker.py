import fitz  # PyMuPDF
import os
import sys
import io
from contextlib import redirect_stderr
from pypdf import PdfReader
import shutil

# MuPDF 오류 메시지 차단
def disable_mupdf_errors():
    sys.stderr = open(os.devnull, "w")

# PDF 파일이 손상되었는지 확인
def check_pdf_integrity(file_path):
    disable_mupdf_errors()
    error_buffer = io.StringIO()

    try:
        # 파일 크기가 0바이트인지 확인
        if os.path.getsize(file_path) == 0:
            return "손상됨 (파일 크기가 0바이트)"

        # MuPDF로 열어서 검사
        with redirect_stderr(error_buffer):
            with fitz.open(file_path) as doc:
                for _ in doc:
                    pass

        # pypdf로 페이지 데이터 접근 확인
        with open(file_path, "rb") as f:
            reader = PdfReader(f)
            _ = reader.pages  # 페이지 접근 테스트

        # 파일 끝에 %%EOF가 있는지 확인
        with open(file_path, "rb") as f:
            f.seek(-10, 2)  # 파일 끝에서 10바이트 전으로 이동
            if b"%%EOF" not in f.read():
                return "손상됨 (EOF 마커 없음)"

        return "정상"  # 모든 검사를 통과하면 정상

    except Exception as e:
        return f"손상됨 ({e})"

# 폴더 내 모든 PDF 검사 후 결과를 파일로 저장
def check_pdfs_in_folder(target_folder):
    pdf_files = [f for f in os.listdir(target_folder) if f.lower().endswith('.pdf')]
    total_files = len(pdf_files)
    results = []

    # "OK" 폴더와 "corrupted" 폴더 경로 생성
    ok_folder = os.path.join(target_folder, "OK")
    corrupted_folder = os.path.join(target_folder, "corrupted")
    os.makedirs(ok_folder, exist_ok=True)
    os.makedirs(corrupted_folder, exist_ok=True)

    for idx, pdf in enumerate(pdf_files):
        file_path = os.path.join(target_folder, pdf)
        status = check_pdf_integrity(file_path)
        
        # 상태에 따라 파일을 복사
        if status == "정상":
            shutil.copy(file_path, os.path.join(ok_folder, pdf))  # 정상 파일은 OK 폴더로 복사
        else:
            shutil.copy(file_path, os.path.join(corrupted_folder, pdf))  # 손상된 파일은 corrupted 폴더로 복사
        
        results.append(f"{pdf}: {status}")

        # 진행 상황 출력
        progress = (idx + 1) / total_files * 100
        print(f"진행 중: {progress:.2f}%")

    # 검사 결과를 파일로 저장
    output_file = os.path.join(SCRIPT_DIR, "pdf_check_results.txt")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(results))

    print(f"📄 검사 결과가 '{output_file}' 파일에 저장되었습니다.")

# 검사할 폴더 경로 설정
target_folder = "C:/TEST~~"
check_pdfs_in_folder(target_folder)
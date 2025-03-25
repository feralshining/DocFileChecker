import os
import shutil
import openpyxl
import xlrd
from contextlib import redirect_stderr
import sys

# 현재 실행 중인 스크립트의 폴더 경로
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 엑셀 오류 메시지 차단
def disable_excel_errors():
    sys.stderr = open(os.devnull, "w")

# 엑셀 파일이 손상되었는지 확인
def check_excel_integrity(file_path):
    disable_excel_errors()

    try:
        # 파일 크기가 0바이트인지 확인
        if os.path.getsize(file_path) == 0:
            return "손상됨 (파일 크기가 0바이트)"

        # 엑셀 파일 열기 (XLSX 형식)
        if file_path.lower().endswith(".xlsx"):
            try:
                openpyxl.load_workbook(file_path)  # XLSX 파일 열기 시도
            except Exception:
                return "손상됨 (XLSX 파일 열기 실패)"
        
        # 엑셀 파일 열기 (XLS 형식)
        elif file_path.lower().endswith(".xls"):
            try:
                xlrd.open_workbook(file_path)  # XLS 파일 열기 시도
            except Exception:
                return "손상됨 (XLS 파일 열기 실패)"

        return "정상"  # 모든 검사를 통과하면 정상

    except Exception as e:
        return f"손상됨 ({e})"

# 폴더 내 모든 엑셀 파일 검사 후 결과를 파일로 저장
def check_excels_in_folder(target_folder):
    excel_files = [f for f in os.listdir(target_folder) if f.lower().endswith(('.xls', '.xlsx'))]
    total_files = len(excel_files)
    results = []

    # "OK" 폴더와 "corrupted" 폴더 경로 생성
    ok_folder = os.path.join(target_folder, "OK")
    corrupted_folder = os.path.join(target_folder, "corrupted")
    os.makedirs(ok_folder, exist_ok=True)
    os.makedirs(corrupted_folder, exist_ok=True)

    for idx, excel in enumerate(excel_files):
        file_path = os.path.join(target_folder, excel)
        status = check_excel_integrity(file_path)
        
        # 상태에 따라 파일을 복사
        if status == "정상":
            shutil.copy(file_path, os.path.join(ok_folder, excel))  # 정상 파일은 OK 폴더로 복사
        else:
            shutil.copy(file_path, os.path.join(corrupted_folder, excel))  # 손상된 파일은 corrupted 폴더로 복사
        
        results.append(f"{excel}: {status}")

        # 진행 상황 출력
        progress = (idx + 1) / total_files * 100
        print(f"진행 중: {progress:.2f}%")

    # 검사 결과를 파일로 저장 (SCRIPT_DIR에 저장)
    output_file = os.path.join(SCRIPT_DIR, "excel_check_results.txt")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(results))

    print(f"📄 검사 결과가 '{output_file}' 파일에 저장되었습니다.")

# 검사할 폴더 경로 설정
target_folder = "C:/TEST~~"
check_excels_in_folder(target_folder)

import os
import olefile
import zlib
import struct
import re
import unicodedata
import shutil

# 현재 실행 중인 스크립트의 폴더 경로
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 중국어 제거
def remove_chinese_characters(s: str):   
    return re.sub(r'[\u4e00-\u9fff]+', '', s)

# 바이트 문자열 제거
def remove_control_characters(s):    
    return "".join(ch for ch in s if unicodedata.category(ch)[0]!="C")

class HWPExtractor(object):
    FILE_HEADER_SECTION = "FileHeader"
    HWP_SUMMARY_SECTION = "\x05HwpSummaryInformation"
    SECTION_NAME_LENGTH = len("Section")
    BODYTEXT_SECTION = "BodyText"
    HWP_TEXT_TAGS = [67]



    def __init__(self, filename):
        self._ole = self.load(filename)
        self._dirs = self._ole.listdir()

        self._valid = self.is_valid(self._dirs)
        if not self._valid:
            raise Exception("Not Valid HwpFile")

        self._compressed = self.is_compressed(self._ole)
        self.text = self._get_text()

    def load(self, filename):
        return olefile.OleFileIO(filename)

    def is_valid(self, dirs):
        if [self.FILE_HEADER_SECTION] not in dirs:
            return False
        return [self.HWP_SUMMARY_SECTION] in dirs

    def is_compressed(self, ole):
        header = self._ole.openstream("FileHeader")
        header_data = header.read()
        return (header_data[36] & 1) == 1

    def get_body_sections(self, dirs):
        m = []
        for d in dirs:
            if d[0] == self.BODYTEXT_SECTION:
                m.append(int(d[1][self.SECTION_NAME_LENGTH:]))
        return ["BodyText/Section" + str(x) for x in sorted(m)]

    def get_text(self):
        return self.text

    def _get_text(self):
        sections = self.get_body_sections(self._dirs)
        text = ""
        for section in sections:
            text += self.get_text_from_section(section)
            text += "\n"
        self.text = text
        return self.text

    def get_text_from_section(self, section):
        bodytext = self._ole.openstream(section)
        data = bodytext.read()

        unpacked_data = zlib.decompress(data, -15) if self._compressed else data
        size = len(unpacked_data)

        i = 0
        text = ""

        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            level = (header >> 10) & 0x3ff
            rec_len = (header >> 20) & 0xfff

            if rec_type in self.HWP_TEXT_TAGS:
                rec_data = unpacked_data[i + 4:i + 4 + rec_len]
                decode_text = rec_data.decode('utf-16')

                # 문자열 정제
                res = remove_control_characters(remove_chinese_characters(decode_text))

                text += res + "\n"

            i += 4 + rec_len

        return text

def get_text(filename):
    try:
        hwp = HWPExtractor(filename) 
        text = hwp.get_text()
        return text.strip()  # 공백 제거 후 리턴
    except:
        return None  # 파일이 손상되었거나 읽을 수 없는 경우 None 반환

def scan_hwp_files(directory):
    result_file = os.path.join(SCRIPT_DIR, "hwp_check_result.txt")
    
    # OK, corrupted 폴더가 없으면 생성
    if not os.path.exists(os.path.join(directory, "OK")):
        os.makedirs(os.path.join(directory, "OK"))
    if not os.path.exists(os.path.join(directory, "corrupted")):
        os.makedirs(os.path.join(directory, "corrupted"))
    
    with open(result_file, "w", encoding="utf-8") as f:
        for filename in os.listdir(directory):
            if filename.lower().endswith(".hwp"):
                filepath = os.path.join(directory, filename)
                extracted_text = get_text(filepath)

                # 파일 상태 판별
                if extracted_text:
                    status = "정상"
                    # 정상 파일은 OK 폴더로 복사
                    try:
                        shutil.copy(filepath, os.path.join(directory, "OK", filename))
                    except PermissionError:
                        print(f"파일 잠금 오류: {filename}")
                        continue  # 다른 프로세스가 파일을 사용 중인 경우 건너뜁니다.
                else:
                    status = "손상됨"
                    # 손상된 파일은 corrupted 폴더로 복사
                    try:
                        shutil.copy(filepath, os.path.join(directory, "corrupted", filename))
                    except PermissionError:
                        print(f"파일 잠금 오류: {filename}")
                        continue  # 다른 프로세스가 파일을 사용 중인 경우 건너뜁니다.

                # 결과 기록
                f.write(f"{filename}: {status}\n")
                print(f"{filename}: {status}")

    print(f"\n검사 완료! 결과는 {result_file}에 저장됨.")

# 실행할 폴더 경로 지정
target_folder = "C:\TEST~~"

# HWP 파일 검사 실행
scan_hwp_files(target_folder)

## 📄 DocFileChecker 📄
> hwp(한글), pdf, xlsx(엑셀) 파일들의 손상 여부를 판단하는 도구

----
###  💡  도구 사용법
- **HWP Checker** : `target_folder 변수에 위치 지정. 정상 파일은 "OK" 폴더로 복사, 손상된 파일은 "corrupted" 폴더로 복사됨. HWP Checker가 있는 위치에 hwp_check_result.txt 생성됨.` 

- **PDF Checker with MuPDF & pypdf** : `target_folder 변수에 위치 지정. 정상 파일은 "OK" 폴더로 복사, 손상된 파일은 "corrupted" 폴더로 복사됨. PDF Checker가 있는 위치에 pdf_check_results.txt 생성됨.` 

- **XLSX Checker with openpyxl & xlrd** : `target_folder 변수에 위치 지정. 정상 파일은 "OK" 폴더로 복사, 손상된 파일은 "corrupted" 폴더로 복사됨. XLSX Checker가 있는 위치에 xlsx_check_results.txt 생성됨.` 

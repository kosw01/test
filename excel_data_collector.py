import os
import pandas as pd
import sys

# 실행 파일인지 스크립트인지 확인
if getattr(sys, 'frozen', False):
    # 실행 파일로 실행된 경우
    script_folder = os.path.dirname(sys.executable)
else:
    # 스크립트로 실행된 경우
    script_folder = os.path.dirname(os.path.abspath(__file__))

output_excel_filename = os.path.join(script_folder, 'output.xlsx')
result_df = pd.DataFrame() #빈 데이터프레임

print(f"현재 작업 폴더: {script_folder}")

for filename in os.listdir(script_folder):
    if filename.endswith('.xlsx') and filename != 'output.xlsx':  # 엑셀 파일인 경우에만 처리 (output.xlsx 제외)
        file_path = os.path.join(script_folder, filename)
        print(f"처리 중인 파일: {filename}")
        
        # 엑셀 파일에서 시트 이름 및 두 번째 열 데이터 추출
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            sheet_df = pd.read_excel(file_path, sheet_name)
            
            # 두 번째 열만 추출
            second_column = sheet_df.iloc[:, 1]
            
            result_df[filename + '_' + sheet_name] = second_column

# 파일저장
result_df.to_excel(output_excel_filename, index=False)
print(f"결과가 {output_excel_filename}에 저장되었습니다.")

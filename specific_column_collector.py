import os
import pandas as pd

#파일 경로 & 저장 파일 명
script_folder = os.path.dirname(os.path.abspath(__file__))
output_csv_filename = 'collect_data.csv'
result_df = pd.DataFrame() #빈 데이터프레임

for filename in os.listdir(script_folder):
    if filename.endswith('.csv'):  # 엑셀 파일인 경우에만 처리
        file_path = os.path.join(script_folder, filename)
        df = pd.read_csv(file_path)
        gp_s3m_01_z = df['AC']

        result_df[filename] = gp_s3m_01_z

# 파일저장
result_df.to_csv(output_csv_filename, index=False)

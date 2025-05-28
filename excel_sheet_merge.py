import pandas as pd
import os

def analyze_measure_date(file_path):
    """
    각 시트의 'Measure Date' 컬럼 데이터 특성을 분석하는 함수
    
    Parameters:
    file_path (str): 엑셀 파일의 경로
    """
    try:
        # 엑셀 파일 읽기
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        print("\n=== Measure Date 컬럼 분석 ===")
        for sheet in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            
            # Measure Date 컬럼이 있는지 확인
            if 'Measure Date' in df.columns:
                print(f"\n[{sheet} 시트]")
                print(f"데이터 타입: {df['Measure Date'].dtype}")
                print(f"고유값 개수: {df['Measure Date'].nunique()}")
                print(f"시작 날짜: {df['Measure Date'].min()}")
                print(f"종료 날짜: {df['Measure Date'].max()}")
                print(f"샘플 데이터 (처음 5개):")
                print(df['Measure Date'].head())
            else:
                print(f"\n[{sheet} 시트] - Measure Date 컬럼이 없습니다.")
                
    except Exception as e:
        print(f"오류가 발생했습니다: {str(e)}")

def merge_excel_sheets(file_path, join_type='outer'):
    """
    엑셀 파일의 모든 시트를 'Measure Date' 컬럼을 기준으로 병합하는 함수
    
    Parameters:
    file_path (str): 엑셀 파일의 경로
    join_type (str): 병합 방식 ('inner', 'outer', 'left', 'right')
    
    Returns:
    None: 결과를 CSV 파일로 저장
    """
    try:
        # 엑셀 파일 읽기
        excel_file = pd.ExcelFile(file_path)
        
        # 시트 이름 리스트 출력
        sheet_names = excel_file.sheet_names
        print("엑셀 파일에 포함된 시트 목록:")
        for sheet in sheet_names:
            print(f"- {sheet}")
        
        # 첫 번째 시트를 기준 데이터프레임으로 설정
        merged_df = pd.read_excel(file_path, sheet_name=sheet_names[0])
        
        # 나머지 시트들을 순차적으로 병합
        for sheet in sheet_names[1:]:
            df = pd.read_excel(file_path, sheet_name=sheet)
            # outer join으로 병합
            merged_df = pd.merge(merged_df, df, on='Measure Date', how=join_type)
        
        # 결과를 CSV 파일로 저장
        output_path = os.path.splitext(file_path)[0] + f'_{join_type}_merged.csv'
        merged_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"\n병합된 데이터가 다음 경로에 저장되었습니다: {output_path}")
        print(f"사용된 병합 방식: {join_type}")
        print(f"최종 데이터프레임 크기: {merged_df.shape}")
        
    except Exception as e:
        print(f"오류가 발생했습니다: {str(e)}")

# 사용 예시
if __name__ == "__main__":
    # 현재 스크립트와 같은 디렉토리에 있는 엑셀 파일 사용
    file_path = '거금대교 계측데이터 다운로드.xlsx'
    # Measure Date 컬럼 분석
    analyze_measure_date(file_path)
    # 병합 방식을 선택할 수 있습니다: 'inner', 'outer', 'left', 'right'
    merge_excel_sheets(file_path, join_type='outer') 

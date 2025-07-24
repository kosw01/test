import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator
import pandas as pd
from scipy import stats

def preprocess(df, df1, stdev):
    for i in range(len(stdev)):
        df[stdev[i]] = df1[stdev[i]]
    return df

def plot_graph(df, month, br_name):
    for col in df.columns[1:]:
        plt.figure(figsize=(7.4, 3))
        plt.plot(df['date'], df[col])
        dateFmt = mdates.DateFormatter('%Y-%m-%d')
        plt.gca().xaxis.set_major_formatter(dateFmt)
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=3))
        plt.gca().yaxis.set_major_locator(MaxNLocator(nbins=10))
        plt.xlabel('Date')
        plt.ylabel(col)
        plt.title(col)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(f'04_report_monthly/2025_0{month+1}/{br_name}/{col}.png')
        plt.close()

def main(df, df1, stdev, month, br_name):
    preprocess(df, df1, stdev)
    plot_graph(df, month, br_name)

def bridge_match(n):
    if n == 1:
        stdev = [
                'SL_HJG_HGZ', 'SL_HJG_HGN', 'SL_HJG_HGE', 'SL_HJQ_HBZ', 'SL_HJQ_HBY', 'SL_HJH_HBZ', 'SL_HJK_HBZ', 'SL_HJK_HBY',
                'SL_HJJ_HBZ', 'SL_HJF_HBX', 'SL_HJE_HBY', 'SL_HJE_HBX', 'SL_HJD_HBX', 'SL_HJC_HBY', 'SL_HJC_HBX', 'SL_HJP_HBY'
                ]
        br_name = 'hangju'
        return br_name, stdev
    elif n == 2:
        stdev = ['TA_1L','TA_1T','TA_1V']
        br_name = 'gayang'
        return br_name, stdev
    elif n == 3:
        stdev = ['AC_S2M_01_Y','AC_S2M_01_Z','EQK_WBG_X','EQK_WBG_Y','EQK_WBG_Z','EQK_WBQ_Y','EQK_WBQ_Z','EQK_WBH_Z','EQK_WBA_X','EQK_WBA_Y','EQK_WBC_X','EQK_WBC_Y','EQK_WBP_Y']
        br_name = 'worldcup'
        return br_name, stdev
    elif n == 4:
        stdev = ['AC_S3_01']
        br_name = 'sungsan'
        return br_name, stdev
    elif n == 5:
        stdev = ['AC_01','AC_02']
        br_name = 'yanghwa'
        return br_name, stdev
    elif n == 6:
        stdev = []
        br_name = 'seogang'
        return br_name, stdev
    elif n == 7:
        stdev = []
        br_name ='wonhyo'
        return br_name, stdev
    elif n == 8:
        stdev = ['AC_P04P05_01','AC_P04P05_02','AC_P04P05_03','AC_P12P13_01','AC_P12P13_02','AC_P12P13_03']
        br_name ='hangang'
        return br_name, stdev
    elif n == 9:
        stdev = ['AC_P8P9_01','AC_P8P9_02','AC_P8P9_03','AC_P8P9_04']
        br_name ='sungsu'
        return br_name, stdev
    elif n == 10:
        stdev = ['AC_P36P37_Z_01', 'AC_P37P38_Z_01']
        br_name = 'chongdam'
        return br_name, stdev
    elif n == 11:
        stdev = ['SL_OPG_HGZ','SL_OPG_HGN','SL_OPG_HGE','SL_OPQ_HBZ','SL_OPQ_HBY','SL_OPH_HBZ','SL_OPK_HBZ','SL_OPK_HBY']
        br_name ='olympic'
        return br_name, stdev
    elif n == 12:
        stdev = ['EST_Z','EST_Y','EST_X','ACC3_Z','ACC4_Z','ACC5_Y','ACC6_X','ACC7_Z']
        br_name ='amsa'
        return br_name, stdev
    elif n == 13:
        stdev = ['AC_01_Y','AC_01_Z','AC_02_Y','AC_02_Z']
        br_name ='setgang'
        return br_name, stdev
    elif  n == 14:
        stdev  = ['AC_S1Q2_01_Z', 'AC_S4Q2_01_Z', 'AC_S7Q2_01_Z', 'AC_S10Q2_01_Z', 'AC_S13Q2_01_Z', 'AC_S16Q2_01_Z', 'AC_S17Q2_01_Z'] 
        br_name = 'cheonho'
        return br_name, stdev
    elif  n == 15:
        stdev  = ['AC_S1Q2_01_Z', 'AC_S4Q2_01_Z', 'AC_S7Q2_01_Z', 'AC_S10Q2_01_Z', 'AC_S13Q2_01_Z', 'AC_S16Q2_01_Z', 'AC_S17Q2_01_Z']  
        br_name = 'yeongdong'
        return br_name, stdev
    elif  n == 16:
        stdev  = ['JA1_X','JA1_Y','JA1_Z','JA1_T','JA2_X','JA2_Y','JA2_Z']
        br_name = 'jamsu'
        return br_name, stdev
    else:
        raise ValueError("Invalid input. n should be between 1 and 16.")

def process_bridge(n):
    """특정 다리 번호에 대한 데이터 처리"""
    print(f"\n{'='*50}")
    print(f"다리 번호 {n} 처리 시작...")
    print(f"{'='*50}")
    
    br_name, stdev = bridge_match(n)
    print(f"다리 이름: {br_name}")
    print(f"처리할 센서 수: {len(stdev)}")
    
    # Average 데이터 처리
    file_paths = [
        f'00_data/02_preprocess/2025_01/{br_name}_average.csv',
        f'00_data/02_preprocess/2025_02/{br_name}_average.csv',
        f'00_data/02_preprocess/2025_03/{br_name}_average.csv',
        f'00_data/02_preprocess/2025_04/{br_name}_average.csv',
        f'00_data/02_preprocess/2025_05/{br_name}_average.csv',
        f'00_data/02_preprocess/2025_06/{br_name}_average.csv',
    ]

    output_file = f'03_append/{br_name}_average.csv'

    # 데이터프레임 초기화
    df_list = []

    # 각 파일을 읽어서 리스트에 추가
    for file_path in file_paths:
        try:
            df = pd.read_csv(file_path, na_values='Null')
            df_list.append(df)
            print(f"  {file_path} 로드 완료")
        except FileNotFoundError:
            print(f"  경고: {file_path} 파일을 찾을 수 없습니다.")

    if df_list:
        # 모든 데이터프레임을 합치기
        df_app = pd.concat(df_list, ignore_index=True)
        print(f"  결측값 확인: {df_app.isna().sum().sum()}개")
        
        # 결과를 CSV 파일로 저장
        df_app.to_csv(output_file, index=False)
        print(f"  {br_name}_average.csv 저장 완료")
    else:
        print(f"  경고: {br_name}의 average 데이터가 없습니다.")

    # Stdev 데이터 처리
    file_paths = [
        f'00_data/02_preprocess/2025_01/{br_name}_stdev.csv',
        f'00_data/02_preprocess/2025_02/{br_name}_stdev.csv',
        f'00_data/02_preprocess/2025_03/{br_name}_stdev.csv',
        f'00_data/02_preprocess/2025_04/{br_name}_stdev.csv',
        f'00_data/02_preprocess/2025_05/{br_name}_stdev.csv',
        f'00_data/02_preprocess/2025_06/{br_name}_stdev.csv',
    ]

    output_file = f'03_append/{br_name}_stdev.csv'

    # 데이터프레임 초기화
    df_list = []

    # 각 파일을 읽어서 리스트에 추가
    for file_path in file_paths:
        try:
            df = pd.read_csv(file_path, na_values='Null')
            df_list.append(df)
            print(f"  {file_path} 로드 완료")
        except FileNotFoundError:
            print(f"  경고: {file_path} 파일을 찾을 수 없습니다.")

    if df_list:
        # 모든 데이터프레임을 합치기
        df_app = pd.concat(df_list, ignore_index=True)
        print(f"  결측값 확인: {df_app.isna().sum().sum()}개")
        
        # 결과를 CSV 파일로 저장
        df_app.to_csv(output_file, index=False)
        print(f"  {br_name}_stdev.csv 저장 완료")
    else:
        print(f"  경고: {br_name}의 stdev 데이터가 없습니다.")

    print(f"다리 번호 {n} ({br_name}) 처리 완료\n")

def main():
    """사용자 입력을 받아서 다리 하나씩 처리"""
    print("한강대교 데이터 처리 프로그램")
    print("="*50)
    print("다리 번호: 1(행주), 2(가양), 3(월드컵), 4(성산), 5(양화)")
    print("          6(서강), 7(원효), 8(한강), 9(성수), 10(청담)")
    print("          11(올림픽), 12(암사), 13(성탄), 14(천호), 15(영동), 16(잠수)")
    print("="*50)
    
    while True:
        try:
            # 사용자로부터 다리 번호 입력받기
            n = input("\n처리할 다리 번호를 입력하세요 (1-16, 종료하려면 'q' 입력): ").strip()
            
            # 종료 조건
            if n.lower() == 'q':
                print("프로그램을 종료합니다.")
                break
            
            # 입력값 검증
            n = int(n)
            if n < 1 or n > 16:
                print("오류: 1~16 사이의 숫자를 입력해주세요.")
                continue
            
            # 다리 처리
            process_bridge(n)
            
            # 다음 다리 처리 여부 확인
            while True:
                next_bridge = input("\n다음 다리를 처리하시겠습니까? (y/n): ").strip().lower()
                if next_bridge in ['y', 'n']:
                    break
                else:
                    print("y 또는 n을 입력해주세요.")
            
            if next_bridge == 'n':
                print("프로그램을 종료합니다.")
                break
                
        except ValueError:
            print("오류: 올바른 숫자를 입력해주세요.")
        except Exception as e:
            print(f"오류 발생: {str(e)}")
            print("다시 시도해주세요.")

if __name__ == "__main__":
    main() 

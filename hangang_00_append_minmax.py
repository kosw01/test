import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator
import pandas as pd

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



n= 3
br_name, stdev = bridge_match(n)

from scipy import stats


file_paths = [
    f'00_data/02_preprocess/2025_01/{br_name}_min.csv',
    f'00_data/02_preprocess/2025_02/{br_name}_min.csv',
    f'00_data/02_preprocess/2025_03/{br_name}_min.csv',
    f'00_data/02_preprocess/2025_04/{br_name}_min.csv',
    f'00_data/02_preprocess/2025_05/{br_name}_min.csv',
    f'00_data/02_preprocess/2025_06/{br_name}_min.csv',
    #f'00_data/02_preprocess/2025_07/{br_name}_min.csv',
    #f'00_data/02_preprocess/2025_08/{br_name}_min.csv',
    #f'00_data/02_preprocess/2025_09/{br_name}_min.csv',
    #f'00_data/02_preprocess/2025_10/{br_name}_min.csv',
    #f'00_data/02_preprocess/2025_11/{br_name}_min.csv',
    #f'00_data/02_preprocess/2025_12/{br_name}_min.csv'
]

output_file = f'03_append/{br_name}_min.csv'

# 데이터프레임 초기화
df_list = []

# 각 파일을 읽어서 리스트에 추가
for file_path in file_paths:
    df = pd.read_csv(file_path, na_values='Null')  # "Null"을 NaN으로 처리
    #df['date'] = pd.to_datetime(df['date'])
    df_list.append(df)
    print(df.info())


# 모든 데이터프레임을 합치기
df_app = pd.concat(df_list, ignore_index=True)

# 결측값이 포함된 열을 다시 확인
print(df_app.isna().sum())

# 결측값을 평균값으로 채우기 (또는 다른 방법으로 처리)
#df_app.fillna(df_app.mean(), inplace=True)

#df_app = df_app.rename(columns={'save_day': 'date'})

# 결과를 CSV 파일로 저장
df_app.to_csv(output_file, index=False)
print(f'{br_name}이 {output_file}에 저장되었습니다.')


file_paths = [
    f'00_data/02_preprocess/2025_01/{br_name}_max.csv',
    f'00_data/02_preprocess/2025_02/{br_name}_max.csv',
    f'00_data/02_preprocess/2025_03/{br_name}_max.csv',
    f'00_data/02_preprocess/2025_04/{br_name}_max.csv',
    f'00_data/02_preprocess/2025_05/{br_name}_max.csv',
    f'00_data/02_preprocess/2025_06/{br_name}_max.csv',
    #f'00_data/02_preprocess/2025_07/{br_name}_max.csv',
    #f'00_data/02_preprocess/2025_08/{br_name}_max.csv',
    #f'00_data/02_preprocess/2025_09/{br_name}_max.csv',
    #f'00_data/02_preprocess/2025_10/{br_name}_max.csv',
    #f'00_data/02_preprocess/2025_11/{br_name}_max.csv',
    #f'00_data/02_preprocess/2025_12/{br_name}_max.csv'
]

output_file = f'03_append/{br_name}_max.csv'

# 데이터프레임 초기화
df_list = []

# 각 파일을 읽어서 리스트에 추가
for file_path in file_paths:
    df = pd.read_csv(file_path, na_values='Null')  # "Null"을 NaN으로 처리
    df_list.append(df)


# 모든 데이터프레임을 합치기
df_app = pd.concat(df_list, ignore_index=True)

# 결측값이 포함된 열을 다시 확인
print(df_app.isna().sum())

# 결측값을 평균값으로 채우기 (또는 다른 방법으로 처리)
#df_app.fillna(df_app.mean(), inplace=True)

#df_app = df_app.rename(columns={'save_day': 'date'})

# 결과를 CSV 파일로 저장
df_app.to_csv(output_file, index=False)
print(f'파일이 {output_file}에 저장되었습니다.')

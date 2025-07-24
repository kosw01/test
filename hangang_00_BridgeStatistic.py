import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm, rc
from scipy import stats
import os
import numpy as np
from matplotlib.ticker import FuncFormatter
import seaborn as sns
import shutil
import matplotlib

# 폰트 크기 상수 정의
FONT_SIZE_LABEL = 14
FONT_SIZE_TICK = 12
FONT_SIZE_LEGEND = 12

# 1. matplotlib 폰트 캐시 관련 코드 수정
try:
    cache_dir = matplotlib.get_cachedir()
    if os.path.exists(cache_dir):
        for file in os.listdir(cache_dir):
            file_path = os.path.join(cache_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'캐시 파일 삭제 중 오류 발생: {e}')
except Exception as e:
    print(f'캐시 디렉토리 처리 중 오류 발생: {e}')

# 2. 폰트 매니저 재생성
fm._load_fontmanager()

# 3. 폰트 설정
font_path = r"C:\Users\Y15599\Desktop\작업\폰트\서울남산 장체M.ttf"

try:
    font_prop = fm.FontProperties(fname=font_path)
    rc('font', family=font_prop.get_name())  # 글로벌 폰트 설정
    print(f"Font loaded successfully: {font_prop.get_name()}")
except Exception as e:
    print(f"Error loading font: {e}")
    
# 0 : 없음 / 
# 1 : 재난, 안전, 사용 / 상한
# 2 : 재난, 안전, 사용 / 하한
# 3 : 재난, 안전, 사용 / 상하한
# 4 : 추세 / 상하한
# 5 : 신뢰도 / 하한 4차관리기준
# 6 : 없음 / 신축이음계
# 7 : 막대그래프 두개(평균, 최대) / 풍속계 상한 25

# 폴더명
# 01_channel_info
# 02_groupbymonthdata
# 03_append
# 04_yearly_report
# 05_monthly_report
# 06_quarterly_report
class BridgeStatistic:
    def __init__(self, br_name):
        self.br_name = br_name
        self.info = pd.read_csv(f'01_channel_info/info_update.csv', encoding='cp949')
        self.average = pd.read_csv(f'03_append/{self.br_name}_average.csv', encoding='cp949')
        self.stdev = pd.read_csv(f'03_append/{self.br_name}_stdev.csv', encoding='cp949')
        

        # date 변환을 초기화 단계에서 수행 (더 유연한 파싱 사용)
        try:
            self.average['date'] = pd.to_datetime(self.average['date'], format='mixed')
        except:
            # 첫 번째 시도가 실패하면 더 유연한 파싱 시도
            self.average['date'] = pd.to_datetime(self.average['date'], errors='coerce')
            
        try:
            self.stdev['date'] = pd.to_datetime(self.stdev['date'], format='mixed')
        except:
            # 첫 번째 시도가 실패하면 더 유연한 파싱 시도
            self.stdev['date'] = pd.to_datetime(self.stdev['date'], errors='coerce')

    def preprocess_average_data(self, stdevlist=None, wind=None, min=None):
        # 중복된 날짜 제거 및 정렬
        self.average = self.average.drop_duplicates(subset=['date']).sort_values('date')
        self.stdev = self.stdev.drop_duplicates(subset=['date']).sort_values('date')
        
        # 평균 데이터 정리
        data_average = self.average.groupby(self.average['date'].dt.to_period('M')).mean()

        if stdevlist is not None:
            data_stdev = self.stdev.groupby(self.stdev['date'].dt.to_period('M')).mean()

            for col in stdevlist:
                if col in data_stdev.columns:
                    data_average[col] = data_stdev[col]

        if wind is not None:
            data_max = self.average.groupby(self.average['date'].dt.to_period('M')).max()

            for windcol in wind:
                if windcol in data_max.columns:
                    data_average[f'{windcol} 최대'] = data_max[windcol]

        if min is not None:
            self.min = pd.read_csv(f'03_append/{self.br_name}_min.csv', encoding='cp949')
            try:
                self.min['date'] = pd.to_datetime(self.min['date'], format='mixed')
            except:
                self.min['date'] = pd.to_datetime(self.min['date'], errors='coerce')

            data_ml = self.min.groupby(self.min['date'].dt.to_period('M')).mean()

            for mincol in min:
                if mincol in data_ml.columns:
                    data_average[mincol] = data_ml[mincol] - data_average[mincol]
                    print(data_average[mincol])

        # date 열을 '월' 형식으로 변경 (인덱스 직접 처리)
        data_average_reset = data_average.reset_index(drop=True)
        data_average_reset['date'] = data_average_reset.index + 1
        data_average_reset['date'] = data_average_reset['date'].astype(str) + '월'

        # 결과 저장
        output_path = f'02_groupbymonthdata/{self.br_name}_2025.csv'
        data_average_reset.to_csv(output_path, index=False, encoding='cp949')
        print(f'{self.br_name}_2025.csv 파일이 {output_path}에 저장되었습니다.')

    def groupbymonthdata(self):
        data = pd.read_csv(f'02_groupbymonthdata/{self.br_name}_2025.csv', encoding='cp949')
        data_ly = None
        data_ly2 = None

        # 전년도 데이터 있는지 확인하기
        if os.path.exists(f'02_groupbymonthdata/{self.br_name}_2024.csv'):
            data_ly = pd.read_csv(f'02_groupbymonthdata/{self.br_name}_2024.csv', encoding='cp949')
        
        if os.path.exists(f'02_groupbymonthdata/{self.br_name}_2023.csv'):
            data_ly2 = pd.read_csv(f'02_groupbymonthdata/{self.br_name}_2023.csv', encoding='cp949')
        
        return data, data_ly, data_ly2

    def calculate_monthly_statistics(self, stdevlist=None):
        # 월별 평균, 최소, 최대 계산
        try:
            self.average['date'] = pd.to_datetime(self.average['date'], format='mixed')
        except:
            self.average['date'] = pd.to_datetime(self.average['date'], errors='coerce')
            
        monthly_stats = self.average.groupby(self.average['date'].dt.to_period('M')).agg(['min', 'mean', 'max'])
        monthly_stats['range'] = monthly_stats[monthly_stats.columns[2]] - monthly_stats[monthly_stats.columns[0]]
        monthly_stats.columns = ['_'.join(col).strip() for col in monthly_stats.columns.values]

        # 표준편차 데이터 처리
        if stdevlist:
            try:
                self.stdev['date'] = pd.to_datetime(self.stdev['date'], format='mixed')
            except:
                self.stdev['date'] = pd.to_datetime(self.stdev['date'], errors='coerce')
                
            stdev_stats = self.stdev.groupby(self.stdev['date'].dt.to_period('M')).agg(['min', 'mean', 'max'])
            stdev_stats.columns = ['_'.join(col).strip() for col in stdev_stats.columns.values]

            for col in stdevlist:
                if col in self.stdev.columns:
                    monthly_stats[f'{col}_max'] = stdev_stats[f'{col}_max']
                    monthly_stats[f'{col}_mean'] = stdev_stats[f'{col}_mean']
                    monthly_stats[f'{col}_min'] = stdev_stats[f'{col}_min']
                    monthly_stats[f'{col}_range'] = stdev_stats[f'{col}_max'] - stdev_stats[f'{col}_min']

        return monthly_stats

    def save_statistics(self, period, monthly_stats):
        # 통계 데이터 저장
        monthly_stats.to_csv(f'05_monthly_report/{period}/{self.br_name}_stat.csv', encoding='utf-8-sig')
        print(f'05_monthly_report/{period}에 {self.br_name}_stat.csv 파일이 저장되었습니다.')

    def plot_yearly_report(self):
        """limit_type에 따라 월통계 그래프를 그리는 함수"""

        # info.csv에서 현재 교량(br_name)의 limit_type에 해당하는 channel_name 필터링
        filtered_info = self.info[self.info['br_name'] == self.br_name]
        channel_names = filtered_info['channel_name'].tolist()

        if not channel_names:
            print(f"br_name에 해당하는 채널이 없습니다: {self.br_name}")
            return

        data, data_ly1, data_ly2 = self.groupbymonthdata()

        for idx, col in enumerate(channel_names):
            if col not in data.columns:
                print(f"{col} 컬럼이 올해 데이터에 없습니다. 스킵합니다.")
                continue
            
            if data_ly1 is None:
                print(f"{self.br_name}의 작년 데이터가 없습니다. 작년 데이터 없이 진행합니다.")
            elif col not in data_ly1.columns:
                print(f"{col} 컬럼이 작년 데이터에 없습니다. 스킵합니다.")
                continue
            
            if data_ly2 is None:
                print(f"{self.br_name}의 제작년 데이터가 없습니다. 제작년 데이터 없이 진행합니다.")
            elif col not in data_ly2.columns:
                print(f"{col} 컬럼이 제작년 데이터에 없습니다. 스킵합니다.")
                continue

            # limit_type에 따른 상한선 또는 하한선 가져오기
            channel_info = filtered_info[filtered_info['channel_name'] == col].iloc[0]
            limit_type = channel_info['limit_type']
            b = channel_info.get('b')
            h = channel_info.get('h')
            up_limit = channel_info.get('up_limit', None)
            low_limit = channel_info.get('low_limit', None)
            limit1 = channel_info.get('limit1', None)
            limit2 = channel_info.get('limit2', None)
            limit3 = channel_info.get('limit3', None)
            limit4 = channel_info.get('limit4', None)
            significant_figure = channel_info.get('significant_figure', None)
            correl = channel_info.get('correl',None)
            legendloc = channel_info.get('legendloc', None)
            bbox_x = channel_info.get('bbox_x', None)
            bbox_y = channel_info.get('bbox_y', None)
            eqk1 = channel_info.get('eqk1', None)
            eqk2 = channel_info.get('eqk2', None)

            # ylabel 구성: 채널종류 + (단위)
            ylabel = f"{channel_info['채널종류']} ({channel_info['unit']})"
            
            def format_y_ticks(value, _):
                if value < 0:
                    return f'-{abs(value):,.{significant_figure}f}'
                else:
                    return f'{value:,.{significant_figure}f}'

            # 그래프 생성
            fig, ax = plt.subplots(figsize=(b, h))

            if limit_type in [0,6,12,13]:  # limit_type 0: 막대그래프 3개년
                if data_ly2 is not None:
                    ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                if data_ly1 is not None:
                    ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')
                ax.bar(data.index + 0.25, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')

            elif limit_type == 1:  # limit_type 1: 상한선이 있는 막대그래프 3개년
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Limit: {round(up_limit,significant_figure)}')
                if data_ly2 is not None:
                    ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                if data_ly1 is not None:
                    ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')
                ax.bar(data.index + 0.25, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')

            elif limit_type == 2:  # limit_type 2: 하한선이 있는 막대그래프 3개년
                if data_ly2 is not None:
                    ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                if data_ly1 is not None:
                    ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')
                ax.bar(data.index + 0.25, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')
                ax.axhline(y=low_limit, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {round(low_limit,significant_figure)}')

            elif limit_type == 3:  # limit_type 3: 꺾은선그래프 3개년 (상하한선 포함)
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Upper Limit: {round(up_limit,significant_figure)}')
                if data_ly2 is not None:
                    ax.plot(data_ly2.index, data_ly2[col], label='2023년', color='dimgray', marker='o')
                if data_ly1 is not None:
                    ax.plot(data_ly1.index, data_ly1[col], label='2024년', color='forestgreen', marker='o')
                ax.plot(data.index, data[col], label='2025년', color='darkorange', marker='o')
                ax.axhline(y=low_limit, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {round(low_limit,significant_figure)}')

            elif limit_type == 4:  # limit_type 4: 꺾은선그래프 3개년 (상하한선 포함)
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Upper Limit: {round(up_limit,significant_figure)}')
                if data_ly2 is not None:
                    ax.plot(data_ly2.index, data_ly2[col], label='2023년', color='dimgray', marker='o')
                if data_ly1 is not None:
                    ax.plot(data_ly1.index, data_ly1[col], label='2024년', color='forestgreen', marker='o')
                ax.plot(data.index, data[col], label='2025년', color='darkorange', marker='o')
                ax.axhline(y=low_limit, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {round(low_limit,significant_figure)}')

            elif limit_type == 5:  # limit_type 5: 막대그래프 3개년 (하한 4차관리기준)
                if data_ly2 is not None:
                    ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                if data_ly1 is not None:
                    ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')                
                ax.bar(data.index + 0.25, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')
                ax.axhline(y=limit1, color='lightsteelblue', linestyle='--', linewidth=2, label=f'Lower Limit: {round(limit1,significant_figure)}')
                ax.axhline(y=limit2, color='cornflowerblue', linestyle='--', linewidth=2, label=f'Lower Limit: {round(limit2,significant_figure)}')
                ax.axhline(y=limit3, color='royalblue', linestyle='--', linewidth=2, label=f'Lower Limit: {round(limit3,significant_figure)}')
                ax.axhline(y=limit4, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {round(limit4,significant_figure)}')

            # elif limit_type == 6:  # limit_type 6: 신축이음계
            #     ax.scatter(data[col], data[correl])
            #     ax.set_xlabel('온도', fontsize=FONT_SIZE_LABEL, fontweight='bold', font=font_prop)

            elif limit_type == 7:  # limit_type 7: 풍속
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Limit: {round(up_limit,significant_figure)}')
                # if data_ly2 is not None:
                #     ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                # if data_ly1 is not None:
                #     ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')
                ax.bar(data.index - 0.125, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')
                ax.bar(data.index + 0.125, data[f'{col} 최대'], width=0.25, label=f'{col} 최대', align='center', alpha=0.8, color='blue')

            elif limit_type == 8:  # limit_type 8 지진 3개
                ax.bar(data.index-0.25, data[col], width=0.25, label=col, align='center', alpha=0.8, color='forestgreen')
                ax.bar(data.index, data[eqk1], width=0.25, label=eqk1, align='center', alpha=0.8, color='darkorange')
                ax.bar(data.index+0.25, data[eqk2], width=0.25, label=eqk2, align='center', alpha=0.8, color='dimgray')

            elif limit_type == 9:  # limit_type 9 지진 2개
                ax.bar(data.index-0.125, data[col], width=0.25, label=col, align='center', alpha=0.8, color='forestgreen')
                ax.bar(data.index+0.125, data[eqk1], width=0.25, label=eqk1, align='center', alpha=0.8, color='dimgray')

            elif limit_type == 10:  # limit_type 10: 막대그래프, 하한선은 없고 label만 표시
                if data_ly2 is not None:
                    ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                if data_ly1 is not None:
                    ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')
                ax.bar(data.index + 0.25, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')
                ax.plot([], [], color='blue', linestyle='-', linewidth=2, label=f'Limit: {round(low_limit,significant_figure)}')

            elif limit_type == 11:  # limit_type 11: 막대그래프, 상한선은 없고 label만 표시
                if data_ly2 is not None:
                    ax.bar(data_ly2.index - 0.25, data_ly2[col], width=0.25, label='2023년', color='dimgray')
                if data_ly1 is not None:
                    ax.bar(data_ly1.index, data_ly1[col], width=0.25, label='2024년', color='forestgreen')
                ax.bar(data.index + 0.25, data[col], width=0.25, label='2025년', align='center', alpha=0.8, color='darkorange')
                ax.plot([], [], color='red', linestyle='-', linewidth=2, label=f'Limit: {round(up_limit,significant_figure)}')

            # elif limit_type == 12:  # limit_type 12: 산점도
            #     ax.scatter(data[col], data[correl])
            #     ax.set_xlabel('온도', fontsize=FONT_SIZE_LABEL, fontweight='bold', font=font_prop)

            ax.set_ylabel(ylabel, fontsize=FONT_SIZE_LABEL, fontweight='bold', font=font_prop)
            ax.legend(loc=legendloc, bbox_to_anchor=(bbox_x, bbox_y), prop=font_prop, fontsize=FONT_SIZE_LEGEND, ncol=5)
            ax.set_xticks(data.index)
            ax.set_xticklabels([f'{i + 1}월' for i in data.index], 
                             fontsize=FONT_SIZE_TICK, 
                             fontweight='bold', 
                             font=font_prop)
            ax.tick_params(axis='y', labelsize=FONT_SIZE_TICK)
            ax.yaxis.set_major_formatter(FuncFormatter(format_y_ticks))
            ax.grid(True)
            fig.tight_layout()

            # 결과 저장
            output_dir = f'04_yearly_report/{self.br_name}'
            os.makedirs(output_dir, exist_ok=True)
            plt.savefig(f'{output_dir}/{col}.jpg')
            plt.close(fig)
            print(f'{idx + 1}/{len(channel_names)}: {col} 그래프 생성 완료')

    def plot_quarterly_report(self, Q):
        """limit_type에 따라 월통계 그래프를 그리는 함수"""

        # info.csv에서 현재 교량(br_name)의 limit_type에 해당하는 channel_name 필터링
        filtered_info = self.info[self.info['br_name'] == self.br_name]
        channel_names = filtered_info['channel_name'].tolist()

        if not channel_names:
            print(f"br_name에 해당하는 채널이 없습니다: {self.br_name}")
            return

        data, data_ly1, data_ly2 = self.groupbymonthdata()

        for idx, col in enumerate(channel_names):
            if col not in data.columns:
                print(f"{col} 컬럼이 올해 데이터에 없습니다. 스킵합니다.")
                continue

            # limit_type에 따른 상한선 또는 하한선 가져오기
            channel_info = filtered_info[filtered_info['channel_name'] == col].iloc[0]
            limit_type = channel_info['limit_type']
            b = channel_info.get('b')
            h = channel_info.get('h')
            up_limit = channel_info.get('up_limit', None)
            low_limit = channel_info.get('low_limit', None)
            limit1 = channel_info.get('limit1', None)
            limit2 = channel_info.get('limit2', None)
            limit3 = channel_info.get('limit3', None)
            limit4 = channel_info.get('limit4', None)
            significant_figure = channel_info.get('significant_figure', None)
            correl = channel_info.get('correl',None)
            legendloc = channel_info.get('legendloc', None)
            bbox_x = channel_info.get('bbox_x', None)
            bbox_y = channel_info.get('bbox_y', None)
            eqk1 = channel_info.get('eqk1', None)
            eqk2 = channel_info.get('eqk2', None)


            # ylabel 구성: 채널종류 + (단위)
            ylabel = f"{channel_info['채널종류']} ({channel_info['unit']})"
            
            def format_y_ticks(value, _):
                if value < 0: 
                    return f'-{abs(value):,.{significant_figure}f}'
                else:
                    return f'{value:,.{significant_figure}f}'

            # 그래프 생성
            fig, ax = plt.subplots(figsize=(b, h))

            if limit_type == 0:  # limit_type 0 막대그래프 3개년
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')

            elif limit_type == 1:  # limit_type 1: 상한선이 있는 막대그래프 3개년
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Limit: {round(up_limit,significant_figure)}')
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')

            elif limit_type == 2:  # limit_type 2: 하한선이 있는 막대그래프 3개년
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')
                ax.axhline(y=low_limit, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {low_limit}')

            elif limit_type == 3:  # limit_type 3: 꺾은선그래프 3개년 (상하한선 포함)
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Upper Limit: {up_limit}')
                ax.plot(data.index, data[col], label=col, color='darkorange', marker='o')
                ax.axhline(y=low_limit, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {low_limit}')

            elif limit_type == 4:  # limit_type 4: 꺾은선그래프 3개년 (상하한선 포함)
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Upper Limit: {up_limit}')
                ax.plot(data.index, data[col], label=col, color='darkorange', marker='o')
                ax.axhline(y=low_limit, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {low_limit}')

            elif limit_type == 5:  # limit_type 5: 막대그래프 3개년 (하한 4차관리기준)
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')
                ax.axhline(y=limit1, color='lightsteelblue', linestyle='--', linewidth=2, label=f'Lower Limit1: {round(limit1,significant_figure)}')
                ax.axhline(y=limit2, color='cornflowerblue', linestyle='--', linewidth=2, label=f'Lower Limit2: {round(limit2,significant_figure)}')
                ax.axhline(y=limit3, color='royalblue', linestyle='--', linewidth=2, label=f'Lower Limit3: {round(limit3,significant_figure)}')
                #ax.axhline(y=limit4, color='blue', linestyle='-', linewidth=2, label=f'Lower Limit: {limit4}')
                ax.plot([], [], color='blue', linestyle='-', linewidth=2, label=f'Lower Limit4: {round(limit4,significant_figure)}')

            elif limit_type in [6, 12, 13, 14]:  # limit_type 6: 신축그래프 
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')

            elif limit_type == 7:  # limit_type 7: 풍속
                ax.axhline(y=up_limit, color='red', linestyle='-', linewidth=2, label=f'Limit: {round(up_limit,significant_figure)}')
                ax.bar(data.index-0.125, data[col], width=0.25, label=col, align='center', alpha=0.8, color='blue')
                ax.bar(data.index+0.125, data[f'{col} 최대'], width=0.25, label=f'{col} 최대', align='center', alpha=0.8, color='darkorange')
            
            elif limit_type == 8:  # limit_type 8 지진 3개
                ax.bar(data.index-0.25, data[col], width=0.25, label=col, align='center', alpha=0.8, color='forestgreen')
                ax.bar(data.index, data[eqk1], width=0.25, label=eqk1, align='center', alpha=0.8, color='darkorange')
                ax.bar(data.index+0.25, data[eqk2], width=0.25, label=eqk2, align='center', alpha=0.8, color='dimgray')

            elif limit_type == 9:  # limit_type 9 지진 2개
                ax.bar(data.index-0.125, data[col], width=0.25, label=col, align='center', alpha=0.8, color='forestgreen')
                ax.bar(data.index+0.125, data[eqk1], width=0.25, label=eqk1, align='center', alpha=0.8, color='dimgray')

            elif limit_type == 10:  # limit_type 10: 막대그래프, 하한선은 없고 label만 표시
                # 막대 그래프 그리기
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')
                
                # 빈 라인을 추가하여 범례에만 상한값 표시 (실제 선은 그리지 않음)
                ax.plot([], [], color='blue', linestyle='-', linewidth=2, label=f'Limit: {round(low_limit,significant_figure)}')

            elif limit_type == 11:  # limit_type 11: 막대그래프, 상한선은 없고 label만 표시
                # 막대 그래프 그리기
                ax.bar(data.index, data[col], width=0.25, label=col, align='center', alpha=0.8, color='darkorange')
                
                # 빈 라인을 추가하여 범례에만 상한값 표시 (실제 선은 그리지 않음)
                ax.plot([], [], color='red', linestyle='-', linewidth=2, label=f'Limit: {round(up_limit,significant_figure)}')


            else:
                pass

            ax.set_ylabel(ylabel, fontsize=FONT_SIZE_LABEL, fontweight='bold', font=font_prop)
            ax.legend(loc=legendloc, bbox_to_anchor=(bbox_x, bbox_y), prop=font_prop, fontsize=FONT_SIZE_LEGEND, ncol=5)
            ax.set_xticks(data.index)
            ax.set_xticklabels([f'{i + 1}월' for i in data.index], 
                            fontsize=FONT_SIZE_TICK, 
                            fontweight='bold', 
                            font=font_prop)
            ax.tick_params(axis='y', labelsize=FONT_SIZE_TICK)
            ax.tick_params(axis='x', labelsize=FONT_SIZE_TICK)
            ax.yaxis.set_major_formatter(FuncFormatter(format_y_ticks))
            ax.grid(True)
            fig.tight_layout()

            # 결과 저장
            output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
            os.makedirs(output_dir, exist_ok=True)
            plt.savefig(f'{output_dir}/{col}.jpg')
            plt.close(fig)
            print(f'{idx + 1}/{len(channel_names)}: {col} 그래프 생성 완료')

    def plot_scatter(self, Q):
        
        # info.csv에서 현재 교량(br_name)의 limit_type에 해당하는 channel_name 필터링
        filtered_info = self.info[(self.info['br_name'] == self.br_name)|(self.info['limit_type'].isin([6, 12, 13, 14]))]
        filtered_info1 = self.info[(self.info['br_name'] == self.br_name)&(self.info['limit_type'].isin([6, 12, 13, 14]))]
        channel_names = filtered_info1['channel_name'].tolist()
        print(channel_names)

        if not channel_names:
            print(f"br_name에 해당하는 채널이 없습니다: {self.br_name}")
            return

        data, data_ly1, data_ly2 = self.groupbymonthdata()

        for idx, col in enumerate(channel_names):
            if col not in data.columns:
                print(f"{col} 컬럼이 올해 데이터에 없습니다. 스킵합니다.")
                continue

            # limit_type에 따른 상관관계정보 가져오기
            channel_info = filtered_info[filtered_info['channel_name'] == col].iloc[0]
            channel_info1 = filtered_info1[filtered_info1['channel_name'] == col].iloc[0]
            limit_type = channel_info1['limit_type']
            b = channel_info1.get('b')
            h = channel_info1.get('h')
            correl = channel_info1['correl']

            # ylabel 구성: 채널종류 + (단위)
            ylabel = f"{channel_info1['채널종류']} ({channel_info1['unit']})"

            def format_y_ticks(value, _):
                if value < 0:
                    return f'-{abs(value):,.0f}'
                return f'{value:,.0f}'
            
            print(idx, col, correl)
            
            if limit_type == 12:  # limit_type 12: 산점도
                # matplotlib 스타일 설정
                plt.style.use('default')
                matplotlib.rcParams['axes.unicode_minus'] = False  # 음수 기호 표시 설정
                
                fig, ax = plt.subplots(figsize=(b/2, h))

                # NaN 제거
                valid_data = self.average[[col, correl]].dropna()

                if valid_data.empty:
                    print(f"데이터가 부족하여 {col} 산점도를 생성할 수 없습니다.")
                    continue

                # x_data 기준으로 정렬
                valid_data = valid_data.sort_values(by=correl)
                x_data = valid_data[correl]
                y_data = valid_data[col]

                # 20을 기준으로 데이터 분리
                below_20 = valid_data[valid_data[correl] <= 20]
                above_20 = valid_data[valid_data[correl] > 20]

                # 20°C 이하 데이터 그래프
                if not below_20.empty:
                    fig_below, ax_below = plt.subplots(figsize=(b/2, h))
                    ax_below.scatter(below_20[correl], below_20[col])
                    
                    # 회귀 분석 수행
                    slope_below, intercept_below, r_value_below, p_value_below, std_err_below = stats.linregress(below_20[correl], below_20[col])
                    ax_below.plot(below_20[correl], intercept_below + slope_below * below_20[correl], color='red', 
                            label=f'y = {slope_below:.2f}x + {intercept_below:.2f}\n$R^2$ = {r_value_below**2:.2f}')

                    # y축 포맷터 설정
                    def format_y_ticks(value, _):
                        if value < 0:
                            return f'-{abs(value):,.0f}'
                        return f'{value:,.0f}'
                    
                    # y축, x축 포맷터 적용
                    formatter = FuncFormatter(format_y_ticks)
                    ax_below.yaxis.set_major_formatter(formatter)
                    ax_below.xaxis.set_major_formatter(formatter)

                    # y축 범위 설정
                    ax_below.set_ylim(-5, 55)       

                    # y축, x축 눈금 설정
                    ax_below.tick_params(axis='y', which='major', labelsize=12)
                    ax_below.tick_params(axis='x', which='major', labelsize=12)
                    
                    ax_below.set_xlabel(f'{channel_info1["correl"]} (ºC)', fontsize=14, fontweight='bold', font=font_prop)
                    ax_below.set_ylabel(ylabel, fontsize=14, fontweight='bold', font=font_prop)
                    ax_below.legend(loc='best', prop=font_prop)
                    ax_below.grid(True)
                    ax_below.invert_xaxis()    
                    fig_below.tight_layout()

                    # 결과 저장
                    output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
                    os.makedirs(output_dir, exist_ok=True)
                    plt.savefig(f'{output_dir}/{col}_below.jpg')
                    plt.close(fig_below)
                    print(f'{correl} vs {col} 그래프 생성 완료')

                # 20°C 초과 데이터 그래프
                if not above_20.empty:
                    fig_above, ax_above = plt.subplots(figsize=(b/2, h))
                    ax_above.scatter(above_20[correl], above_20[col])
                    
                    # 회귀 분석 수행
                    slope_above, intercept_above, r_value_above, p_value_above, std_err_above = stats.linregress(above_20[correl], above_20[col])
                    ax_above.plot(above_20[correl], intercept_above + slope_above * above_20[correl], color='red', 
                            label=f'y = {slope_above:.2f}x + {intercept_above:.2f}\n$R^2$ = {r_value_above**2:.2f}')

                    # y축 포맷터 설정
                    def format_y_ticks(value, _):
                        if value < 0:
                            return f'-{abs(value):,.0f}'
                        return f'{value:,.0f}'
                    
                    # y축, x축 포맷터 적용
                    formatter = FuncFormatter(format_y_ticks)
                    ax_above.yaxis.set_major_formatter(formatter)
                    ax_above.xaxis.set_major_formatter(formatter)

                    # y축 범위 설정
                    ax_above.set_ylim(-5, 55)
                    ax_above.set_xlim(20, 50)
                    
                    # y축 눈금 설정
                    ax_above.tick_params(axis='y', which='major', labelsize=12)
                    ax_above.tick_params(axis='x', which='major', labelsize=12)
                    
                    ax_above.set_xlabel(f'{channel_info1["correl"]} (ºC)', fontsize=14, fontweight='bold', font=font_prop)
                    ax_above.set_ylabel(ylabel, fontsize=14, fontweight='bold', font=font_prop)
                    ax_above.legend(loc='best', prop=font_prop)
                    ax_above.grid(True)
                    ax_above.invert_xaxis() 
                    fig_above.tight_layout()

                    # 결과 저장
                    output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
                    os.makedirs(output_dir, exist_ok=True)
                    plt.savefig(f'{output_dir}/{col}_above.jpg')
                    plt.close(fig_above)

                print(f'{idx + 1}/{len(channel_names)}: {col} 그래프 생성 완료')

            elif limit_type == 13:  # limit_type 13: 18도를 기준으로 데이터를 나누는 산점도
                # matplotlib 스타일 설정
                plt.style.use('default')
                matplotlib.rcParams['axes.unicode_minus'] = False  # 음수 기호 표시 설정
                
                fig, ax = plt.subplots(figsize=(b/2, h))

                # NaN 제거
                valid_data = self.average[[col, correl]].dropna()

                if valid_data.empty:
                    print(f"데이터가 부족하여 {col} 산점도를 생성할 수 없습니다.")
                    continue

                # x_data 기준으로 정렬
                valid_data = valid_data.sort_values(by=correl)
                x_data = valid_data[correl]
                y_data = valid_data[col]

                # 18을 기준으로 데이터 분리
                below_18 = valid_data[valid_data[correl] <= 18]
                above_18 = valid_data[valid_data[correl] > 18]

                # 18°C 이하 데이터 그래프
                if not below_18.empty:
                    fig_below, ax_below = plt.subplots(figsize=(b/2, h))
                    ax_below.scatter(below_18[correl], below_18[col])
                    
                    # 회귀 분석 수행
                    slope_below, intercept_below, r_value_below, p_value_below, std_err_below = stats.linregress(below_18[correl], below_18[col])
                    ax_below.plot(below_18[correl], intercept_below + slope_below * below_18[correl], color='red', 
                            label=f'y = {slope_below:.2f}x + {intercept_below:.2f}\n$R^2$ = {r_value_below**2:.2f}')

                    # y축 포맷터 설정
                    def format_y_ticks(value, _):
                        if value < 0:
                            return f'-{abs(value):,.0f}'
                        return f'{value:,.0f}'
                    
                    # y축, x축 포맷터 적용
                    formatter = FuncFormatter(format_y_ticks)
                    ax_below.yaxis.set_major_formatter(formatter)
                    ax_below.xaxis.set_major_formatter(formatter)
                    
                    # y축 범위 설정
                    ax_below.set_ylim(-5,20)

                    # y축, x축 눈금 설정
                    ax_below.tick_params(axis='y', which='major', labelsize=12)
                    ax_below.tick_params(axis='x', which='major', labelsize=12)
                    
                    ax_below.set_xlabel(f'{channel_info1["correl"]} (ºC)', fontsize=14, fontweight='bold', font=font_prop)
                    ax_below.set_ylabel(ylabel, fontsize=14, fontweight='bold', font=font_prop)
                    ax_below.legend(loc='best', prop=font_prop)
                    ax_below.grid(True)
                    ax_below.invert_xaxis()
                    fig_below.tight_layout()

                    # 결과 저장
                    output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
                    os.makedirs(output_dir, exist_ok=True)
                    plt.savefig(f'{output_dir}/{col}_below.jpg')
                    plt.close(fig_below)

                # 18°C 초과 데이터 그래프
                if not above_18.empty:
                    fig_above, ax_above = plt.subplots(figsize=(b/2, h))
                    ax_above.scatter(above_18[correl], above_18[col])
                    
                    # 회귀 분석 수행
                    slope_above, intercept_above, r_value_above, p_value_above, std_err_above = stats.linregress(above_18[correl], above_18[col])
                    ax_above.plot(above_18[correl], intercept_above + slope_above * above_18[correl], color='red', 
                            label=f'y = {slope_above:.2f}x + {intercept_above:.2f}\n$R^2$ = {r_value_above**2:.2f}')

                    # y축 포맷터 설정
                    def format_y_ticks(value, _):
                        if value < 0:
                            return f'-{abs(value):,.0f}'
                        return f'{value:,.0f}'
                    
                    # y축, x축 포맷터 적용
                    formatter = FuncFormatter(format_y_ticks)
                    ax_above.yaxis.set_major_formatter(formatter)
                    ax_above.xaxis.set_major_formatter(formatter)
                    
                    # y축 범위 설정
                    ax_above.set_ylim(-5, 20)
                    ax_above.set_xlim(15, 40)

                    # y축, x축 눈금 설정
                    ax_above.tick_params(axis='y', which='major', labelsize=12)
                    ax_above.tick_params(axis='x', which='major', labelsize=12)
                    
                    ax_above.set_xlabel(f'{channel_info1["correl"]} (ºC)', fontsize=14, fontweight='bold', font=font_prop)
                    ax_above.set_ylabel(ylabel, fontsize=14, fontweight='bold', font=font_prop)
                    ax_above.legend(loc='best', prop=font_prop)
                    ax_above.grid(True)
                    ax_above.invert_xaxis()
                    fig_above.tight_layout()

                    # 결과 저장
                    output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
                    os.makedirs(output_dir, exist_ok=True)
                    plt.savefig(f'{output_dir}/{col}_above.jpg')
                    plt.close(fig_above)

                print(f'{idx + 1}/{len(channel_names)}: {col} 그래프 생성 완료')

            elif limit_type == 6:  # limit_type 6: 산점도
                # matplotlib 스타일 설정
                plt.style.use('default')
                matplotlib.rcParams['axes.unicode_minus'] = False  # 음수 기호 표시 설정
                
                fig, ax = plt.subplots(figsize=(b, h))

                # NaN 제거
                valid_data = self.average[[col, correl]].dropna()

                if valid_data.empty:
                    print(f"데이터가 부족하여 {col} 산점도를 생성할 수 없습니다.")
                    continue

                x_data = valid_data[correl]
                y_data = valid_data[col]

                ax.scatter(x_data, y_data)

                # 회귀 분석 수행
                slope, intercept, r_value, p_value, std_err = stats.linregress(x_data, y_data)

                # 회귀선 추가
                ax.plot(x_data, intercept + slope * x_data, color='red', 
                        label=f'y = {slope:.2f}x + {intercept:.2f}\n$R^2$ = {r_value**2:.2f}')

                # y축 포맷터 설정
                def format_y_ticks(value, _):
                    if value < 0:
                        return f'-{abs(value):,.0f}'
                    return f'{value:,.0f}'
                
                # y축, x축 포맷터 적용
                formatter = FuncFormatter(format_y_ticks)
                ax.yaxis.set_major_formatter(formatter)
                ax.xaxis.set_major_formatter(formatter)
                
                # y축 범위 설정
                y_min = min(y_data.min(), 0)  # 음수 값이 있다면 포함
                y_max = max(y_data.max(), 0)  # 양수 값이 있다면 포함
                ax.set_ylim(y_min - 0.1 * abs(y_min), y_max + 0.1 * abs(y_max))

                # y축, x축 눈금 설정
                ax.tick_params(axis='y', which='major', labelsize=12)
                ax.tick_params(axis='x', which='major', labelsize=12)
                
                ax.set_xlabel(f'{channel_info1["correl"]} (ºC)', fontsize=14, fontweight='bold', font=font_prop)
                ax.set_ylabel(ylabel, fontsize=14, fontweight='bold', font=font_prop)
                ax.legend(loc='best', prop=font_prop)
                ax.grid(True)
                fig.tight_layout()

                # 결과 저장
                output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
                os.makedirs(output_dir, exist_ok=True)
                plt.savefig(f'{output_dir}/{col}_.jpg')
                plt.close(fig)
                print(f'{idx + 1}/{len(channel_names)}: {col} 그래프 생성 완료')

            elif limit_type == 14:  # limit_type 14: 산점도
                # matplotlib 스타일 설정
                plt.style.use('default')
                matplotlib.rcParams['axes.unicode_minus'] = False  # 음수 기호 표시 설정
                
                fig, ax = plt.subplots(figsize=(b, h))

                # NaN 제거
                valid_data = self.average[[col, correl]].dropna()

                if valid_data.empty:
                    print(f"데이터가 부족하여 {col} 산점도를 생성할 수 없습니다.")
                    continue

                x_data = valid_data[correl]
                y_data = valid_data[col]

                ax.scatter(x_data, y_data)

                # 회귀 분석 수행
                slope, intercept, r_value, p_value, std_err = stats.linregress(x_data, y_data)

                # 회귀선 추가
                ax.plot(x_data, intercept + slope * x_data, color='red', 
                        label=f'y = {slope:.2f}x + {intercept:.2f}\n$R^2$ = {r_value**2:.2f}')

                # y축 포맷터 설정
                def format_y_ticks(value, _):
                    if value < 0:
                        return f'-{abs(value):,.0f}'
                    return f'{value:,.0f}'
                
                # y축, x축 포맷터 적용
                formatter = FuncFormatter(format_y_ticks)
                ax.yaxis.set_major_formatter(formatter)
                ax.xaxis.set_major_formatter(formatter)
                
                # y축 범위 설정
                y_min = min(y_data.min(), 0)  # 음수 값이 있다면 포함
                y_max = max(y_data.max(), 0)  # 양수 값이 있다면 포함
                ax.set_ylim(y_min - 0.1 * abs(y_min), y_max + 0.1 * abs(y_max))

                # y축, x축 눈금 설정
                ax.tick_params(axis='y', which='major', labelsize=12)
                ax.tick_params(axis='x', which='major', labelsize=12)
                
                ax.set_xlabel(f'{channel_info1["correl"]} (ºC)', fontsize=14, fontweight='bold', font=font_prop)
                ax.set_ylabel(ylabel, fontsize=14, fontweight='bold', font=font_prop)
                ax.legend(loc='best', prop=font_prop)
                ax.grid(True)
                ax.invert_xaxis()  # x축 반전
                fig.tight_layout()

                # 결과 저장
                output_dir = f'06_quarterly_report/{Q}/{self.br_name}'
                os.makedirs(output_dir, exist_ok=True)
                plt.savefig(f'{output_dir}/{col}_.jpg')
                plt.close(fig)
                print(f'{idx + 1}/{len(channel_names)}: {col} 그래프 생성 완료')
            else:
                pass

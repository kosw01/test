import pandas as pd
import os
import win32com.client as win32

# 엑셀 파일에서 데이터 불러오기
def load_excel_data(excel_file):
    # 각 시트를 데이터프레임으로 불러오기
    reception_df = pd.read_excel(excel_file, sheet_name='데이터 수신율')
    outlier_df = pd.read_excel(excel_file, sheet_name='이상치 분석')
    limit_df = pd.read_excel(excel_file, sheet_name='관리기준 초과 여부')
    
    return {
        'reception': reception_df,
        'outlier': outlier_df,
        'limit': limit_df
    }

# 한글문서 텍스트입력 
def insert_text(text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

# i번째 교량 문서 작성 중; 점검내용 중 원인 작성 칸으로 이동
def move_to_start_position(hwp, i):
    hwp.MoveToField(f'carrot{{{{{i}}}}}',True,False,False)
    for _ in range(3):
        hwp.HAction.Run("TableRightCell")

# 다음 행으로 이동
def move_to_next_row(hwp, count=8):
    for _ in range(count):
        hwp.HAction.Run("TableRightCell")


# 한글파일 열기
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True # 아래한글 첫번째 파일 보이게해줘

# 파일경로 꼭 확인하기(절대경로로 작성)
hwp.Open(r"C:\Users\Y15599\Desktop\작업\통계데이터(25년 01~04월)\월간모니터링보고서(5월).hwp")

# 참조값 설정
field = pd.read_csv('02_hwp_ref/hwp_ref.csv', encoding='cp949')              # 월간모니터링 결과 총괄 필드명
input = pd.read_csv('02_hwp_ref/hwp_input.csv', encoding='cp949')            # 월간모니터링 결과 총괄 입력값
info = pd.read_csv('01_channel_info/channel_info.csv', encoding='cp949')     # 채널정보
pic_path = r'C:\Users\Y15599\Desktop\작업\통계데이터(25년 01~04월)'               # 가동율 그래프를 담고있는 폴더의 절대경로


# 월간모니터링 결과 총괄 + 계측시설 성능지수 + 교량별 정상이상 수량 집계
for i in range(len(input)-1):
    for j in range(len(input.columns)-1):
        hwp.PutFieldText(field.iloc[i+1,j+1],input.iloc[i+1,j+1])

# 교량별 센서 현황을 점검 내용에 작성하기(교량 수량만큼 루프돌기)
for i in range(2):
    filtered_info = info[info['no.'] == i+1]
    channel_names = filtered_info['channel_name'].tolist()
    br_name = filtered_info['br_name'].iloc[0]
    
    # 엑셀 파일 경로
    excel_file = f'{br_name}/{br_name}_요약보고서.xlsx'
    
    # 엑셀 데이터 불러오기
    if os.path.exists(excel_file):
        excel_data = load_excel_data(excel_file)
        bridge_report = excel_data['reception']  # 데이터 수신율 시트 사용
        limit_report = excel_data['limit']       # 관리기준 초과 시트 사용
    else:
        print(f"엑셀 요약 보고서가 존재하지 않습니다: {excel_file}")
        continue
    
    hwp.MoveToField(f'가동율{{{{{i+1}}}}}',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}\{br_name}_주별가동율.png'), Embedded=False, sizeoption=2)

    # # 테이블 생성 (지정한 캐럿으로 이동해서 행추가를 (교량개수-1)회 수행
    # print(f'carrot{{{{{i}}}}}')
    # for j in range(len(channel_names)-1):
    #     hwp.MoveToField(f'carrot{{{{{i}}}}}',True,False,False)
    #     hwp.HAction.Run("TableCellBlock")
    #     hwp.HAction.Run("TableAppendRow")

    # # 채널구분 및 채널명 작성
    # hwp.MoveToField(f'carrot{{{{{i}}}}}',True,False,False)
    # for j in range(len(channel_names)):
    #     channel_type = filtered_info['센서종류'].iloc[j]
    #     insert_text(channel_type)
    #     hwp.HAction.Run("TableRightCell")
    #     insert_text(j+1)
    #     hwp.HAction.Run("TableRightCell")
    #     insert_text(channel_names[j])
    #     hwp.HAction.Run("TableRightCell")
    #     hwp.HAction.Run("TableRightCell")
    #     hwp.HAction.Run("TableRightCell")
    #     hwp.HAction.Run("TableRightCell")
    #     hwp.HAction.Run("TableRightCell")
    #     hwp.HAction.Run("TableRightCell")

    # # 데이터 수신율 확인
    # move_to_start_position(hwp, i)
    # for j in range(len(channel_names)):
    #     reception_rate = float(bridge_report['Unnamed: 3'].iloc[j+1])
    #     if reception_rate <= 20:
    #         insert_text(f'데이터 수신불량 수신율:{reception_rate}%')
    #         move_to_next_row(hwp, 2)
    #         insert_text('O')
    #         print(f'{br_name} 채널명 {channel_names[j]}에서 데이터수신율 {reception_rate}%')
    #         move_to_next_row(hwp, 6)   
    #     elif reception_rate <= 50:
    #         insert_text('일부구간 결측')
    #         move_to_next_row(hwp, 1)
    #         insert_text('O')
    #         print(f'{br_name} 채널명 {channel_names[j]}에서 데이터수신율 {reception_rate}%')
    #         move_to_next_row(hwp, 7)    
    #     elif reception_rate <= 80:
    #         insert_text('일부구간 결측')
    #         move_to_next_row(hwp, 3)
    #         insert_text('O')
    #         print(f'{br_name} 채널명 {channel_names[j]}에서 데이터수신율 {reception_rate}%')
    #         move_to_next_row(hwp, 5)  
    #     else:
    #         move_to_next_row(hwp)

    # # 관리기준 초과 확인
    # move_to_start_position(hwp, i)
    # for j in range(len(channel_names)):
    #     try:
    #         if limit_report['Unnamed: 5'].iloc[j+1] == '초과':
    #             insert_text('관리기준 초과\r\n')
    #         move_to_next_row(hwp)
    #     except:
    #         print(f'{br_name} 채널명 {channel_names[j]}에서 관리기준 초과 확인 중 오류 발생')

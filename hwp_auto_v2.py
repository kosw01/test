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

def insert_rate_operate(hwp, br_name):
    hwp.MoveToField(f'가동율',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/{br_name}_주별가동율.png'), Embedded=False, sizeoption=2)    
def insert_daq_pic1(hwp, br_name):
    hwp.MoveToField(f'운영프로그램1',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/기타프로그램.jpg'), Embedded=False, sizeoption=3)  
def insert_daq_pic2(hwp, br_name):
    hwp.MoveToField(f'운영프로그램2',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/URMS.jpg'), Embedded=False, sizeoption=3)  
def insert_v3_pic(hwp, br_name):
    hwp.MoveToField(f'백신',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/바이러스.jpg'), Embedded=False, sizeoption=3)  
def insert_eqk_pic(hwp, br_name):
    hwp.MoveToField(f'지진점검',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/행정안전부.jpg'), Embedded=False, sizeoption=3)
def insert_res_speed_pic(hwp, br_name):
    hwp.MoveToField(f'응답속도',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/작업관리자.jpg'), Embedded=False, sizeoption=3)  
def insert_vol_afford_pic(hwp, br_name):
    hwp.MoveToField(f'자원사용율',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/D속성.jpg'), Embedded=False, sizeoption=3)   

# 그래프 그리는 테이블로 이동
def move_to_startofgraph_position(hwp,i):
    hwp.MoveToField(f'carrot2{{{{{i}}}}}',True,False,False)

def insert_jth_graph(hwp, j):
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/{channel_names[j+1]}_월통계.png'), Embedded=False, sizeoption=2)


# 한글파일 열기
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True # 아래한글 첫번째 파일 보이게해줘

# 파일경로 꼭 확인하기(절대경로로 작성)
hwp.Open(r"C:\Users\Y15599\Desktop\작업\Cablebridge_project_01\월간모니터링보고서(개별교량).hwp")

# 참조테이블 불러오기
info = pd.read_csv('01_channel_info/channel_info.csv', encoding='cp949')     # 채널정보
pic_path = r"C:\Users\Y15599\Desktop\작업\Cablebridge_project_01\2025-05"               # 가동율 그래프를 담고있는 폴더의 절대경로

n = 7 # 보고서 작성할 교량 번호
filtered_info = info[info['no.'] == n]
channel_names = filtered_info['channel_name'].tolist()
br_name = filtered_info['br_name'].iloc[0]
hwp.PutFieldText('교량명', br_name)

# 엑셀 파일 경로
excel_file = f'{br_name}/{br_name}_요약보고서.xlsx'

# 엑셀 데이터 불러오기
if os.path.exists(excel_file):
    excel_data = load_excel_data(excel_file)
    bridge_report = excel_data['reception']  # 데이터 수신율 시트 사용
    limit_report = excel_data['limit']       # 관리기준 초과 시트 사용
else:
    print(f"엑셀 요약 보고서가 존재하지 않습니다: {excel_file}")
    pass

insert_rate_operate(hwp, br_name)
i = 0
# 테이블 생성 지정된 캐럿으로 이동 표의 첫번째
print(f'carrot{{{{{i}}}}}')
for j in range(len(channel_names)-1):
    hwp.MoveToField(f'carrot{{{{{i}}}}}',True,False,False)
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableAppendRow")

# 그래프를 위한 테이블 생성; 지정된 캐럿(carrot2)으로 이동 표의 첫번째
print(f'carrot2{{{{{i}}}}}')
for j in range(len(channel_names)-1):
    hwp.MoveToField(f'carrot2{{{{{i}}}}}',True,False,False)
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableAppendRow")
    print(f'{j}/{len(channel_names)}')

# 채널구분 및 채널명 작성
hwp.MoveToField(f'carrot{{{{{i}}}}}',True,False,False)
for j in range(len(channel_names)):
    channel_type = filtered_info['센서종류'].iloc[j]
    insert_text(channel_type)
    hwp.HAction.Run("TableRightCell")
    insert_text(j+1)
    hwp.HAction.Run("TableRightCell")
    insert_text(channel_names[j])
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")

# 데이터 수신율 확인
move_to_start_position(hwp, i)
abnormal_count = 0
print(bridge_report.info())
with open(f'{br_name}/{br_name}_log.txt', 'w', encoding='utf-8') as f:
    for j in range(len(channel_names)):
        try:
            reception_rate = float(bridge_report['Unnamed: 3'].iloc[j])
            if reception_rate <= 20:
                insert_text(f'데이터 수신불량 수신율:{reception_rate}%')
                move_to_next_row(hwp, 2)
                insert_text('O')
                log_msg = f'{br_name} 채널명 {channel_names[j]}에서 데이터수신율 {reception_rate}%'
                print(log_msg)
                f.write(log_msg + '\n')
                move_to_next_row(hwp, 6)   
                abnormal_count += 1
            elif reception_rate <= 50:
                insert_text(F'일부구간 결측 {reception_rate}%')
                move_to_next_row(hwp, 1)
                insert_text('O')
                log_msg = f'{br_name} 채널명 {channel_names[j]}에서 데이터수신율 {reception_rate}%'
                print(log_msg)
                f.write(log_msg + '\n')
                abnormal_count += 1
                move_to_next_row(hwp, 7)    
            elif reception_rate <= 80:
                insert_text('일부구간 결측')
                move_to_next_row(hwp, 3)
                insert_text('O')
                log_msg = f'{br_name} 채널명 {channel_names[j]}에서 데이터수신율 {reception_rate}%'
                print(log_msg)
                f.write(log_msg + '\n')
                move_to_next_row(hwp, 5)  
            else:
                move_to_next_row(hwp)
        except Exception as e:
            log_msg = f'{br_name} 채널명 {channel_names[j]}에서 데이터 수신율 확인 중 오류 발생: {str(e)}'
            print(log_msg)
            f.write(log_msg + '\n')
            move_to_next_row(hwp)

    # 관리기준 초과 확인
    move_to_start_position(hwp, i)
    for j in range(len(channel_names)):
        try:
            if limit_report['Unnamed: 5'].iloc[j+1] == '초과':
                insert_text('관리기준 초과\r\n')
            move_to_next_row(hwp)
        except:
            log_msg = f'{br_name} 채널명 {channel_names[j]}에서 관리기준 초과 확인 중 오류 발생'
            print(log_msg)
            f.write(log_msg + '\n')

    # 그림넣는 루프
    hwp.MoveToField(f'carrot2{{{{{i}}}}}',True,False,False)
    hwp.Run("SelectAll")
    hwp.Run("Delete")
    hwp.InsertPicture(os.path.join(pic_path, f'{br_name}/{channel_names[0]}_월통계.png'), Embedded=False, sizeoption=2)
    hwp.HAction.Run("TableRightCell")
    for j in range(len(channel_names)-1):
        try:
            insert_jth_graph(hwp, j)
            hwp.HAction.Run("TableRightCell")
        except:
            log_msg = f'{br_name} 채널명 {channel_names[j]}에서 그래프 삽입중 오류 발생'
            print(log_msg)
            f.write(log_msg + '\n')
            insert_text('데이터 없음')
            hwp.HAction.Run("TableRightCell")

print(f'{abnormal_count}개 채널 일부구간 결측')

insert_daq_pic1(hwp, br_name)
insert_daq_pic2(hwp, br_name)
insert_v3_pic(hwp, br_name)
insert_eqk_pic(hwp, br_name)
insert_res_speed_pic(hwp, br_name)
insert_vol_afford_pic(hwp, br_name)

# 월간모니터링 결과 총괄  + 교량별 정상이상 수량 집계
hwp.PutFieldText('기준수량', len(channel_names))
hwp.PutFieldText('정상수량', len(channel_names)-abnormal_count)
hwp.PutFieldText('비정상수량', abnormal_count)
hwp.PutFieldText('전달대비변동', '-')
hwp.PutFieldText('점검내용 작성', f'{len(channel_names)} 중 {abnormal_count}개 채널 일부구간 결측')
hwp.SaveAs(f'C:/Users/Y15599/Desktop/작업/Cablebridge_project_01/2025-04/거북선대교/{br_name}_월간모니터링보고서(4월).hwp')

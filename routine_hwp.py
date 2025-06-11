import pandas as pd
import os
import win32com.client as win32

# 한글문서 텍스트입력 
def insert_text(text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

# 다음 행으로 이동
def move_to_next_row(hwp, count=9):
    for _ in range(count):
        hwp.HAction.Run("TableRightCell")



# 한글파일 열기
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True # 아래한글 첫번째 파일 보이게해줘

# 파일경로 꼭 확인하기(절대경로로 작성)
hwp.Open(r"C:\Users\USER\Desktop\특수교성과품\정기점검표(양식).hwp")

# 참조값 설정
info = pd.read_csv('01_channel_info/channel_info.csv', encoding='cp949')     # 채널정보
path = r"C:\Users\USER\Desktop\특수교성과품\Cablebridge_project_01\2025-05"
def insert_signiture(name, name1):
    hwp.MoveToField(f'{name}',True,False,False)
    hwp.Run('SelectAll')
    hwp.Run('Delete')
    hwp.InsertPicture(os.path.join(path, f'{name1}.png'))

# 센서 현황 작성하기
i = 1
hwp.PutFieldText('점검일시','-')
hwp.PutFieldText('점검자1','-')
hwp.PutFieldText('점검자2','-')
hwp.PutFieldText('확인자','-')
insert_signiture('서명1','점검자1')
insert_signiture('서명2','점검자2')
insert_signiture('서명3','확인자')

filtered_info = info[info['no.'] == i]
channel_names = filtered_info['channel_name'].tolist()
br_name = filtered_info['br_name'].iloc[0]

# 테이블 생성 (지정한 캐럿으로 이동해서 행추가를 (교량개수-2)회 수행
for j in range(len(channel_names)-2):
    hwp.MoveToField('carrot1',True,False,False)
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableAppendRow")

# 채널구분 및 채널명 작성
hwp.MoveToField('carrot1',True,False,False)
for j in range(len(channel_names)):
    channel_type = filtered_info['센서종류'].iloc[j]
    insert_text(channel_type)
    hwp.HAction.Run("TableRightCell")
    insert_text(channel_names[j])
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    hwp.HAction.Run("TableRightCell")
    print(f'{br_name} {channel_names[j]}')

hwp.SaveAs(f"C:/Users/USER/Desktop/특수교성과품/Cablebridge_project_01/2025-05/정기점검표({br_name}).hwp")

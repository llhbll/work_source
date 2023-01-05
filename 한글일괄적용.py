import win32com.client as win32
import pandas as pd

excel = pd.read_excel('D:\파이썬\study_source\data\sample\일괄입력명단.xlsx')

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerMoudle')


hwp.Run("FileNew")
hwp.Open("D:\파이썬\study_source\data\sample\일괄입력.hwp")
field_list = [i for i in hwp.GetFieldList().split('\x02')]
hwp.Run('SelectAll')
hwp.Run('Copy')

hwp.MovePos(3)

for i in range(len(excel) - 1):
    hwp.Run('Paste')
    hwp.MovePos(3)

for page in range(len(excel)):
    for field in field_list:
        hwp.PutFieldText(f'{field}{{{{{page}}}}}', excel[field].iloc[page])
#일상의 코딩 님 소스 (유튜브 문서비교 툴 구현해보기)
import difflib

import pyperclip as cb
import win32com.client as win32

def 글자색(Color):
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run(f"CharShapeTextColor{Color.capitalize()}")
    hwp.HAction.Run("Cancel")

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL","SecurityMoudle")

hwp.Run("FileNew")
hwp.Open(r"D:\파이썬\study_source\data\sample\aaa.hwp")
hwp.Run("FileNew")
hwp.Open(r"D:\파이썬\study_source\data\sample\bbb.hwp")
hwp.Run("FileNew")
hwp.Open(r"D:\파이썬\study_source\data\sample\ccc.hwp")


원본 = hwp.XHwpDocuments.Item(1)
사본 = hwp.XHwpDocuments.Item(2)
비교 = hwp.XHwpDocuments.Item(3)


원본.SetActive_XHwpDocument()
hwp.InitScan()
original_full_text = ""
stop_signal = True
while stop_signal:
    signal, text = hwp.GetText()
    original_full_text += text
    if signal == 1:
        break
hwp.ReleaseScan()
original_full_text = original_full_text.split('\r\n')[:-1]

사본.SetActive_XHwpDocument()
hwp.InitScan()
copy_full_text = ""
stop_signal = True
while stop_signal:
    signal, text = hwp.GetText()
    copy_full_text += text
    if signal == 1:
        break
hwp.ReleaseScan()
copy_full_text = copy_full_text.split('\r\n')[:-1]

비교.SetActive_XHwpDocument()
for original_statement in original_full_text:
    cb.copy(original_statement)
    hwp.Run('Paste')
    hwp.Run('TableRightCellAppend')
    coupled_dict = dict()
    for copy_statement in copy_full_text:
        coupled_dict[difflib.SequenceMatcher(None, original_statement.split(' ', 1)[1],
                                             copy_statement.split(' ', 1)[1]).ratio()] = copy_statement
    ratio = max(k for k, v in coupled_dict.items())
    cb.copy(coupled_dict[max(k for k, v in coupled_dict.items())].strip())
    hwp.Run('Paste')
    if ratio < 1.0:
        글자색('red')

    hwp.Run('TableRightCellAppend')

hwp.SaveAs(r'D:\파이썬\study_source\data\sample\ddd.hwp')
hwp.Quit()


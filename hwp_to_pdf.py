# -*- coding: utf-8 -*-
"""사업 계획서 수정본.hwp → PDF 변환 (시각 검증용)"""
import win32com.client as win32

HWP = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산\사업 계획서 수정본.hwp"
PDF = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산\사업 계획서 수정본.pdf"

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass
hwp.XHwpWindows.Item(0).Visible = False
hwp.SetMessageBoxMode(0x2FF1)
hwp.Open(HWP, "HWP", "")

print("Pages:", hwp.PageCount)

hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = PDF
hwp.HParameterSet.HFileOpenSave.Format = "PDF"
hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
print("PDF saved:", PDF)
hwp.Quit()

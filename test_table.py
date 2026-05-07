# -*- coding: utf-8 -*-
"""표 생성 단독 테스트 - 셀에 글자가 들어가는지 확인"""
import os
import win32com.client as win32

OUT = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산\_test_table.hwp"

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass
hwp.XHwpWindows.Item(0).Visible = False
hwp.SetMessageBoxMode(0x2FF1)
hwp.HAction.Run("FileNew")

def insert_text(t):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = t
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

# 본문에 텍스트
insert_text("Before table")
hwp.HAction.Run("BreakPara")

# 표 만들기 직전 위치 저장 (부모 단락의 표 삽입 지점)
saved_pos = hwp.GetPos()
print("Pos before table:", saved_pos)

# 표 생성 - HeightType=0 (자동)로 시도
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
tc = hwp.HParameterSet.HTableCreation
tc.Rows = 3
tc.Cols = 3
tc.WidthType = 2
tc.HeightType = 0   # 자동
tc.WidthValue = hwp.MiliToHwpUnit(150)
tc.HeightValue = hwp.MiliToHwpUnit(36)
tc.CreateItemArray("ColWidth", 3)
for i in range(3):
    tc.ColWidth.SetItem(i, hwp.MiliToHwpUnit(50))
tc.CreateItemArray("RowHeight", 3)
for i in range(3):
    tc.RowHeight.SetItem(i, hwp.MiliToHwpUnit(12))
hwp.HAction.Execute("TableCreate", tc.HSet)

print("After TableCreate, GetPos:", hwp.GetPos())

# 9개 셀 채우기
labels = ["A1","A2","A3","B1","B2","B3","C1","C2","C3"]
for i, lab in enumerate(labels):
    insert_text(lab)
    print(f"After {lab}, pos:", hwp.GetPos())
    if i < 8:
        hwp.HAction.Run("TableRightCell")
        print(f"  After TableRightCell, pos:", hwp.GetPos())

# 표 빠져나오기 - SetPos로 부모 단락 복귀 후 MoveLineEnd로 표 통과
print("Pos in last cell:", hwp.GetPos())
hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
print("After SetPos, pos:", hwp.GetPos())
hwp.HAction.Run("MoveLineEnd")
print("After MoveLineEnd, pos:", hwp.GetPos())
hwp.HAction.Run("BreakPara")
print("After BreakPara, pos:", hwp.GetPos())
insert_text("After table")

hwp.SaveAs(OUT, "HWP", "")
print("Saved:", OUT)

# 다시 열어서 표 컨트롤들의 셀 크기 확인
hwp.HAction.Run("FileNew")
hwp.Open(OUT, "HWP", "")

ctrl = hwp.HeadCtrl
while ctrl:
    if ctrl.CtrlID == "tbl":
        print("Found table")
        prop = ctrl.Properties
        try:
            print("  CellLineColor:", prop.Item("CellLineColor"))
        except: pass
        try:
            print("  PageBreak:", prop.Item("PageBreak"))
        except: pass
        # iterate properties
        try:
            for k in ["CellMarginLeft", "CellMarginRight", "CellMarginTop", "CellMarginBottom", "AutoMakeListNum"]:
                try:
                    print(f"  {k}:", prop.Item(k))
                except: pass
        except: pass
    ctrl = ctrl.Next

# PDF 변환으로 시각 확인
PDF = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산\_test_table.pdf"
try:
    hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = PDF
    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
    hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
    print("PDF saved:", PDF)
except Exception as e:
    print("PDF err:", e)

hwp.Quit()

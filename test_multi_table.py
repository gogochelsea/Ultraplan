# -*- coding: utf-8 -*-
"""연속 표 생성 테스트 - 두번째 TableCreate 실패 원인 추적"""
import os
import win32com.client as win32

OUT = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산\_test_multi.hwp"

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

def make_table(rows, cols, col_widths_mm, row_h_mm=12):
    nrows = rows
    total_w = sum(col_widths_mm)

    saved_pos = hwp.GetPos()
    print(f"  Before TableCreate({rows}x{cols}): pos={saved_pos}")

    # Use unique action key approach: HAction.GetDefault then Execute
    act = hwp.CreateAction("TableCreate")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Rows", nrows)
    pset.SetItem("Cols", cols)
    pset.SetItem("WidthType", 2)
    pset.SetItem("HeightType", 0)
    pset.SetItem("WidthValue", hwp.MiliToHwpUnit(total_w))
    pset.SetItem("HeightValue", hwp.MiliToHwpUnit(row_h_mm * nrows))
    # Try with col widths via CreateItemArray on fresh pset
    arr = pset.CreateItemArray("ColWidth", cols)
    for i, w in enumerate(col_widths_mm):
        arr.SetItem(i, hwp.MiliToHwpUnit(w))
    arr2 = pset.CreateItemArray("RowHeight", nrows)
    for i in range(nrows):
        arr2.SetItem(i, hwp.MiliToHwpUnit(row_h_mm))
    act.Execute(pset)
    print(f"  TableCreate OK")

    # fill
    for r in range(nrows):
        for c in range(cols):
            insert_text(f"R{r}C{c}")
            if not (r == nrows - 1 and c == cols - 1):
                hwp.HAction.Run("TableRightCell")

    # exit
    hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.Run("BreakPara")
    print(f"  After exit: pos={hwp.GetPos()}")

print("=== Table 1 (4x4) ===")
insert_text("Section 1")
hwp.HAction.Run("BreakPara")
make_table(4, 4, [30, 40, 30, 40])

print("=== Table 2 (5x2) ===")
insert_text("Section 2")
hwp.HAction.Run("BreakPara")
make_table(5, 2, [40, 100])

print("=== Table 3 (3x3) ===")
insert_text("Section 3")
hwp.HAction.Run("BreakPara")
make_table(3, 3, [50, 50, 50])

hwp.SaveAs(OUT, "HWP", "")
print("Saved.")

PDF = OUT.replace(".hwp", ".pdf")
hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = PDF
hwp.HParameterSet.HFileOpenSave.Format = "PDF"
hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
print("PDF saved:", PDF)
hwp.Quit()

# -*- coding: utf-8 -*-
"""
사업 계획서.hwp 글자 크기 검증·교정
- 표지(첫 페이지)는 예외, 그 외 모든 단락의 첫 글자 크기를 점검
- 12pt(=Height 1200)가 아닌 단락은 12pt로 강제 변환
"""
import win32com.client as win32
PATH = r"C:\Users\filk2\Desktop\공부자료\학교수업\화공산\사업 계획서.hwp"

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass
hwp.XHwpWindows.Item(0).Visible = False
hwp.SetMessageBoxMode(0x2FF1)
hwp.Open(PATH, "HWP", "")

# 총 페이지
total_pages = hwp.PageCount
print("Total pages:", total_pages)

# 본문 페이지(2 페이지)부터 끝까지 단락 단위 점검
hwp.HAction.Run("MoveDocBegin")
# 표지(1쪽) 건너뜀 → 2쪽 시작으로
hwp.MovePos(3)  # MoveTopOfFile
# 다음 페이지 시작점으로 이동
hwp.HAction.Run("MovePageDown")

# 본문 시작 위치 기록
hwp.HAction.Run("MoveLineStart")

# 단락 순회: 한 단락씩 선택→Height 검사→1200 아니면 1200으로 변경
fixed_count = 0
checked_count = 0
problems = []

para_no = 0
while True:
    # 현재 단락 시작으로
    hwp.HAction.Run("MoveLineStart")
    # 한 글자 선택해서 그 글자의 CharShape를 읽음
    hwp.HAction.Run("MoveSelLineEnd")
    # 빈 단락이면 길이 0 — 그 경우 그냥 다음으로
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    h = hwp.HParameterSet.HCharShape.Height
    # 선택 해제
    hwp.HAction.Run("Cancel")
    checked_count += 1
    if h != 0 and h != 1200:
        problems.append((para_no, h))
        # 단락 전체 선택해서 Height만 변경 (다른 속성 유지)
        hwp.HAction.Run("MoveLineStart")
        hwp.HAction.Run("MoveSelLineEnd")
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        hwp.HParameterSet.HCharShape.Height = 1200
        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
        hwp.HAction.Run("Cancel")
        fixed_count += 1
    para_no += 1
    # 다음 줄로
    moved = hwp.HAction.Run("MoveDown")
    if not moved:
        break
    if para_no > 1500:
        break

print(f"Checked: {checked_count}, Problems: {len(problems)}, Fixed: {fixed_count}")
if problems[:30]:
    print("Problem samples:", problems[:30])

# 저장
hwp.Save()
hwp.Quit()
print("Done.")

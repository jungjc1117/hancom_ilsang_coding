import PIL.Image as pilimg
import numpy as np
import win32com.client as win32

# 이미지 불러오기
im = pilimg.open(r"C:\Users\jjc\Desktop\한컴오피스 업무자동화 입문\[15]_[파이썬-아래아한글 자동화 입문] 셀채우기 코드로 아름다운 모자이크아트 만들어보기.png")
# 픽셀정보를 np.array로
pix = np.array(im)

hwp =win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
# 100×50 표가 생성되어 있는 한글파일 불러오기
hwp.Open(r"C:\Users\jjc\Desktop\한컴오피스 업무자동화 입문\[15]_[파이썬-아래아한글 자동화 입문] 셀채우기 코드로 아름다운 모자이크아트 만들어보기.hwpx")
#%%
hwp.Run("TableCellBlock")
for row in pix:
    for col in row:
        hwp.HAction.GetDefault("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
        hwp.HParameterSet.HCellBorderFill.FillAttr.type = hwp.BrushType("NullBrush|WinBrush")
        hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = \
        hwp.RGBColor(col[0], col[1], col[2])
        hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = hwp.RGBColor(0, 0, 0)
        hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = hwp.HatchStyle("None")
        hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
        hwp.HAction.Execute("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
        hwp.Run("TableRightCell")
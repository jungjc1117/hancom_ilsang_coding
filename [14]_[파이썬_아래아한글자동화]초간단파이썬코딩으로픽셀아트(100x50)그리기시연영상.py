import PIL.Image as pilimg
import numpy as np
import win32com.client as win32
from random import randrange

# 이미지 불러오기
im = pilimg.open(r"C:\Users\jjc\Desktop\한컴오피스 업무자동화 입문\[14]_[파이썬-아래아한글자동화] 초간단 파이썬코딩으로 픽셀아트(100x50) 그리기 시연영상.png")
# 이미지를 numpy 배열로 변환
pix = np.array(im)

hwp =win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
hwp.Open(r"C:\Users\jjc\Desktop\한컴오피스 업무자동화 입문\[14]_[파이썬-아래아한글자동화] 초간단 파이썬코딩으로 픽셀아트(100x50) 그리기 시연영상.hwpx")
hwp.HAction.Run("FrameFullScreen")
hwp.XHwpWindows.Item(0).Visible = True
#%%
target_coords = []
for i in range(50):
    for j in range(100):
        target_coords.append(f"{i}_{j}")

for row in pix:
    for col in row:
        target_coord = target_coords.pop(randrange(len(target_coords)))
        hwp.MoveToField(target_coord, select=False)
        hwp.HAction.GetDefault("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
        hwp.HParameterSet.HCellBorderFill.FillAttr.type = hwp.BrushType("NullBrush|WinBrush")
        hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = hwp.RGBColor(*pix[int(target_coord.split("_")[0])][int(target_coord.split("")[1])])
        hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = hwp.RGBColor(0, 0, 0)
        hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = hwp.HatchStyle("None")
        hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
        hwp.HAction.Execute("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
        print(len(target_coords))
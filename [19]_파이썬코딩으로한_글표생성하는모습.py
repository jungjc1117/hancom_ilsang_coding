"""
아래의 148mm는 종이여백 210mm에서 60mm(좌우 각 30mm)를 뺀 150mm에다가,
표 바깥여백 각 1mm를 뺀 148mm이.(TableProperties.Width = 41954)
각 열의 너비는 5개 기준으로 26mm인데, 이는 셀마다 안쪽여백 좌우 각각 1.8mm를 뺀 값으로,
148 - (1.8 x 10 =) 18mm= 130mm
그래서 셀 너비의 총 합은 130이 되어야 한다.
아래의 라인28~32까지 셀너비의 합은 16+36+46+16+16=130
표를 생성하는 시점에는 표 안팎의 여백을 없애거나 수정할 수 없으므로
이는 고정된 값으로 간주해야 한다.
"""
#%%
import win32com.client as win32
#%%
hwp =win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp =win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True

# %%
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 5
hwp.HParameterSet.HTableCreation.Cols = 5
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HParameterSet.HTableCreation.HeightType = 1
hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(148.0)
hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(150)
hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 5)
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(16.0))
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(36.0))
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(46.0))
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3, hwp.MiliToHwpUnit(16.0))
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(4, hwp.MiliToHwpUnit(16.0))
hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(40.0))
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(20.0))
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(2, hwp.MiliToHwpUnit(50.0))
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(3, hwp.MiliToHwpUnit(20.0))
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(4, hwp.MiliToHwpUnit(20.0))
hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1 # 셀을 문자로 취급
hwp.HParameterSet.HTableCreation.TableProperties.Width = hwp.MiliToHwpUnit(148)
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

hwp.MovePos(3) # 문서 끝으로 이동
# %%

import datetime as dt
import win32com.client as win32

def init_hwp():
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True
        hwp.Open(r"C:\Users\jjc\Desktop\[09]_.hwpx")
        return hwp
    except Exception as e:
        print(f"Error initializing HWP: {e}")
        return None

def insert_text(text):
    act = hwp.CreateAction("InsertText")
    pset = act.CreateSet()
    pset.SetItem("Text", text)
    act.Execute(pset)

def week(month, day):
    weekday = "월화수목금토일"
    return weekday[dt.date(2022, month, day).weekday()]

def get_weekday(text):
    month, day = [int(i) for i in text.split(".")[:2]]
    return f"( {week(month, day)} )"

if __name__ == "__main__":
    hwp = init_hwp()
    if hwp is not None:
        hwp.FindCtrl()
        hwp.Run("ShapeObjTableSelCell")
        while True:
            hwp.InitScan(Range=0xff)
            state, text = hwp.GetText()
            hwp.ReleaseScan()

            if text.endswith("."):
                hwp.Run("Cancel")
                hwp.Run("MoveLineEnd")
                insert_text(get_weekday(text))
                hwp.Run("TableCellBlock")

            if not hwp.HAction.Run("TableRightCell"):
                break
        hwp.Save()

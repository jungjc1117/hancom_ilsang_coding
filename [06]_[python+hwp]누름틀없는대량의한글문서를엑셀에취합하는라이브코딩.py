from tkinter.filedialog import askopenfilenames
import win32com.client as win32

def get_text():
    hwp.InitScan(Range=0xff)
    total_text = ""

    state = 2
    while state not in [0, 1]:
        state, text = hwp.GetText()
        total_text += text

    hwp.ReleaseScan()
    return total_text

filelist = askopenfilenames()

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

excel = win32.gencache.EnsureDispatch("Excel.Application")
wb = excel.Workbooks.Open(r"C:\users\jjc\Desktop\[06]_.xlsx")
ws = wb.Worksheets(1)
excel.Visible = True

for file in filelist:
    hwp.Open(file)

    ctrl = hwp.HeadCtrl
    while ctrl:
        if ctrl.CtrlID == "tbl":
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            break
        else:
            ctrl = ctrl.Next

    # 표 안에 진입 후 작업
    hwp.FindCtrl()
    hwp.Run("Shape0bjTableSelCell")
    contents = []
    contents.append(get_text())
    while hwp.HAction.Run("TableRightCell"):
        contents.append(get_text())

    과제명 = contents[1]
    신청부서 = contents[3].split("\r\n")[0].replace("/", "")
    과제담당관 = contents[3].split("\r\n")[1].replace("/", "")
    담당공무원 = contents[5]

    연구방식_ = ["위탁형", "공동연구형", "자문형"]
    연구방식 = [i.strip() for i in contents[7].split("(")][1:]
    for idx, text in enumerate(연구방식):
        if not text.startswith(")"):
            연구방식 = 연구방식_[idx]
            break

    연구시작 = contents[9].split("~")[0].strip()
    연구종료 = contents[9].split("~")[1].split("(")[0].strip()
    연구기간 = contents[9].split("(")[1].replace(")", "")

    예산항목_ = ["포괄", "사업별"]
    예산항목 = [i.strip() for i in contents[12].split("(")[1:]]
    for idx, text in enumerate(예산항목):
        if not text.startswith(")"):
            예산항목 = 예산항목_[idx]
            break

    예상금액 = contents[15]
    연구필요성 = contents[17]
    중복검토방법 = contents[21].split("\r\n")[1].replace("-", "").strip()

    중복성여부_ = ["있다", "없다"]
    중복성여부 = [i.strip() for i in contents[21].split("\r\n")[2].split("(")][1:]
    for idx, text in enumerate(중복성여부):
        if not text.startswith(")"):
            중복성여부 = 중복성여부_[idx]
            break

    연구내용 = contents[23]
    연구결과활용방안 = contents[25]

    입력행 = ws.UsedRange.Rows.Count + 1

    ws.Range(ws.Cells(입력행, 1), ws.Cells(입력행, 15)).Value = \
        과제명, 신청부서, 과제담당관, 담당공무원, 연구방식, 연구시작, 연구종료, 연구기간, 예산항목, 예상금액, 연구필요성, 중복검토방법, 중복성여부, 연구내용, 연구결과활용방안

wb.Save()
hwp.Clear(isDirty=True)
hwp.Quit()

import win32com.client as win32

# 한글 객체 생성 및 문서 열기
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

hwp.Open(r"C:\Users\jjc\Desktop\[07]_.hwpx")

# 첫 번째 컨트롤로 이동
ctrl = hwp.HeadCtrl
while ctrl:
    if ctrl.CtrlID == "tbl":  # 테이블 컨트롤을 찾음
        hwp.SetPosBySet(ctrl.GetAnchorPos(0))  # 테이블 위치로 이동
        hwp.FindCtrl()  # 컨트롤 찾기
        hwp.Run("ShapeObjTableSelect")  # 테이블 선택
        hwp.Run("SelectAll")  # 전체 선택
        hwp.Run("Cut")  # 잘라내기
        try:
            hwp.DeleteCtrl(ctrl)  # 컨트롤 삭제 시도
        except Exception as e:
            print(f"Error deleting control: {e}")
        hwp.Run("Paste")  # 잘라낸 내용 붙여넣기
        break
    ctrl = ctrl.Next  # 다음 컨트롤로 이동

# 변경 사항 저장
hwp.Save()

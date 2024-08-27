'''
아래아한글 파일을 열어서-
각각의 페이지를 한 개의 한글파일로 저장하는 프로그램입니다.
기본 저장폴더는 바탕화면의 result 폴더이며, 폴더가 없다면 생성합니다.
'''

import os
from time import sleep
from tkinter.filedialog import askopenfilename

import win32com.client as win32

#askdirectory

# 클래스 정의
class Hwp:
    def __init__(self):
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    def __del__(self) :
        self.hwp.Clear(option=1) # 0:팝업, 1:버리기, 2:저장팝업, 3:무조건저장(빈 문서#은 버림)
        self.hwp.Quit()

    def open_file(self, filename, view=False):
        self.name = filename
        self.hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        if view == True:
            self.hwp.Run("FileNew")
        self.hwp.Open(self.name)

    def split_save(self, name):
        self.hwp.MovePos(0)
        self.pagecount = self.hwp.PageCount
        hwp_docs = self.hwp.XHwpDocuments

        target_folder = os.path.join(os.environ['USERPROFILE'], 'desktop', 'result')
        try:
            os.mkdir(target_folder)
        except FileExistsError:
            print("바탕화면에 result 폴더가 이미 생성되어 있습니다.")

        for i in range(self.pagecount):
            hwp_docs.Item(0).SetActive_XHwpDocument() # 이걸 실행해주고
            sleep(1) # 이건 컴퓨터 사양을 좀 타는데 ...왠만하면 생략해도 될듯
            self.hwp.Run("CopyPage")
            sleep(1)
            hwp_docs.Add(isTab=True) # True를 넣으면 탭이 생긴다. False일 때는 새 창
            hwp_docs.Item(1).SetActive_XHwpDocument() # 이걸 실행해주고
            self.hwp.Run("Paste")
            self.hwp.SaveAs(
                os.path.join(target_folder, name.rsplit('/')[-1].rsplit('.')[0]+"_"+str(i+1)+".hwp"))
            self.hwp.Run("FileClose")
            self.hwp.Run("MovePageDown")
            print(f"{i +1}/{self.pagecount}")

    def quit(self):
        self.hwp.Quit()

# 메인함수 정의
def main():
    """..."""
    name = askopenfilename(initialdir=
                           os.path.join(os.environ["USERPROFILE"], "desktop"),
                            filetypes=(("아래아한글 파일", "*.hwp"), ("모든 파일", "*.*")),
                            title="HWP파일을 선택하세요.")

    hwp = Hwp ()
    hwp.open_file(name)
    hwp.split_save(name)
    hwp.quit()
    print('HWP 파일의 페이지별 개별저장이 완료되었습니다.창을 닫으셔도 좋습니다.')
    input()

# 메인함수 실행
if __name__ == '__main__':
    main()
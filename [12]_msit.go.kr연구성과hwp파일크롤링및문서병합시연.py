"""
과기정통부 연구성과 페이지의 모든 한/글 문서를 크롤링한 후
한 파일로 취합하는 코드.
pyautoaui 등 특정 키를 외부에서 입력하거나 특정 좌표를 클릭하는 코드를 완전히 배제함
"""
import win32com.client as win32
import winreg
import os
from time import sleep
import win32con
import win32gui
import win32ui
import shutil
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

과정보기 = True


def 시작(과정보기=True):
    """
    아래아한글을 여는 함수.
    백그라운드에서 자
    백그라운드에서 작업하려면 과정보이기 파라미터를 False로 선택하면 됨
    :return: hwp 오브젝트
    """
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject") # 아래아한글 열기
    hwp.RegisterModuLe("FilePathCheckDLL", "FilePathCheckerModule") # 보안모듈 실행
    hwp.XHwpWindows.Item(0).Visible = 과정보기 # 숨김해제(기본적으로 숨김상태에서 시작함)
    return hwp

def 새_여백구역_추가(hwp, margin_dict, ApplyTo=4) :
    """
    문서 하단에 새 여백구역을 추가하는 함수
    :param hwp: hwp.오브젝트
    :param margin_dict: 여백과 종횡, 용지설정이 들어있는 사전객체.하단의 "페이지별_여백추출"의 원소
    :param ApplyTo: 4는 새 구역으로 여백설정 추가(기본), 3은 문서전체의 여백을 수정하도록 함
    :return: 반환값 없음
    """
    hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
    hwp.HParameterSet.HSecDef.PageDef.PaperWidth = margin_dict["PaperWidth"] # 용지너비
    hwp.HParameterSet.HSecDef.PageDef.PaperHeight = margin_dict["PaperHeight"] # 8=0
    hwp.HParameterset.HSecDef.Pagebef.Papereight = margin_dict["PaperHeight"] # 8|0|
    hwp.HParameterSet.HSecDef.PageDef.LeftMargin =margin_dict["LeftMargin"] # 외쪽여백
    hwp.HParameterSet.HSecDef.PageDef.RightMargin=margin_dict["RightMargin"] # 오른쪽여백
    hwp.HParameterSet.HSecDef.PageDef.TopMargin= margin_dict["TopMargin"] # 오른쪽여백
    hwp.HParameterSet.HSecDef.PageDef.BottomMargin = margin_dict["BottomMargin"] # 오른쪽여백
    hwp.HParameterSet.HSecDef.PageDef.HeaderLen=margin_dict["HeaderLen"] # 오른쪽여백
    hwp.HParameterSet.HSecDef.PageDef.FooterLen= margin_dict["FooterLen"] # 오른쪽여백
    hwp.HParameterSet.HSecDef.PageDef.Gutterlen = margin_dict["Gutterlen"] # 제본길이
    hwp.HParameterset.HSecDef.Pagebel.Gutterlype = margin_dict["GutterType"] # 제본타입
    hwp.HParameterSet.HSecDef.PageDef.Landscape=margin_dict["Landscape"] # 종횡
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", ApplyTo) # 4는 새 구역으로, 3은 문서전체
    hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)

def 페이지별_여백추출(hwp):
    """
    모든 페이지의 여백정보를 저장한 2중dict를 반환하는 함수.
    "PageSetup", "PageDef" 및 "PaperWidth" 등의 문자열은 API문서에서, 검색가능
    :param hwp: 한/글 오브젝트
    ireturn: 2중dict.
    """
    margin_dict = {} # 빈 사전 생성
    hwp.MovePos(2) # 문서의 시작으로 이동
    for page in range(hwp.PageCount): # 모든 페이지번호마다 반복하면서
        Act = hwp.CreateAction("PageSetup") # 페이지설정 정보를 열고
        Set = Act.CreateSet() # 빈 파라미터셋(그릇?) 생성
        Act.GetDefault(Set) # 현재 쪽의 여백값을 파라미터셋에 채음
        margin_dict[page] ={
        "PaperWidth" : Set.Item("PageDef").Item("PaperWidth"), # 용지너비 추출
        "PaperHeight": Set.Item("PageDef").Item("PaperHeight"), # 용지높이
        "LeftMargin" : Set.Item("PageDef").Item("LeftMargin"),# 외쪽여백
        "RightMargin" : Set.Item("PageDef").Item("TopMargin"),# 오른쪽여백
        "TopMargin" : Set.Item("PageDef").Item("TopMargin"), # 윗쪽여백
        "BottomMargin": Set.Item("PageDef").Item("BottomMargin"), # 아래쪽여백
        "HeaderLen":Set.Item("PageDef").Item("HeaderLen"),# 오른쪽여백
        "FooterLen": Set.Item("PageDef").Item("FooterLen"), # 오른쪽여백
        "GutterLen" : Set.Item( "Pageber" ).Item("GutterLen"), # 제본길이
        "GutterType": Set.Item("PageDef").Item("GutterType"), # 제본타입
        "Landscape": Set.Item("PageDef").Item("Landscape") # 종획
        }
        hwp.Run("MovePageDown") # 다음 페이지로 이동
    return margin_dict


def 모든페이지_여백이_동일한지_체크(리스트):
    """
    페이지별_여백추출 함수에서 반환한 2중사전에서
    모든 페이지의 여백을 비교해서 모두 동일한 경우 True를 반환하고
    여백설정이 두 가지 이상인 경우에는 False를 반환
    :param 리스트:
    .return:
    """
    결과 = True # 초기값은 True
    # current = None # 없어도 되는 라인이지만 else물에 문법경고가 떠서 놔뒀음
    for i in range(len(리스트)):
        if i == 0: # 첫 페이지의 여백을 current에 대입함
            current = 리스트[i]
        else:
            if i == current: # 첫 페이지와 다음페이지의 값이 같으면
                pass # 통과
            else: # 하나라도 다른 값이 있다면?
                결과 = False
        return 결과

def 글자삽입(hwp, 텍스트):
    """
    간단히 글자를 삽입하는 아래아한글 ARI명령었·
    아래의 _물임번호수정" 함수를 간소화하기 위해 별도로 정의함.
    스크립트매크로 중 가장 간단하고 일반적인 형태임.(액션생성-> 값설정 -> 실행)
    param hwp: 아래아한글 오브젝트
    :param 텍스트: 삽입하고자 하는 문자열
    :return: None
    """
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet) # 액션 생성
    hwp.HParameterSet.HInsertText.Text = 텍스트 # 액션에 값 설정
    hwp.HAction.Execute("InsertText",hwp.HParameterSet.HInsertText.HSet) # 액션 실행

def 붙임번호수정(hwp, 인덱스):
    """
    "붙임#" 안의 #인덱스를 순서대로 매기기 위한 함수.
    :param hwp:
    :param 인덱스: int(순서)
    :return: Non
    """
    hwp.InitScan() # 문서 탐색 시작
    while True:
        result = hwp.GetText() # 문단별로 텍스트와 상태코드 언기
        if result[0] == 1: # 상태코드1 == 문서 끝에 도달하며
            break # while물 종
        elif result[0] == 4 and result[1].startswith("붙임"): # 상태코드4 == 표 안으로 진입
            hwp.MovePos(201) # 탐색된 위치로 캐럿 이동
            hwp.Run("SelectAll") # Ctrl-A를 놀러서
            hwp.Run("Delete") # 셀 내용 삭제 후
            글자삽입(hwp, f"붙임{인덱스}") # 새로 "붙임#" 문자열 입력
            break # while 종료
        else: # 그 외에는
            pass # 그냥 넘어감

def 붙임번호재설정(hwp):
    """
    최종취합문서 붙임문서 번호를 처음부터 재설정하는 함수.
    참고:InitScan 메서드는 문서 수정 발생시 탐색이 중지되므로 수정 후 InitScan()을 재실행해야 함.
    :param hwp: 아래아한글 오브젝트
    :return: None
    """
    인덱스 = 1 # 붙임문서번호 초기값 1로 설정
    hwp.InitScan() # 문서 탐색 시작
    while True:
        result = hwp.GetText() # 문단별로 텍스트와 상태코드 얻기
        if result[0] == 1: # 상태코드1 == 문서 끝에 도달하면
            break # while문 종료
        elif result[0] == 4:
            if result[1].startswith(f"붙임{인덱스}"): # 인덱스 순서가 맞으면
                인덱스 +=1
            elif result[1].startswith("붙임"): # 인덱스 순서가 틀리면
                hwp.MovePos(201) # 탐색된 위치로 캐럿 이동
                hwp.Run("SelectAll") # Ctrl-A를 눌러서
                hwp.Run("Delete") # 셀 내용 삭제 후
                글자삽입(hwp, f"붙임{인덱스}") # 새로 "붙임#" 문자열 입력
                hwp.InitScan() # 처음부터 탐색 재시작
                인덱스 = 1 # 이데스번호도 1부터 재시간
            else:
                pass
        else: # 그 외에는
            pass # 그냥 넘어감
    hwp.ReleaseScan() # InitScan 탐색을 마친 후 꼭 실행해줘야 함.

def 특정문자열_표함한_표_삭제(hwp, 문자열):
    """
    문서 내에 특정 문자열을 포함한 표를 찾아서 모두 삭제해 주는 함수
    사용 예: "각 기관 스마트워크 시스템 개선방안이 잘 나타날 수 있도록 위 양식 일부 조정 가능" 표 삭제.
    :param hwp: 아래아하글 오브젝트
    :param 문자열: 이 문자열을 포함한 모든 표를 삭제· 유의하여 사용해야 함.
    :return: None.몇 개의 표를 삭제했는지 콘솔에 출력
    """
    hwp.InitScan() # 문서 탐색 시작
    del_num = 0
    while True:
        result = hwp.GetText() # 문단별로 텍스트와 상태코드 얻기
        if result[0] == 1: # 상태코드1 == 문서 끝에 도달하면
            break # while문 종료
        elif result[0] in [3, 4] and result[1] .__contains__(문자열): # 상태코드3:표 내부, 4:표 안으로 진입,
            hwp.MovePos (201) # 탐색된 위치로 캐럿 이동
            hwp.Run("CloseEx") # 표 밖으로 나가서
            hwp.FindCtrl() # 표 전체를 선택하고
            hwp.Run("Delete") # 삭제
            hwp.InitScan()
            del_num += 1
        else: #그 외에는
            pass # 그냥 넘어감
    print(f"{del_num}개의 표를 삭제하였습니다.")
    
def 익스플로러_다운경로():
    key =winreg.HKEY_CURRENT_USER
    key_value = r"Software\Microsoft\Internet Explorer\Main"
    open = winreg.OpenKey(key, key_value, 0, winreg.KEY_ALL_ACCESS)
    value = winreg.QueryValueEx(open, "Default Download Directory") [0]
    winreg.CloseKey(open)
    return value

저장경로 = 익스플로러_다운경로()
옮길경로 = os.path.join(os.environ["USERPROFILE"], "desktop", "저장경로")

def f_click(pycwnd):
    x = 300
    y = 300
    lParam = y << 15 | x
    pycwnd.SendMessage(win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
    pycwnd.SendMessage(win32con.WM_LBUTTONUP, 0, lParam)
    
def get_whndl():
    whndl = win32gui.FindWindowEx(0, 0, None, '다운로드 보기 - Internet Explorer')
    return whndl

def make_pycwnd(hwnd) :
    PyCWnd = win32ui.CreateWindowFromHandle (hwnd)
    return PyCWnd

def send_input (pycwnd):
    f_click(pycwnd)
    sleep(1)
    pycwnd.SendMessage(win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    pycwnd.SendMessage(win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    pycwnd.UpdateWindow()
    
def callback(hwnd, hwnds):
    if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
        hwnds[win32gui.GetClassName(hwnd) ] = hwnd
    return True

def 다운로드_완료시까지_대기(파일명):
    for root, dirs, files in os.walk(저장경로):
        for file in files:
            if 파일명 != "" and file.startswith(파일명.rsplit(".")[0].replace(" ", "+")):
                filepath = os.path.join(root, file)
                print(filepath)
                while True:
                    before = os.path.getsize(filepath)
                    sleep(1)
                    if os.path.getsize(filepath) - before == 0:
                        break
                    print(f"{파일명} 다운로드 중입니다.")
                print(f"{파일명} 다운로드가 완료되었습니다.")
                
if __name__ == '__main__':
    driver = webdriver .Ie(r"C:\Users\User\PycharmProjects\hwptest\IEDriverServer.exe")
    driver.get("about:blank")
    driver.maximize_window()
    
    hwnd_browser = win32gui.FindWindow(None, "빈 페이지 - Internet Explorer")
    # win32gui.ShowWindow(hwnd_browser, win32con.SW_HIDE);print("익스플로러 창을 숨깁니다.크롤링이 완료된 후 다시 나타납니다."
    
    주소 = "https://ww.msit.go.kr/"
    driver.get(주소)
    print(f"{주소} 로 이동합니다.")
    driver.execute_script("javascript: fn_menulast_go('user', '118', '119', '#');")
    driver.implicitly_wait(3)
    print("크롤링을 시작합니다.")
    페이지번호 = 1
    while True:
        a_list = [i.get_attribute("onclick") for i in
                  driver.find_elements_by_css_selector("div.board_list > div.toggle > a[title='|']")]
        글번호 = 1
        for js_art in a_list:
            print(f"{페이지번호}페이지의 {글번호}번 글을 열었습니다.")

            sleep(1)
            driver.execute_script(js_art)
            sleep(1)
            print(f"글 제목은 {driver.find_element_by_tag_name('h2').text}입니다.")
            try:
                js링크 = driver.find_element_by_css_selector("a[onclick^='fn_download(']").get_attribute("onclick")
                driver.find_element_by_tag_name("body").send_keys(Keys.END)
                sleep(1)
                파일명 = driver.find_element_by_css_selector("ul.down_file > li > a[title='새창열림']").text.replace(" ", "+")
                print(f"첨부파일 이름은 {파일명}입니다.")
                if 파일명:
                    pass
                else:
                    raise FileNotFoundError
                driver.execute_script(js링크)
                sleep(1)
                다운로드_완료시까지_대기(파일명=파일명)
                driver.find_element_by_tag_name("body").send_keys(Keys.LEFT_CONTROL + "j")
                hwnd_download =win32gui.FindWindow(None, "다운로드 보기 - Internet Explorer")
                # win32gui.ShowWindow(hwnd_download, win32con.SW_HIDE)
                sleep(3)

                whndl = get_whndl()
                hwnds = {}
                win32gui.EnumChildWindows(whndl, callback, hwnds)
                whndl = hwnds['DirectUIHWND']

                pycwnd = make_pycwnd(whndl)
                send_input(pycwnd)
                win32gui.PostMessage(hwnd_download, win32con.WM_CLOSE, 0, 0)
                print(f"{저장경로}에 {파일명} 다운로드를 완료하였습니다.")
                sleep(1)
                shutil.move(src=os.path.join(저장경로, 파일명), dst=os.path.join(옮길경로, 파일명))
            except NoSuchElementException:
                print("본 글에는 첨부파일이 없습니다.다음 글로 넘어갑니다.")
                pass
            sleep(3)
            driver.back()
            글번호 += 1
        sleep(1)
        try:
            driver.execute_script(
                driver.find_element_by_css_selector("div.board_paging span.btn > a.next").get_attribute("onclick"))
            print("다음 페이지로 이동합니다.")
            sleep(1)
            페이지번호 += 1
        except NoSuchElementException:
            print("마지막 페이지까지 크롤링을 완료하였습니다.")
            break
    win32gui.ShowWindow(hwnd_browser, win32con.SW_SHOW)
    
    #.문서취합 시작
    취합보고서 = r"C:\Users\User\Desktop\취합프로그램\취합보고서.hwp"
    os.chdir(r"C:\Users\User\Desktop\저장경로") # 문서경로 폴더로 이동
    붙임문서리스트 = [i for i in os.listdir() if (i.endswith(".hwp") or i.endswith(".hwpx"))] # 폴더 안의 한/글 파일명 추출.

    hwp = 시작(과정보기) # 취합보고서용 아래아한글, 오브젝트 열기
    hwp.Open(취합보고서) # 취합보고서 불러오기
    hwp.MovePos(3) # 문서의 끝으로 이동
    hwp1 = 시작(과정보기) # 개별파일 열기 위한 아래아한글 오브젝트 하나 더 열고
    for 인덱스, 문서 in enumerate(붙임문서리스트, start=1): # enumerate는 (인덱스번호, 내용) 튜플을 매번 리턴함
        hwp1.Open(os.path.join(os.getcwd(), 문서)) # (주의)파일을 불러올 때는 무조건 절대경로 전체를 입력해야 함.
        붙임번호수정(hwp1, 인덱스)
        문서여백 = 페이지별_여백추출(hwp1)
        if 모든페이지_여백이_동일한지_체크(문서여백) == False: # 문서 내 여백이 동일하지 않으면?
            raise AttributeError # 문서여백을 "사전 리스트"로.한 페이지씩 복사해서 넣어야 함.현재는 생략.
        else: # 문서 내 모든 페이지의 여백이 동일하면?
            문서여백 = 문서여백[0] # 문서여백을 "사전"으로
            새_여백구역_추가(hwp, 문서여백) # 위 여백설정으로 새 페이지 생성
            hwp1.Run("SelectAll") # 개별문서의 전체내용 선택
            hwp1.Run("Copy") # 클립보드에 복사
            hwp1.Run("CloseEx") # 선택해제(해제하지 않으면 Close 실행시 오류발생하는 경우가 종종 있음)
            hwp.Run("Paste") # 취합용 문서에 붙여넣기
        hwp1.XHwpDocuments.Item(0).Close(isDirty=False) # 아래 두 줄과 의미 동일
        # hwp1.Clear(option=1) # 변경된 내용을 버림
        # hwp1.Run("FileClose") # 파일 닫기
        print(f"{문서} 붙여넣기 완료[{인덱스}/{len(붙임문서리스트)}]")

    특정문자열_표함한_표_삭제(hwp, "각 기관 스마트워크 시스템 개선방안이 잘 나타날 수 있도록 위 양식 일부 조정 가능")
    hwp1.Quit() # 한/글 프로그램 종료
    붙임번호재설정(hwp)
    hwp.XHwpWindows.Item(0).Visible = True
    # hwp.Save() # 검토 필요하므로 취합보고서 저장은 생략
    # hwp.Quit() # 한/글 프로그램 종료도 생략
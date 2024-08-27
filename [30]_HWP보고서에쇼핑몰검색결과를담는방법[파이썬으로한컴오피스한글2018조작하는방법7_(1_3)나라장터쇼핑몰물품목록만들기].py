'''
1.나라장터 종합쇼핑몰에 들어가서, "작업용의자"로 검색한다.
2.조건은, 등판과 좌판이 "메시" 또는 "망사"여야 하며, "가죽"이 들어있지 않아야 한다.팔걸이와 머리받침판이 있어야 한다.
단, 우선구매대상 및 의무구매대상 인증을 받은 제품이어야 한다.
3.명세서는 엑셀로 작성해놓고, 이미지도 각각 다운받아놓는다.(한글보고서에 첨부예정)
4.아래한글 보고서를 작성한다.끝.
'''
import os
from io import BytesIO
from time import sleep
from urllib.request import urlretrieve as download

import pandas as pd
import win32clipboard
import win32com.client as win32
from PIL import Image
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

chromedriver_path = r"C:\Users\jjc\Desktop\한컴오피스 업무자동화 입문\chromedriver-win64\chromedriver.exe"
chrome_options = Options()
service = Service(executable_path=chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get("https://www.g2b.go.kr:8092/sm/ma/mn/SMMAMnF.do")
driver.switch_to.frame('sub')

search_box = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "input#kwd.srch_txt"))
)
search_box.send_keys("작업용 의자")

# Submitting the search form using the updated method
search_box.submit()

def execute_script(script):
    while True:
        try:
            driver.execute_script(script)
            break
        except JavascriptException:
            sleep(0.5)

DOWNLOAD_DIR = 'c:/users/jjc/desktop/imgs'
필수옵션 = ['메시', '망사']
제외옵션 = '가죽'

# 더보기 클릭
# sleep(3)
execute_script("javascript:fnGoodsAttrNmFold('show');")

# 좌판재질 클릭
sleep(1)
execute_script("javascript:attrNmValLink('5611210201', '좌판재질', 'ATTR_264449' , '' ); ")
체크박스리스트 = driver.find_elements(By.CSS_SELECTOR, 'ul#dLstDiv>li>input[type="checkbox"]')
for 체크박스 in 체크박스리스트:
    parent = 체크박스.find_element(By.XPATH, './..')
    for j in 필수옵션:
        if j in parent.text and 제외옵션 not in parent.text:
            체크박스.click()
execute_script("javascript: toSMPPIntgrSrchGoodsList('')")

# 등판재질 클릭
sleep(1)
execute_script("javascript:attrNmValLink('5611210201', '등판재질', 'ATTR_269556', '' );")
체크박스리스트 = driver.find_elements(By.CSS_SELECTOR, 'ul#dLstDiv>li>input[type="checkbox"]')
for 체크박스 in 체크박스리스트:
    parent = 체크박스.find_element(By.XPATH, './..')
    for j in 필수옵션:
        if j in parent.text and '가죽' not in parent.text:
            체크박스.click()
execute_script("javascript: toSMPPIntgrSrchGoodsList('')")

# 머리바치파 유
sleep(1)
execute_script("javascript:attrNmValLink('5611210201', '머리받침판부착유무', 'ATTR_106171', '' );")
체크박스리스트 = driver.find_elements(By.CSS_SELECTOR, 'ul#dLstDiv>li>input[type="checkbox"]')
for 체크박스 in 체크박스리스트:
    parent = 체크박스.find_element(By.XPATH, './..')
    for j in 필수옵션:
        if j in parent.text and '가죽' not in parent.text:
            체크박스.click()
execute_script("javascript: toSMPPIntgrSrchGoodsList('')")

# 팔걸이 유
sleep(1)
execute_script("javascript:attrNmValLink('5611210201', '팔걸이유무', 'ATTR_259429', '' );")
체크박스리스트 = driver.find_elements(By.CSS_SELECTOR, 'ul#dLstDiv>li>input[type="checkbox"]')
for 체크박스 in 체크박스리스트:
    parent = 체크박스.find_element(By.XPATH, './..')
    for j in 필수옵션:
        if j in parent.text and '가죽' not in parent.text:
            체크박스.click()
execute_script("javascript:toSMPPIntgrSrchGoodsList('')")

# 100개 리스트
sleep(1)
driver.find_element_by_css_selector('option[value="100"]').click()
sleep(1)
execute_script("javascript:searchForNewPageSize();")

# 우선구매대상 + 의무구매대상
sleep(1)
driver.find_element(By.CSS_SELECTOR, "input[id='prior0bligPrdCrtfcCheck']").click()
sleep(1)
execute_script("javascript:toSMPPIntgrSrchGoodsList('');")

sleep(1)
아이템_리스트 = driver.find_elements(By.CSS_SELECTOR, "tbody>tr>td>a[href^='javascript:toSMPPGoodsDtlInfo(']")
스크립트_리스트 = []

for i in 아이템_리스트:
    script = i.get_attribute("href")
    스크립트_리스트.append(script)

filename = r"c:\users\jjc\desktop\chair_list.xlsx"
book = Workbook()
book.save(filename)

with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
    writer.book = book
    writer.sheets = {ws.title: ws for ws in book.worksheets}

    for idx, script in enumerate(스크립트_리스트):
        execute_script(script)
        print(idx)
        sleep(1)
        spec = pd.read_html(driver.page_source, index_col=0)[0].transpose()
        if idx == 0:
            spec.columns = map(lambda a: a.replace(" :", ''), spec.columns)
            spec.to_excel(writer, startrow=0, sheet_name='Sheet', index=False)
        else:
            spec.to_excel(writer, startrow=writer.sheets['Sheet'].max_row, sheet_name='Sheet', index=False, header=False)

        img = driver.find_element_by_css_selector('img[src"http://img.g2b.go.kr:7070/Resource/CataAttach/XezCatalog/XZMOK/"]').get_attribute('src')
        download(img, os.path.join(DOWNLOAD_DIR, f'{idx}.png'))
    writer.save()

def 클립보드로_이미지_복사하기(i):
    filepath = rf"C:\Users\smj02\Desktop\imgs\{i}.png"
    image = Image.open(filepath)
    output = BytesI0()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue() [14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

spec=pd.read_excel(r"C: \Users\smj02\Desktop\chair list.xlsx")
hwp = win32.gencache.EnsureDispatch("HWPFrame.Hwp0bject")
hwp.RegisterModule("FilePathCheckDLL", "securitymodule")
def 번호_붙이기(i):
    hwp.HAction.GetDefault("ComposeChars", hwp.HParameterSet.HChCompose.HSet)
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeLatin = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeHanja = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeJapanese = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontType0ther = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeSymbol = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeUser = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CircleType = 1
    hwp.HParameterSet.HChCompose.CheckCompose = 0
    hwp.HParameterSet.HChCompose.Chars = f"{i}"
    hwp.HAction.Execute("ComposeChars", hwp.HParameterSet.HChCompose.HSet)

for i in range(len(spec)):
    클립보드로_이미지_복사하기(i)
    hwp.Run('Paste')
    sleep(0.1)

    if i % 2 == 0:
        hwp.Run ("TableAppendRow")
    else:
        hwp.Run ("MoveDown")
    번호_붙이기(i + 1)

    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "    {}\r\n{}\r\n{}".format(spec.업체명[i].replace(" [중소기업]", ""),
                                                                        spec.규격명[i].splilt(',')[3].split()[0],
                                                                        spec.가격[i])
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    hwp.Run("TableRightCellAppend")
    if i % 2 == 0:
        hwp.Run ("MoveUp")
    else:
        pass




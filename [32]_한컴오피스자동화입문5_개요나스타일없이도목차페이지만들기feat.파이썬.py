import re
import os
import win32com.client as win32

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule') # 레지스트리에 등록한 보안모듈을 로드
hwp.Open(r"C:\Users\jjc\Desktop\한컴오피스 업무자동화 입문\첨부파일\개요없이목차.hwpx")

제목위치_리스트 = []

hwp.InitScan()
while True:
    text = hwp.GetText()
    if text[0] == 0: # [0]은 status, [1]은 content, 문서 끝에 도달(0)하면,
        hwp.ReleaseScan() # 문서스캔 종료
        break
    else: # 그 전까지는 1.2.3.가.나.더.힣.
        if re.match(r'[\d+가-힣]\.', text[1].strip()): # 숫자나 가나다라인 경우에 
            hwp.MovePos(201) # 해당 문단으로 이동
            제목위치_리스트.append(hwp.GetPos()) # 현재 위치를
        else:
            pass

for 제목위치 in 제목위치_리스트:
    hwp.SetPos(*제목위치) # *은 튜플 원소를 풀어서 삽입
    hwp.Run("MarkTitle") # [제목차례] 컨트를 삽입

hwp.MovePos(2) #= moveTopOfFile
hwp.Run('BreakPage') # 페이지 나누기
hwp.MovePos(2)

# 목차 생성을 위한 최소한의 파라미터.보다 자세한 설정은 parameterset에서 makecontents
hwp.HAction.GetDefault("MakeContents", hwp.HParameterSet.HMakeContents.HSet)
hwp.HParameterSet.HMakeContents.Make = 65 # 스타일 설정.
hwp.HParameterSet.HMakeContents.Leader = 6 # 탭채우기(점선, 파선, 실선 등이 있음)
hwp.HParameterSet.HMakeContents.type = 1
hwp.HAction.Execute("MakeContents", hwp.HParameterSet.HMakeContents.HSet)

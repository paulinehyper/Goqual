import os
from time import sleep
from tkinter import Tk
import tkinter
from tkinter.filedialog import askopenfilename
import re
import win32com.client as win32
import pyperclip as cb

#BASE_DIR = 'c:/hwp3'
#tkinter.Tk().withdraw()
#file_name = askopenfilename(initialdir=BASE_DIR)

#hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
#hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
#hwp.RegisterModule('FilePathCheckDLL','FileAuto') # 보안 승인창 뜨지 않도록 모듈 등록
#hwp.SetMessageBoxMode(0x00020000) # 메세지 창 뜨지 않도록 설정
##[출처] 파이썬으로 한글(hwp)내에 미주 개수 세기|작성자 코딩헤윰
#hwp.Open(os.path.join(BASE_DIR, file_name)) #파일열기
#current_window=hwp.XHwpWindows.Item(0)
#current_window.Visible=True

BASE_DIR = 'C:/Users/leeha.LAPTOP-FKRDOM42/Downloads/hwp'
def hwp_init2():
    
    tkinter.Tk().withdraw()
    file_name = askopenfilename(initialdir=BASE_DIR)
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    #hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.RegisterModule('FilePathCheckDLL','FileAuto') # 보안 승인창 뜨지 않도록 모듈 등록
    hwp.SetMessageBoxMode(0x00020000) # 메세지 창 뜨지 않도록 설정 ##[출처] 파이썬으로 한글(hwp)내에 미주 개수 세기|작성자 코딩헤윰
    hwp.Open(os.path.join(BASE_DIR, file_name)) #파일열기
    current_window=hwp.XHwpWindows.Item(0)
    current_window.Visible=True
    hwp.HAction.Run("MoveDocBegin") #문서 시작으로 커서 이동
    
    return hwp


#모든 빈줄 제거(미사용)
def check_empty_line(hwp):  
    dAct = hwp.CreateAction("AllReplace") ###빈줄 (^n을 찾아서 모두 ""아무것도 없는 것으로 replace하면 엔터가 사라짐)
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("FindString", "^n")
    dSet.SetItem("ReplaceString", "")
    dAct.Execute(dSet) 

    return hwp

#문단 속성( 문단위: 0, 문단아래: 0, 글)
def default_paragraph(hwp):
    hwp.HAction.Run("SelectAll")
    dAct = hwp.CreateAction("ParagraphShape") ## 1.문단 모양 바꾸기 시작.PageSetup
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("PrevSpacing", 0) 
    dSet.SetItem("NextSpacing", 0)
    dSet.SetItem("LineSpacingType", 0)
    dSet.SetItem("LineSpacing", 150)
    dSet.SetItem("LeftMargin", 0)
    dSet.SetItem("RightMargin", 0)
    dSet.SetItem("Indentation", 0.0)
    dAct.Execute(dSet)
    return hwp

def default_charShape(hwp):
    hwp.HAction.Run("SelectAll")
    dAct = hwp.CreateAction("CharShape") ## 1.문단 모양 바꾸기 시작.PageSetup
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    ##장평 : 기본값 100%
    dSet.SetItem("RatioUser", 100)
    dSet.SetItem("RatioSymbol", 100)
    dSet.SetItem("RatioOther", 100)
    dSet.SetItem("RatioJapanese", 100)
    dSet.SetItem("RatioHanja", 100)
    dSet.SetItem("RatioLatin", 100)
    dSet.SetItem("RatioHangul", 100)
    dAct.Execute(dSet)

    hwp.HAction.Run("SelectAll")
    dAct = hwp.CreateAction("CharShape")  
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    # 자간: 기본값 -5 
    dSet.SetItem("SpacingSymbol", -5)
    dSet.SetItem("SpacingJapanese", -5)
    dSet.SetItem("SpacingHanja", -5) 
    dSet.SetItem("SpacingLatin", -5)
    dSet.SetItem("SpacingHangul", -5)
    dSet.SetItem("SpacingUser", -5)
    dSet.SetItem("SpacingOther", -5)
    dAct.Execute(dSet)

    dAct = hwp.CreateAction("CharShape")
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("TextColor", 0xFF0000)
    dSet.SetItem("Bold",0)
    dSet.SetItem("FaceNameUser", "휴먼명조")
    dSet.SetItem("FontTypeUser", 1)
    dSet.SetItem("FaceNameSymbol", "휴먼명조")
    dSet.SetItem("FontTypeSymbol", 1)
    dSet.SetItem("FaceNameOther", "휴먼명조")
    dSet.SetItem("FontTypeOther", 1)
    dSet.SetItem("FaceNameJapanese", "휴먼명조")
    dSet.SetItem("FontTypeJapanese", 1)
    dSet.SetItem("FaceNameHanja", "휴먼명조")
    dSet.SetItem("FontTypeHanja", 1)
    dSet.SetItem("FaceNameLatin", "휴먼명조")
    dSet.SetItem("FontTypeLatin", 1)
    dSet.SetItem("FaceNameHangul", "휴먼명조")
    dSet.SetItem("FontTypeHangul", 1)
    dSet.SetItem("Height", 1500)
    dAct.Execute(dSet)

    
    return hwp


#인쇄 페이지 속성( 문단위: 0, 문단아래: 0, 글)
def pagesetup(hwp):
    dAct = hwp.CreateAction("PageSetup") ## 2.인쇄 페이지 속성 pageSetup
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("ApplyClass", 24)
    dSet.SetItem("ApplyTo", 3)
    ## 1mm = 283.465 HWPUNITs
    _dset = dSet.CreateItemSet ("PageDef", "PageDef")
    _dset.SetItem ("TopMargin", 2834.65)
    _dset.SetItem ("BottomMargin", 10)
    _dset.SetItem ("LeftMargin", 6236.23)
    _dset.SetItem ("RightMargin",6236.23)
    _dset.SetItem ("HeaderLen", 50)
    _dset.SetItem ("FooterLen", 10000)
    _dset.SetItem ("GutterLen", 0)

    dAct.Execute(dSet)

def hwp_center_align_and_insert_blank_line(hwp, dir, target):
    if target == "장":
        sleep(0.2)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        sleep(0.1)
    elif target =="□":
        #hwp.HAction.Run("ParagraphShapeAlignCenter")
        #hwp.HAction.Run("MoveSelListEnd")  ## SELECTION: 이전단어
        dAct = hwp.CreateAction("CharShape")
        dSet = dAct.CreateSet()
        dAct.GetDefault(dSet)
        dSet.SetItem("TextColor", 0xFF0000)
        dSet.SetItem("Bold",0)
        dSet.SetItem("FaceNameUser", "HY헤드라인M")
        dSet.SetItem("FontTypeUser", 1)
        dSet.SetItem("FaceNameSymbol", "HY헤드라인M")
        dSet.SetItem("FontTypeSymbol", 1)
        dSet.SetItem("FaceNameOther", "HY헤드라인M")
        dSet.SetItem("FontTypeOther", 1)
        dSet.SetItem("FaceNameJapanese", "HY헤드라인M")
        dSet.SetItem("FontTypeJapanese", 1)
        dSet.SetItem("FaceNameHanja", "HY헤드라인M")
        dSet.SetItem("FontTypeHanja", 1)
        dSet.SetItem("FaceNameLatin", "HY헤드라인M")
        dSet.SetItem("FontTypeLatin", 1)
        dSet.SetItem("FaceNameHangul", "HY헤드라인M")
        dSet.SetItem("FontTypeHangul", 1)
        dSet.SetItem("Height", 1500)
        dAct.Execute(dSet)
        


    else:
        pass

    if dir == "above":
        hwp.HAction.Run("MoveLineBegin")
        hwp.HAction.Run("BreakPara")
    elif dir == "below":
        hwp.HAction.Run("MoveLineEnd")
        hwp.HAction.Run("BreakPara")
        
        
        
    

    else:
        raise ValueError
    sleep(0.3)


def hwp_check_if_blank_exists_above(hwp):
    current_position = hwp.GetPos() #현위치 저장(깐혹 다음 검색위치로 튀는 문제 조치)
    hwp.HAction.Run("MoveLineBegin")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position) #방금위치 복원

    if cb.paste() == "r\n\r\n":
        return True
    else:
        return False
    

def hwp_check_if_blank_exists_below(hwp):
    current_position = hwp.GetPos()  # 현위치 저장(간혹 다음 검색위치로 튀는문제 조치)
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position)  # 방금위치 복원
    if cb.paste() == "\r\n\r\n":
        return True
    else:
        return False



   


#줄끝에서 몇칸 띄어서 그걸 복사한다음에, 그게 "\r\n\r\n"에 해당하면 1장 뒤에 엔터가 있다고 판단함. 
def hwp_check_if_blank_exists_below_then_delete(hwp): 
    current_position = hwp.GetPos() #현위치 저장(간혹, 다음 검색위치로 튀는 문제 조치)
    #print(current_position)
    #hwp.HAction.Run("MoveLineEnd")       ##띄어쓰기 포함 맨끝.
    #hwp.HAction.Run("MoveDocEnd")         ## 1. 문장 끝으로 이동.
    hwp.MovePos(3)
    dAct = hwp.CreateAction("InsertText")
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("Text", "MoveDocEnd")
    dAct.Execute(dSet)

    hwp.SetPos(*current_position)

    while True :
        
         ## hwp.GetPos()!=(0,129,28)   
        #current_position = hwp.GetPos()
        
        hwp.HAction.Run("MoveSelRight")
        #hwp.HAction.Run("MoveSelRight")
        hwp.HAction.Run("Copy")
        line_end_point = hwp.GetPos()
        print("hi")
        print(cb.paste())
      
        #hwp.SetPos(*current_position) #받금위치 복원
        if cb.paste() != "\r\n": ##뒤에 빈줄이 없고, 글자가 있으면 : 
            hwp.HAction.Run("MoveNextParaBegin")
            print("NextParaBegin")
            print(hwp.GetPos())
            if hwp.GetPos() == line_end_point:
                break
            else:
                hwp.HAction.Run("MoveLineEnd")
                line_end_point= hwp.GetPos()
                print("MoveLineEnd")
                print(hwp.GetPos())
        else:
            hwp.HAction.Run("Delete")
            print("delete")
            print(hwp.GetPos())
       
    hwp.ReleaseScan()
    hwp.MovePos(2)
    #hwp.MovePos(*(0, 32, 62))
    return True

def hwp_check_if_blank_exists_above_then_delete(hwp): 
    current_position = hwp.GetPos() #현위치 저장(간혹, 다음 검색위치로 튀는 문제 조치)
    #print(current_position)
    #hwp.HAction.Run("MoveLineEnd")       ##띄어쓰기 포함 맨끝.
    #hwp.HAction.Run("MoveDocEnd")         ## 1. 문장 끝으로 이동.
    hwp.MovePos(3)
    dAct = hwp.CreateAction("InsertText")
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("Text", "MoveDocEnd")
    dAct.Execute(dSet)

    hwp.SetPos(*current_position)

    while True :
        
         ## hwp.GetPos()!=(0,129,28)   
        #current_position = hwp.GetPos()
        
        hwp.HAction.Run("MoveSelRight")
        hwp.HAction.Run("MoveSelRight")
        hwp.HAction.Run("Copy")
        line_end_point = hwp.GetPos()
        print(cb.paste())
      
        #hwp.SetPos(*current_position) #받금위치 복원
        if cb.paste() != "\r\n\r\n": ##뒤에 빈줄이 없고, 글자가 있으면 : 
            hwp.HAction.Run("MoveNextParaBegin")
            print("NextParaBegin")
            print(hwp.GetPos())
            if hwp.GetPos() == line_end_point:
                break
            else:
                hwp.HAction.Run("MoveLineEnd")
                line_end_point= hwp.GetPos()
                print("MoveLineEnd")
                print(hwp.GetPos())
        else:
            hwp.HAction.Run("Delete")
            print("delete")
            print(hwp.GetPos())
       
    hwp.ReleaseScan()
    hwp.MovePos(2)
    return True


def hwp_char_headline(hwp, target):
    if target == "□":
        print(target)
        sleep(0.2)
        #hwp.HAction.Run("MoveSelRight") ## 빈줄이지만, (셀) 선택
        dAct = hwp.CreateAction("CharShape")
        dSet = dAct.CreateSet()
        dAct.GetDefault(dSet)
        dSet.SetItem("TextColor", 0xFF0000)
        dSet.SetItem("Bold",1)
        dSet.SetItem("FaceNameUser", "HY엽서M")
        dSet.SetItem("FontTypeUser", 1)
        dSet.SetItem("FaceNameSymbol", "HY엽서M")
        dSet.SetItem("FontTypeSymbol", 1)
        dSet.SetItem("FaceNameOther", "HY엽서M")
        dSet.SetItem("FontTypeOther", 1)
        dSet.SetItem("FaceNameJapanese", "HY엽서M")
        dSet.SetItem("FontTypeJapanese", 1)
        dSet.SetItem("FaceNameHanja", "HY엽서M")
        dSet.SetItem("FontTypeHanja", 1)
        dSet.SetItem("FaceNameLatin", "HY엽서M")
        dSet.SetItem("FontTypeLatin", 1)
        dSet.SetItem("FaceNameHangul", "HY엽서M")
        dSet.SetItem("FontTypeHangul", 1)
        dAct.Execute(dSet)
      
    else:
        pass


def insert_empty_line(hwp, level):

    dAct = hwp.CreateAction("CharShape")
    dSet = dAct.CreateSet()
    dAct.GetDefault(dSet)
    dSet.SetItem("FaceNameUser", "HY엽서M")
    dSet.SetItem("FontTypeUser", 1)
    dSet.SetItem("FaceNameSymbol", "HY엽서M")
    dSet.SetItem("FontTypeSymbol", 1)
    dSet.SetItem("FaceNameOther", "HY엽서M")
    dSet.SetItem("FontTypeOther", 1)
    dSet.SetItem("FaceNameJapanese", "HY엽서M")
    dSet.SetItem("FontTypeJapanese", 1)
    dSet.SetItem("FaceNameHanja", "HY엽서M")
    dSet.SetItem("FontTypeHanja", 1)
    dSet.SetItem("FaceNameLatin", "HY엽서M")
    dSet.SetItem("FontTypeLatin", 1)
    dSet.SetItem("FaceNameHangul", "HY엽서M")
    dSet.SetItem("FontTypeHangul", 1)

    if level ==1:
        dSet.SetItem("Height", 1500)
    elif level ==2:
        dSet.SetItem("Height", 1000)
    elif level==3:
        dSet.SetItem("Height", 800)


    
    
    dAct.Execute(dSet)

    hwp.HAction.Run("BreakPara")
    
    return hwp

def char_headline(hwp, level):
    if level == 1:
        sleep(0.2)
        dAct = hwp.CreateAction("CharShape")
        dSet = dAct.CreateSet()
        dAct.GetDefault(dSet)
        dSet.SetItem("TextColor", 0xFF0000)
        dSet.SetItem("Bold",1)
        dSet.SetItem("FaceNameUser", "HY울릉도M")
        dSet.SetItem("FontTypeUser", 1)
        dSet.SetItem("FaceNameSymbol", "HY울릉도M")
        dSet.SetItem("FontTypeSymbol", 1)
        dSet.SetItem("FaceNameOther", "HY울릉도M")
        dSet.SetItem("FontTypeOther", 1)
        dSet.SetItem("FaceNameJapanese", "HY울릉도M")
        dSet.SetItem("FontTypeJapanese", 1)
        dSet.SetItem("FaceNameHanja", "HY울릉도M")
        dSet.SetItem("FontTypeHanja", 1)
        dSet.SetItem("FaceNameLatin", "HY울릉도M")
        dSet.SetItem("FontTypeLatin", 1)
        dSet.SetItem("FaceNameHangul", "HY울릉도M")
        dSet.SetItem("FontTypeHangul", 1)
        dAct.Execute(dSet)

def hwp_find_and_go(hwp):
    scan_position = (0,0,0)    
    hwp.InitScan()
    장번호 = 1
    조번호 = 1
  ###"□ ("
    while True :
        hwp.SetPos(*scan_position)
        text = hwp.GetText()
        if text[0] ==1:
            print(text[0])
            break
        else:
            if re.match(rf"^제{장번호}장.+", text[1].strip().replace(" ", "")):
                print("print text[0]=")
                print(text[0])
                장번호 +=1
                hwp.MovePos(201) #moveScanPos : GetText()실행 후 위치로 이동한다
                dAct = hwp.CreateAction("InsertText")
                dSet = dAct.CreateSet()
                dAct.GetDefault(dSet)
                dSet.SetItem("Text", "HEYJEY")
                dAct.Execute(dSet)
                
                hwp.MovePos(20) ## 한줄 아래로 이동. 2021.08.26
                sleep(0.2)

                if not hwp_check_if_blank_exists_above(hwp):
                    hwp_center_align_and_insert_blank_line(hwp, "above", "장")
                
                if not hwp_check_if_blank_exists_below(hwp):
                    hwp_center_align_and_insert_blank_line(hwp, "below", "장")
                hwp.InitScan()

            pattern = re.escape("□" + "(")

            if re.match(pattern, text[1].strip().replace(" ", "")): ##앞단의 빈칸을 정리해줘야 함.
                #print("print text[0]=")
                #print(text[1])
                #print("처음")
                #print(hwp.GetPos())
                current_position=hwp.GetPos()
                hwp.MovePos(201) #moveScanPos : GetText()실행 후 위치로 이동한다
                #dAct = hwp.CreateAction("InsertText")
                #dSet = dAct.CreateSet()
                #dAct.GetDefault(dSet)
                #dSet.SetItem("Text", "메롱")
                #dAct.Execute(dSet)
                
                position_end = text[1].find(')') +1 #)의 위치 POS
                #print(position_end)
                movetoend=(hwp.GetPos()[0], hwp.GetPos()[1], hwp.GetPos()[2]+position_end)
                hwp.SetPos(*movetoend)
                #print(hwp.GetPos())
                #hwp.MovePos(16)
                #print(hwp.GetPos())
                #hwp.MovePos(16)
                #print(hwp.GetPos())
                #hwp.MovePos(16)
                #print(hwp.GetPos())

                
               
                #dAct = hwp.CreateAction("InsertText")
                #dSet = dAct.CreateSet()
                #dAct.GetDefault(dSet)
                #dSet.SetItem("Text", "짱")
                #dAct.Execute(dSet)

                hwp.HAction.Run("MoveSelPrevParaBegin")

                char_headline(hwp, 1)
                
                hwp.MovePos(20) ## 한줄 아래로 이동. 2021.08.26
                hwp.MovePos(23) ## 한줄 아래로 이동. 2021.08.26
                
                #dAct = hwp.CreateAction("InsertText")
                #dSet = dAct.CreateSet()
                #dAct.GetDefault(dSet)
                #dSet.SetItem("Text", "여기")
                #dAct.Execute(dSet)
                sleep(0.2)
                insert_empty_line(hwp, 1) ##level ==1, 15pt
                scan_position = hwp.GetPos()

                #if not hwp_check_if_blank_exists_above(hwp):
                #    hwp_center_align_and_insert_blank_line(hwp, "above", "조")
                

                #if not hwp_check_if_blank_exists_below(hwp):
                #    hwp_center_align_and_insert_blank_line(hwp, "below", "조")
                #hwp.InitScan() 이걸 제거해야, 무한순환이 없어진다.
            
            
                

                

              
           
            else:
                pass

    hwp.ReleaseScan()
    hwp.MovePos(2)




## 칸을 한칸이라도 앞에 넣으면, main 함수 실행이 안된다.ㅡ.ㅡ 20218.08.15

if __name__ == '__main__': 
    #root = Tk()     
    #filename = askopenfilename()
    #root.destroy()
    #hwp = hwp_init(filename = filename)
    hwp2 = hwp_init2()
    #check_empty_line(hwp2)
    #delete_empty_line(hwp2)
    #hwp_center_align_and_insert_blank_line(hwp2, "below","장")
    #hwp_char_headline(hwp2, "□(")
    #default_preference(hwp2)
    #pagesetup(hwp2)
    
    #hwp_check_if_blank_exists_below_then_delete(hwp2)
    #default_paragraph(hwp2) ##문단 위/아래 0포인트로 맞추기. ##줄간격 160%
    #default_charShape(hwp2) ##글 자간 설정
   
    hwp_find_and_go(hwp2)
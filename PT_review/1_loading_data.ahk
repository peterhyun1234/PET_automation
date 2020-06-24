; 1. gui 위치 조정
; 2. 몇개가져올 지 파악 ㄱㄱ
; 3. 





currentId := ""
currentApprover := ""
currentCI := ""
currentAllComments := ""






Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text, x21 y34 w100 h19 ,Progress 
Gui, Add, Edit, x75 y30 w300 h19 vProcess,
Gui, Add, Text, x21 y74 w100 h19 ,Contents 
Gui, Add, Edit, x75 y70 w400 h300 vMyEdit, 빈칸
Gui, Add, Button, x220 y400 w100 h30 gLoad_data, load reviews
Gui, Show, x1000 y800


Load_data:

global currentId
global currentApprover
global currentCI
global currentAllComments

Gui,Submit,nohide


temp_window := StrSplit(MyEdit, "***********************************************************************************", ".")

; 1. 일단 ************ 기준으로 나누기
AllComments := temp_window[1]
attributes := temp_window[2]

; 2. allcomments에서 CI 뽑아내기
temp_CI := StrSplit(AllComments, "(Clinical information:", ".")
tCI := StrSplit(temp_CI[2], ")", ".")

currentAllComments := AllComments
currentCI := tCI[1]

;MsgBox, 4, , % currentAllComments , 3
;MsgBox, 4, , % currentCI , 3


; 3. attributes에서 ID, Approver 뽑아내기
;    - 이때 이전 ID랑 같은지 검사해서 같으면 파일의 끝으로 인식 ㄱㄱ
temp_attributes := StrSplit(attributes, "Patient Name / ID :", ".")
tattributes := StrSplit(temp_attributes[2], "/ ", ".")
tempID := StrSplit(tattributes[2], "`n", ".")

; MsgBox, 4, , % tempID[1] , 3
currentId := tempID[1]

temp_attributes := StrSplit(attributes, "Approver : ", ".")
temp_approver := StrSplit(temp_attributes[2], "`n", ".")

; MsgBox, 4, , % temp_approver[1] , 3

currentApprover := temp_approver[1]


return



^ESC::
	ExitApp

GuiClose:
	ExitApp
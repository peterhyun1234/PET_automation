; 1. gui 위치 조정
; 2. 몇개가져올 지 파악 ㄱㄱ
; 3. 


xpos := ""
ypos := ""


currentId := "sample"
currentApprover := ""
currentCI := ""
currentAllComments := ""

duplicateFlag := false

currentIdx = 0


reviewFileName = 202005\202005_reviews.xlsx


Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text, x21 y34 w100 h19 ,Date 
Gui, Add, Edit, x75 y30 w300 h19 vStudy_Date,
Gui, Add, Text, x21 y74 w100 h19 ,Contents 
Gui, Add, Edit, x75 y70 w400 h300 vMyEdit, 빈칸
Gui, Add, Button, x220 y400 w100 h30 gLoad_data, load reviews
Gui, Show, x1000 y800




; init exel file
; Indexpath:= "F:\Nuclear Medicine\판독문 정리\process_1\202005\202005_reviews.xlsx"
; IndexExcel := ComObjCreate("Excel.Application") ;오브젝트생성
; IndexExcel.Workbooks.Open(Indexpath) ;엑셀열기
; IndexExcel.Visible:= true  ;false     ;육안으로 보이게 할 지 설정

; IndexExcel.ActiveWorkbook.Save() ;저장
; IndexExcel.ActiveWorkBook.Close  ;닫기
; IndexExcel.Quit ;오브젝트종료











Load_data:

	global currentId
	global currentApprover
	global currentCI
	global currentAllComments

	global duplicateFlag

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
	;    - 이때 이전 ID랑 같은지 검사해서 같으면 파일의 끝으로 인식 ㄱㄱ by using duplicateFlag
	temp_attributes := StrSplit(attributes, "Patient Name / ID :", ".")
	tattributes := StrSplit(temp_attributes[2], "/ ", ".")
	tempID := StrSplit(tattributes[2], "`n", ".")

	if (currentId = tempID[1]){
		duplicateFlag := true
	}else{
		; MsgBox, 4, , % tempID[1] , 3
		currentId := tempID[1]
	}

	temp_attributes := StrSplit(attributes, "Approver : ", ".")
	temp_approver := StrSplit(temp_attributes[2], "`n", ".")

	; MsgBox, 4, , % temp_approver[1] , 3

	currentApprover := temp_approver[1]


	; bring EOF to idx

	if (currentAllComments = "빈칸"){
		MsgBox, 4, , % "currentAllComments = 빈칸" , 1
	}else{

		Indexpath:= "F:\Nuclear Medicine\판독문 정리\process_1\202005\202005_reviews.xlsx"
		IndexExcel := ComObjCreate("Excel.Application") ;오브젝트생성
		IndexExcel.Workbooks.Open(Indexpath) ;엑셀열기
		IndexExcel.Visible:= false  ;true     ;육안으로 보이게 할 지 설정

		; currentIdx := 
		LastRow := IndexExcel.Cells.SpecialCells(11).Row

		; MsgBox, 4, , % "lastrow is " LastRow , 3

		IndexExcel.Cells(LastRow + 1,1).Value := currentId
		IndexExcel.Cells(LastRow + 1,2).Value := currentApprover
		IndexExcel.Cells(LastRow + 1,3).Value := currentCI
		IndexExcel.Cells(LastRow + 1,4).Value := currentAllComments


		IndexExcel.ActiveWorkbook.Save() ;저장
		IndexExcel.ActiveWorkBook.Close  ;닫기
		IndexExcel.Quit ;오브젝트종료

		MsgBox, 4, , % "ID: " currentId " completed"  , 1
	}

	; initialize values
	currentId := ""
	currentApprover := ""
	currentCI := ""
	currentAllComments := ""

return


^1:: ; 판독문 따오기 세팅

	studyTerm = 10

	MouseClick, left, 1100, 864
	Send, ^a
	Send, ^c

	MouseClick, left, -332, -621
	Send, ^v

	MouseClick, left, 833, 180
	Send, PT

	MouseClick, left, 1498, 160

	MouseClick, left, 1318, 468

	; 스크롤 10번 	
	while(studyTerm != 0){
        MouseClick, WheelDown, , , 50
        Sleep, 2000
        studyTerm --
    }

	MouseClick, left, 1095, 208
	Sleep, 1000
	MouseClick

	MouseClick, left, 1400, 208
	Sleep, 1000
	MouseClick

	MsgBox, 4, , "Setting End" , 2

return





^2:: ; 판독문 따오기 시작

	Coordmode,Mouse,Screen

	Loop, 10
	{
		if duplicateFlag
		{
			MsgBox, 4, , % "Duplication occured in ID, Process end" , 2
			break
		}else{
			
			MouseGetPos, xpos, ypos ; store position

			;MsgBox, 4, , % "ypos: " ypos ", xpos: " xpos, 2

			MouseClick, left, 200, 1240

			Send {WheelDown}
			Sleep, 200
			Send {WheelDown}
			Sleep, 200
			Send {WheelDown}
			Sleep, 200
			Send {WheelDown}
			Sleep, 200
			Send {WheelDown}
			Sleep, 200


			; shift 누른 상태에서
			SENDINPUT {SHIFT DOWN}

			
			; 스크롤 내리고

			MouseClick, left, 562, 1958

			SENDINPUT {SHIFT UP}

			Send, ^c

			MouseClick, left, 1150, 900
			Send, ^a
			Send, ^v

			Sleep, 200


			ypos := ypos + 20

			;;여기서 듀플리케이트 걷어내야함!!!!

			if (ypos > 752) ; end of scroll
			{
				MouseClick, left, %xpos%, %ypos%
				Sleep, 2000
				MouseClick
				Sleep, 2000
				
				Loop, 24
				{
					Send, {Down}
					Sleep, 1000
				}

				ypos := 230
				
				MouseClick, left, %xpos%, %ypos% ; return to origin of mouse position + y
			}else{
				MouseClick, left, %xpos%, %ypos% ; return to origin of mouse position
			}



			Sleep, 500
		}
	}

	
return




^ESC::

	ExitApp

GuiClose:
	
	ExitApp
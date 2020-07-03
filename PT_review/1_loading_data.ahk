; 1. gui ��ġ ����
; 2. ������� �� �ľ� ����
; 3. 


xpos := ""
ypos := ""


currentId := "sample"
currentApprover := ""
currentCI := ""
currentAllComments := ""

endFlag := false
inputEndID := ""
currentIdx = 0


reviewFileName = 202005\202005_reviews.xlsx


Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text, x21 y34 w100 h19 ,Date 
Gui, Add, Edit, x75 y30 w300 h19 vStudy_Date,
Gui, Add, Text, x21 y74 w100 h19 ,last ID 
Gui, Add, Edit, x75 y70 w300 h19 vlastID,
Gui, Add, Text, x21 y114 w100 h19 ,Contents 
Gui, Add, Edit, x75 y110 w400 h300 vMyEdit, ��ĭ
Gui, Add, Button, x220 y440 w100 h30 gLoad_data, load reviews
Gui, Show, x1000 y840




Load_data:

	global currentId
	global currentApprover
	global currentCI
	global currentAllComments

	global endFlag

	global inputEndID

	Gui,Submit,nohide

	inputEndID := lastID

	temp_window := StrSplit(MyEdit, "***********************************************************************************", ".")

	; 1. �ϴ� ************ �������� ������
	AllComments := temp_window[1]
	attributes := temp_window[2]

	; 2. allcomments���� CI �̾Ƴ���
	temp_CI := StrSplit(AllComments, "(Clinical information:", ".")
	tCI := StrSplit(temp_CI[2], ")", ".")

	currentAllComments := AllComments
	currentCI := tCI[1]

	;MsgBox, 4, , % currentAllComments , 3
	;MsgBox, 4, , % currentCI , 3


	; 3. attributes���� ID, Approver �̾Ƴ���
	;    - �̶� ���� ID�� ������ �˻��ؼ� ������ ������ ������ �ν� ���� by using endFlag
	temp_attributes := StrSplit(attributes, "Patient Name / ID :", ".")
	tattributes := StrSplit(temp_attributes[2], "/ ", ".")
	tempID := StrSplit(tattributes[2], "`n", ".")
	
	MsgBox, 4, , % "inputEndID: " inputEndID ", tempID: " tempID[1], 2

	if (tempID[1] = inputEndID){
		MsgBox, 4, , % "inputEndID: " inputEndID ", tempID: " tempID[1] " endFlag changed", 3
		endFlag := true
	}else{
		; MsgBox, 4, , % tempID[1] , 3
		currentId := tempID[1]
	}

	temp_attributes := StrSplit(attributes, "Approver : ", ".")
	temp_approver := StrSplit(temp_attributes[2], "`n", ".")

	; MsgBox, 4, , % temp_approver[1] , 3

	currentApprover := temp_approver[1]


	; bring EOF to idx

	if (currentAllComments = "��ĭ"){
		MsgBox, 4, , % "currentAllComments = ��ĭ" , 1
	}else{

		Indexpath:= "F:\Nuclear Medicine\�ǵ��� ����\process_1\202005\202005_reviews.xlsx"
		IndexExcel := ComObjCreate("Excel.Application") ;������Ʈ����
		IndexExcel.Workbooks.Open(Indexpath) ;��������
		IndexExcel.Visible:= false  ;true     ;�������� ���̰� �� �� ����

		; currentIdx := 
		LastRow := IndexExcel.Cells.SpecialCells(11).Row

		; MsgBox, 4, , % "lastrow is " LastRow , 3

		IndexExcel.Cells(LastRow + 1,1).Value := currentId
		IndexExcel.Cells(LastRow + 1,2).Value := currentApprover
		IndexExcel.Cells(LastRow + 1,3).Value := currentCI
		IndexExcel.Cells(LastRow + 1,4).Value := currentAllComments


		IndexExcel.ActiveWorkbook.Save() ;����
		IndexExcel.ActiveWorkBook.Close  ;�ݱ�
		IndexExcel.Quit ;������Ʈ����

		MsgBox, 4, , % "ID: " currentId " completed"  , 1
	}

	; initialize values
	currentId := ""
	currentApprover := ""
	currentCI := ""
	currentAllComments := ""

return


^1:: ; �ǵ��� ������ ����

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

	; ��ũ�� 10�� 	
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





^2:: ; �ǵ��� ������ ����

	Coordmode,Mouse,Screen

	Loop, 100
	{
		if endFlag = true
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


			; shift ���� ���¿���
			SENDINPUT {SHIFT DOWN}

			
			; ��ũ�� ������

			MouseClick, left, 562, 1958

			SENDINPUT {SHIFT UP}

			Send, ^c

			MouseClick, left, 1154, 1050
			Send, ^a
			Send, ^v

			Sleep, 1000

			; load reviews btn
			MouseClick, left, 1269, 1320

			Sleep, 1000

			ypos := ypos + 20

			;;���⼭ ���ø�����Ʈ �Ⱦ����!!!!

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
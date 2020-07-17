; 1. gui ��ġ ����
; 2. ������� �� �ľ� ����
; 3. 


xpos := ""
ypos := ""


currentId := "sample"
currentApprover := ""
currentCI := ""
currentAllComments := ""

progress := 0
endFlag := 0
inputEndID := ""
currentIdx = 0


reviewFileName = F:\Nuclear Medicine\�ǵ��� ����\process_1\202004_reviews_1.txt

Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text, x21 y34 w100 h19 ,Date 
Gui, Add, Edit, x75 y30 w300 h19 vStudy_Date,
Gui, Add, Text, x21 y74 w100 h19 ,lastID-1
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
	
	; MsgBox, 4, , % "inputEndID: " inputEndID ", tempID: " tempID[1], 2

	if (tempID[1] = inputEndID){

		if (progress != 0){
			endFlag := 1
			MsgBox, 4, , % "inputEndID: " inputEndID ", tempID: " tempID[1] " endFlag changed", 3
			currentId := tempID[1]
		}
		
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

		FileAppend, ID: %currentId% `nApprover: %currentApprover% `nCI: %currentCI% `nAllComments: %currentAllComments%///`n, %reviewFileName%


		; Indexpath:= "F:\Nuclear Medicine\�ǵ��� ����\process_1\202005\test2.xlsx"
		; IndexExcel := ComObjCreate("Excel.Application") ;������Ʈ����
		; IndexExcel.Workbooks.Open(Indexpath) ;��������
		; IndexExcel.Visible:= false  ;true     ;�������� ���̰� �� �� ����

		; ; currentIdx := 
		; ; LastRow := IndexExcel.Cells.SpecialCells(11).Row
		; LastRow := IndexExcel.Sheets(1).UsedRange.Rows.Count

		; MsgBox, 4, , % "lastrow is " LastRow , 2

		; IndexExcel.Cells(LastRow + 1,1).Value := currentId
		; IndexExcel.Cells(LastRow + 1,2).Value := currentApprover
		; IndexExcel.Cells(LastRow + 1,3).Value := currentCI
		; IndexExcel.Cells(LastRow + 1,4).Value := currentAllComments


		; IndexExcel.ActiveWorkbook.Save() ;����
		; IndexExcel.ActiveWorkBook.Close  ;�ݱ�
		; IndexExcel.Quit ;������Ʈ����

		; MsgBox, 4, , % "ID: " currentId " completed"  , 1
		progress = progress + 1
	}

	; initialize values
	currentId := ""
	currentApprover := ""
	currentCI := ""
	currentAllComments := ""

return


^1:: ; �ǵ��� ������ ����

	Coordmode,Mouse,Screen

	studyTerm = 10

	MouseClick, left, 1226, 900
	Send, ^a
	Send, ^c

	Sleep, 200

	MouseClick, left, 727, 181
	Send, ^v

	Sleep, 200

	MouseClick, left, 833, 181
	Send, PT

	Sleep, 200

	MouseClick, left, 1500, 150
	
	Sleep, 200

	MouseClick, left, 1318, 468

	Sleep, 200

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

	Loop
	{

		MouseGetPos, xpos, ypos ; store position

		;MsgBox, 4, , % "ypos: " ypos ", xpos: " xpos, 2

		MouseClick, left, 200, 1240

		Send {WheelDown}
		Sleep, 400
		Send {WheelDown}
		Sleep, 400
		Send {WheelDown}
		Sleep, 400
		Send {WheelDown}
		Sleep, 400
		Send {WheelDown}
		Sleep, 400


		; shift ���� ���¿���
		SENDINPUT {SHIFT DOWN}
		Sleep, 400
		
		
		; ��ũ�� ������

		MouseClick, left, 562, 1958
		Sleep, 400

		SENDINPUT {SHIFT UP}
		Sleep, 400

		Send, ^c
		Sleep, 400
		
		MouseClick, left, 1154, 1050
		Send, ^a
		Send, ^v

		Sleep, 400

		; load reviews btn
		MouseClick, left, 1269, 1320

		Sleep, 400

		ypos := ypos + 22

		;;���⼭ ���ø�����Ʈ �Ⱦ����!!!!

		if (ypos > 752) ; end of scroll
		{
			MouseClick, left, %xpos%, %ypos%

			if endFlag = 1
			{
				MouseClick, left, 200, 1240

				Send {WheelDown}
				Sleep, 400
				Send {WheelDown}
				Sleep, 400
				Send {WheelDown}
				Sleep, 400
				Send {WheelDown}
				Sleep, 400
				Send {WheelDown}
				Sleep, 400


				; shift ���� ���¿���
				SENDINPUT {SHIFT DOWN}
				Sleep, 400
				
				
				; ��ũ�� ������

				MouseClick, left, 562, 1958
				Sleep, 400

				SENDINPUT {SHIFT UP}
				Sleep, 400

				Send, ^c
				Sleep, 400
				
				MouseClick, left, 1154, 1050
				Send, ^a
				Send, ^v

				Sleep, 400

				; load reviews btn
				MouseClick, left, 1269, 1320

				Sleep, 400

				MsgBox, 4, , % "Duplication occured in ID, Process end" , 2

				break

			}else{

				Sleep, 2000
				MouseClick
				Sleep, 2000
				
				Loop, 24
				{
					Send, {Down}
					Sleep, 1500
				}

				ypos := 230
				
				MouseClick, left, %xpos%, %ypos% ; return to origin of mouse position + y
			}

		}else{
			MouseClick, left, %xpos%, %ypos% ; return to origin of mouse position
		}
	}


MsgBox, 4, , % "Parsing Process end" , 2

	
return



^ESC::

	ExitApp

GuiClose:
	
	ExitApp
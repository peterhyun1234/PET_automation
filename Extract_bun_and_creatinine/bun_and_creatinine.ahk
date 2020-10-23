idIdx = 1
idArray:= Object()
idArrayLen = 0
currentID = 0

currentReadLine := ""
bun := ""
creatinine := ""

idTxtName = Infi\input_id.txt
destTxtName = Infi\output_BAC.txt

Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Checking Sharness 
Gui, Add, Edit, x11 y45 w500 h19 vName, Edit
Gui, Add, Text, x11 y74 w100 h19 ,Progress: 
Gui, Add, Edit, x75 y70 w300 h19 vProcess,
Gui, Add, Button, x50 y150 w250 h30 gPush_bun_and_creatinine, read bun_and_creatinine
Gui, Show, x1635 y0

Push_bun_and_creatinine:

	global bun
	global creatinine

	Gui,Submit,nohide

	bun_and_creatinine := StrSplit(Name, A_Space, ".")

	bun := bun_and_creatinine[1]
	creatinine := bun_and_creatinine[2]
	
    MsgBox, 4, , % "bun: " bun ", creatinine: " creatinine , 3

return


^r:: ; 1. load IDs in file data

	Loop, Read, C:\Users\ajounm\Desktop\hb_workspace\PET_automation\Extract_bun_and_creatinine\data\%idTxtName%
	{
		idArray.insert(A_LoopReadLine)
		idArrayLen++
	}

	Msgbox, "Reading proccess complete"

return 


a:: ; 2. when press some keys, then Search Id and


	currentID := idArray[idIdx]


	if !currentID 
	{
		Msgbox, "currentID empty"
		return
	}
		
	Mousemove, 465, 176
	Sleep, 100

	; remove
	MouseClick
	Sleep, 100
	Send, {BS}

	; insert
	MouseClick
	Sleep, 100
	Send, %currentID%
	Sleep, 100
	Send, {Enter}
	Sleep, 100

	Mousemove, 30, 236
	MouseClick

	idIdx ++

return



s::

	Mousemove, 1612, 592
	MouseClick

return



d:: ; Start Settings and Write bun
	
	Mousemove, 18, 780
	MouseClick

	Mousemove, 248, 618
	MouseClick, ,,, 2 ; double click
	Send, 50
	Send, {Enter}

	MouseClick, left, 1750, 77 ; temp Window
	Send, ^a
	Send, {BS}
	Sleep, 50


return 


f:: ; write bun for each ID

	push_bun()

	write_data()

	initialize_count()

	initialize_progress()

	close_reviews()

return 



StartSetting(X, Y, Zoom, Bright){

	Mousemove, 18, 780
	MouseClick

	Mousemove, 46, 662
	MouseClick
		
	Send, %X%
	Send, {Enter}
	
	Mousemove, 49, 775
	MouseClick
	Send, {BS}%Zoom%
	Send, {Enter}

	Mousemove, 248, 618
	MouseClick, ,,, 2 ; double click
	Send, %Bright%
	Send, {Enter}
}


push_bun(){
	Coordmode, Mouse, Screen

	MouseClick, left, 3400, 185 ; bun push button

	Sleep, 50

	MouseClick, left, 3060, 30 ; return to xeleris process

}


write_data(){

	global destTxtName 
	global currentID
	global bun

	FileAppend, %currentID% %bun%`n, C:\Users\ajounm\Desktop\hb_workspace\PET_automation\Extract_bun_and_creatinine\data\%destTxtName%
}


initialize_count(){

	global bun 
	bun = 0
}

initialize_progress(){

	Coordmode, Mouse, Screen
	global idIdx
	global idArrayLen

	MouseClick, left, 3442, 100 ; progress
	Send, ^a
	Send, {BS}

	Send, %idIdx%
	Send, {Space}
	Send, /
	Send, {Space}
	Send, %idArrayLen%
	
	Sleep, 50
}


close_reviews(){

	Coordmode, Mouse, Screen
	MouseClick, left, 2182, 123

	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send, {Enter}

	Sleep, 100

	MouseClick, left, 1000, 125
}



^ESC::
	ExitApp

GuiClose:
	ExitApp
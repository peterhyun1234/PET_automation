idIdx = 1
idArray:= Object()
idArrayLen = 0
currentID = 0

currentReadLine := ""
sharpness := ""

idTxtName = 201912_202007\201912_202007_ID.txt
destTxtName = 201912_202007\201912_202007_result.txt

Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Checking Sharness 
Gui, Add, Edit, x21 y45 w500 h19 vName, Edit
Gui, Add, Text, x21 y74 w100 h19 ,Progress: 
Gui, Add, Edit, x75 y70 w300 h19 vProcess,
Gui, Add, Button, x50 y150 w100 h30 gPush_sharpness, read sharpness
Gui, Show, x3300 y0

Push_sharpness:

	global sharpness

	Gui,Submit,nohide

	sharpness := Name

	; MsgBox, 4, , % "sharpness: " sharpness , 3

return


^r:: ; 1. load IDs in file data

	Loop, Read, C:\Program Files\AutoHotkey\1_data_mining\result\%idTxtName%
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

	if currentID = #Infi
	{
		;Msgbox, %currentID%
		idIdx ++
		write_machine()
		return
	}

	if currentID = #Sym
	{
		;Msgbox, %currentID%
		idIdx ++
		write_machine()
		return
	}

	if currentID = #NMCT670
	{
		;Msgbox, %currentID%
		idIdx ++
		write_machine()
		return
	}

	if currentID = #NM830
	{
		;Msgbox, %currentID%
		idIdx ++
		write_machine()
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



d:: ; Start Settings and Write sharpness
	
	Mousemove, 18, 780
	MouseClick

	Mousemove, 248, 618
	MouseClick, ,,, 2 ; double click
	Send, 28
	Send, {Enter}

	MouseClick, left, 1750, 77 ; temp Window
	Send, ^a
	Send, {BS}
	Sleep, 50


return 


f:: ; write sharpness for each ID

	push_sharpness()

	write_statistics()

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


push_sharpness(){
	Coordmode, Mouse, Screen

	MouseClick, left, 3400, 185 ; sharpness push button

	Sleep, 50

	MouseClick, left, 3060, 30 ; return to xeleris process

}

write_machine(){

	global destTxtName 
	global currentID

	FileAppend, %currentID%`n, C:\Program Files\AutoHotkey\1_data_mining\result\%destTxtName%
}

write_statistics(){

	global destTxtName 
	global currentID
	global sharpness

	FileAppend, %currentID% %sharpness%`n, C:\Program Files\AutoHotkey\1_data_mining\result\%destTxtName%
}


initialize_count(){

	global sharpness 
	sharpness = 0
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
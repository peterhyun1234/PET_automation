; 1. load IDs in file data
; 2. when press some keys, then Search Id and
; 3. Open WB file in review (not macro)
; 4. Decide that img can be used (not macro)
; 5. if can, press each leg macro key
; 6. else, enter the reason why can't (not macro)
; 7. process 2 restart



stopFlag = 0
reIndicateFlag = 0
sideOfleg = ""

idIdx = 1
idArray:= Object()
idArrayLen = 0
currentID = 0

currentReadLine := ""

idTxtName = 201912_202007\201912_202007_ID.txt
destTxtName = 201912_202007\201912_202007_result.txt

xpos = 0
ypos = 0

Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Temp window %idIdx%
Gui, Add, Edit, x21 y45 w500 h19 vName, Edit
Gui, Add, Text, x21 y74 w100 h19 ,Progress: 
Gui, Add, Edit, x75 y70 w300 h19 vProcess,
Gui, Add, Button, x50 y150 w100 h30 gPush_sharpness, read sharpness
Gui, Show, x3300 y0


Push_sharpness:

	global sharpness

	global stopFlag

	Gui,Submit,nohide

	sharpness = Name

	MsgBox, 4, , % "sharpness: " sharpness , 3

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



d:: ; Start Settings

	X = 0
	Y = 0
	Zoom = 1
	Bright = 28

	StartSetting(X, Y, Zoom, Bright)

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

	global xpos
	global ypos

	MouseClick, left, 1482, 189 ; femur push button

	Sleep, 50

	MouseClick, left, -900, 175

	Sleep, 50
	
	MouseClick, left, %xpos%, %ypos% ; return to origin of mouse position

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
	
	global idIdx
	global idArrayLen

	MouseClick, left, 1750, 100 ; temp Window
	Send, ^a
	Send, {BS}

	Send, %idIdx%
	Send, {Space}
	Send, /
	Send, {Space}
	Send, %idArrayLen%
	
	Sleep, 50
	Mousemove, -900, 175
	MouseClick

}


close_reviews(){

	MouseClick, left, 262, 125

	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send, {Enter}

	Sleep, 100

	MouseClick, left, -740, 50
}



^ESC::
	ExitApp

GuiClose:
	ExitApp
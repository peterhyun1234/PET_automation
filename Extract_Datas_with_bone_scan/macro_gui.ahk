; 1. load IDs in file data
; 2. when press some keys, then Search Id and
; 3. Open WB file in review (not macro)
; 4. Decide that img can be used (not macro)
; 5. if can, press each leg macro key
; 6. else, enter the reason why can't (not macro)
; 7. process 2 restart



idIdx = 1
idArray:= Object()
currentID = 0

currentReadLine := ""

clear := ""

idTxtName = test.txt

destTxtName = march_result.txt



Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Temp window
Gui, Add, Edit, x21 y45 w500 h19 vName, Edit
Gui, Add, Button, x150 y150 w100 h30 gPush, please push
Gui, Add, Button, x350 y150 w100 h30 gClear, Clear
Gui, Show



Push:
Gui,Submit,nohide

Msgbox, %Name%

return




Clear:
Guicontrol,, Edit, %clear%
return



^r:: ; 1. load IDs in file data

Loop, Read, C:\Users\Administrator\Desktop\autohotkey\AutoHotkey\result_of\%idTxtName%
{
	idArray.insert(A_LoopReadLine)
}

Msgbox, "Reading proccess complete"

;For index, element in idArray
;{
;	Msgbox, % "Element number" . index . " is " . element
;}
return 


g:: ; 2. when press some keys, then Search Id and

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

idIdx ++

return



a:: ; Start Settings with Left leg

X = -4
Y = 0
Zoom = 4
Bright = 28

StartSetting(X, Y, Zoom, Bright)

; Create ROI with Left leg

sideOfleg = left

Create_ROI(sideOfleg)

open_statistics()

return 






s:: ; 5. if can, press each leg macro key

write_temp_femur()

Sleep, 200

;move_to_thigh()

Sleep, 200

;write_temp_thigh()

Sleep, 200

;write_statistics()

Sleep, 200

return 





d:: ; move ROI to right with Left leg

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , 5, 0, ,R

return 





f:: ; Switch Left leg to Right leg

Mousemove, 18, 780
Sleep, 200
MouseClick
Sleep, 100

Mousemove, 50, 662
Sleep, 100
MouseClick
Sleep, 100
Send, {BS}
Sleep, 100
Send, {Enter}
Sleep, 100	

return 






q:: ; delete Cards
cardNum = 8
closeCards(cardNum)
return






z:: ; Start Settings with right leg

X = 4
Y = 0
Zoom = 4
Bright = 28

StartSetting(X, Y, Zoom, Bright)

; Create ROI with right leg

sideOfleg = right

Create_ROI(sideOfleg)

open_statistics()

return 







x:: ; move ROI from bone to muscle with right leg

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , 85, 0, ,R

return 





StartSetting(X, Y, Zoom, Bright){

	Mousemove, 18, 780
	Sleep, 50
	MouseClick
	Sleep, 50

	Mousemove, 46, 662
	Sleep, 50
	MouseClick
		
	Sleep, 50
	Send, %X%
	Sleep, 50
	Send, {Enter}
	Sleep, 50
	
	Mousemove, 49, 775
	Sleep, 50
	MouseClick
	Sleep, 50
	Send, {BS}%Zoom%
	Sleep, 50
	Send, {Enter}
	Sleep, 50

	Mousemove, 248, 618
	Sleep, 50
	MouseClick, ,,, 2 ; double click
	Sleep, 50
	Send, %Bright%
	Sleep, 50
	Send, {Enter}
}


Create_ROI(sideOfleg){
	
	anglePosition = 580

	if(sideOfleg = "right"){
		anglePosition -= 5
	}else{
		anglePosition += 5
	}

	MouseClick, left, 19, 687
	Sleep, 50
	MouseClick, left, 61, 635
	Sleep, 50

	MouseClickDrag, L, 545, 580, 553, 617
	Sleep, 200
	MouseClickDrag, L, 552, 580, 599, %anglePosition%

}

open_statistics(){

	Sleep, 50
	Mousemove, 19, 646
	Sleep, 50
	MouseClick
	Sleep, 50
	MouseClick, left, 545, 580

}


closeCards(cardNum){

	Mousemove, 263, 123
	Sleep, 300


	while(cardNum != 0){
		MouseClick
		Sleep, 200
		Send,{TAB}
		Sleep, 200
		Send,{TAB}
		Sleep, 200
		Send, {Enter}
		Sleep, 300
		cardNum --
	}
}



write_temp_femur(){

	Mousemove, 1750, 185 ; clear
	MouseClick

	Mousemove, 157, 699 ; total count
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}

	Sleep, 100
	Mousemove, 157, 819 ; Std Dev
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}

	Sleep, 100
	Mousemove, 157, 848 ; Area
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}
}


write_temp_thigh(){
	; 
}

move_to_thigh(){

MouseClick
Sleep, 200
MouseClickDrag, L, , , -85, 0, ,R

; read statistics by using reading button
; compare if equal by parsing the datas

if not correct 
	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , 5, 0, ,R

}

write_statistics(){
	;destTxtName 
}


^ESC::
	ExitApp

GuiClose:
	ExitApp
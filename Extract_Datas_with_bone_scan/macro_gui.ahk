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




Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Temp window
Gui, Add, Edit, x21 y45 w500 h19 vName, Edit
Gui, Add, Button, x150 y150 w100 h30 gPush, please push
Gui, Show



Push:
Gui,Submit,nohide

Msgbox, %Name%

return



^r:: ; 1. load IDs in file data

Loop, Read, C:\Users\Administrator\Desktop\autohotkey\AutoHotkey\result_of\test.txt
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




s:: ; move ROI from bone to muscle with Left leg

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , -85, 0, ,R

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



^ESC::
	ExitApp

GuiClose:
	ExitApp
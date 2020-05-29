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


femur_total_cnt = 0
femur_std_dev = 0
femur_area = 0

thigh_total_cnt = 0
thigh_std_dev = 0
thigh_area = 0

idTxtName = test.txt
destTxtName = march_result.txt


xpos = 0
ypos = 0



Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Temp window
Gui, Add, Edit, x21 y45 w500 h19 vName, Edit
Gui, Add, Button, x50 y150 w100 h30 gPush_femur, read femur
Gui, Add, Button, x220 y150 w100 h30 gPush_thigh, read thigh
Gui, Show





Push_femur:

global femur_total_cnt
global femur_std_dev 
global femur_area 

Gui,Submit,nohide

; name parsing needed
parsed_temp_window := StrSplit(Name, A_Space, ".")

femur_total_cnt := parsed_temp_window[1]
femur_std_dev := parsed_temp_window[2]
femur_area := parsed_temp_window[3]

;MsgBox % "TC is " femur_total_cnt ", std is " femur_std_dev ", area is " femur_area

return






Push_thigh:

global thigh_total_cnt 
global thigh_std_dev 
global thigh_area 

Gui,Submit,nohide

; name parsing needed
parsed_temp_window := StrSplit(Name, A_Space, ".")

thigh_total_cnt := parsed_temp_window[1]
thigh_std_dev := parsed_temp_window[2]
thigh_area := parsed_temp_window[3]

;MsgBox % "TC is " thigh_total_cnt ", std is " thigh_std_dev ", area is " thigh_area

return





^r:: ; 1. load IDs in file data

Loop, Read, C:\Program Files\AutoHotkey\result_of\%idTxtName%
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

Mousemove, 30, 236
MouseClick

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

Sleep, 100

push_femur_data()

Sleep, 100

move_to_thigh()

Sleep, 100

write_temp_thigh()

Sleep, 100

if( push_thigh_data() = 0){
	MsgBox, "ROI was not moved!"
	return
}

Sleep, 100

write_statistics()

Sleep, 100

initialize_count()

Sleep, 100

close_reviews()

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

	global xpos
	global ypos

	MouseGetPos, xpos, ypos ; store position of mouse for move femur to thigh 

	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send, ^a
	Send, {BS}
	Sleep, 100


	Mousemove, -900, 175
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

	Sleep, 100
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick
}




write_temp_thigh(){

	global xpos
	global ypos

	MouseGetPos, xpos, ypos ; store position of mouse for move femur to thigh 

	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send, ^a
	Send, {BS}
	Sleep, 100


	Mousemove, -900, 175 
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

	Sleep, 100
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick
}





move_to_thigh(){

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , -85, 0, ,R

}








little_move_to_right(){

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , 5, 0, ,R

}








push_femur_data(){

	global xpos
	global ypos

	;MouseGetPos, xpos, ypos ; store position of mouse for move femur to thigh 

	Mousemove, 1482, 189 ; femur push button
	MouseClick

	Sleep, 200

	Mousemove, -900, 175
	MouseClick

	Sleep, 100
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick

}

push_thigh_data(){


	global femur_total_cnt 
	global femur_std_dev 
	global femur_area 

	global thigh_total_cnt 
	global thigh_std_dev 
	global thigh_area 

	global xpos
	global ypos
	
	MouseGetPos, xpos, ypos ; store position of mouse for move femur to thigh 

	Mousemove, 1651, 189 ; thigh push button
	MouseClick

	Sleep, 200

	Mousemove, -900, 175
	MouseClick

	Sleep, 100
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick

	Sleep, 100

	; Check if it moved well 
	;MsgBox, % "femur_total_cnt: " femur_total_cnt ", thigh_total_cnt: " thigh_total_cnt
	
	if (femur_total_cnt = thigh_total_cnt){
		;MsgBox, "same"
		return 0
	}else {
		if(femur_area != thigh_area){
			;MsgBox, "little_move_to_right"
			
			little_move_to_right()

			Sleep, 200	
			
			push_thigh_data()
		}else {
			;MsgBox, "fine"
			return 1
		}
	}
}






write_statistics(){

	global destTxtName 

	global currentID

	global femur_total_cnt 
	global femur_std_dev 
	global femur_area 

	global thigh_total_cnt 
	global thigh_std_dev 
	global thigh_area 

	FileAppend, %currentID% %femur_total_cnt% %femur_std_dev% %femur_area% %thigh_total_cnt% %thigh_std_dev% %thigh_area%`n, C:\Program Files\AutoHotkey\result_of\%destTxtName%

}











initialize_count(){

	global femur_total_cnt 
	global femur_std_dev 
	global femur_area 

	global thigh_total_cnt 
	global thigh_std_dev 
	global thigh_area 

	femur_total_cnt = 0
	femur_std_dev = 0
	femur_area = 0

	thigh_total_cnt = 0
	thigh_std_dev = 0
	thigh_area = 0

}



close_reviews(){

	Mousemove, 262, 125
	Sleep, 300
	MouseClick
	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send, {Enter}

}



^ESC::
	ExitApp

GuiClose:
	ExitApp
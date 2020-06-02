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


femur_total_cnt = 0
femur_std_dev = 0
femur_area = 0

thigh_total_cnt = 0
thigh_std_dev = 0
thigh_area = 0

idTxtName = March_ID.txt
destTxtName = march_result.txt


xpos = 0
ypos = 0



Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text,, Temp window %idIdx%
Gui, Add, Edit, x21 y45 w500 h19 vName, Edit
Gui, Add, Text, x21 y74 w100 h19 ,Process: 
Gui, Add, Edit, x75 y70 w300 h19 vProcess,
Gui, Add, Button, x50 y150 w100 h30 gPush_femur, read femur
Gui, Add, Button, x220 y150 w100 h30 gPush_thigh, read thigh
Gui, Show, x3300 y0





Push_femur:

global femur_total_cnt
global femur_std_dev 
global femur_area

global stopFlag

Gui,Submit,nohide

; name parsing needed
parsed_temp_window := StrSplit(Name, A_Space, ".")

femur_total_cnt := parsed_temp_window[1]
femur_std_dev := parsed_temp_window[2]
femur_area := parsed_temp_window[3]

;MsgBox, 4, , % "femur_area: " femur_area , 3

if (femur_area = 99 or femur_area = 100){
	return
}else{
	Msgbox, 4, , Area is not appropriate. Do you want to continue?
	IfMsgBox No 
		stopFlag = 1
} 

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
	idArrayLen++
}

Msgbox, "Reading proccess complete"

;For index, element in idArray
;{
;	Msgbox, % "Element number" . index . " is " . element
;}

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







d:: ; Start Settings with Left leg

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

















f:: ; 5. if can, press each leg macro key in left side

sideOfleg = left

write_temp_femur()


push_femur_data()


if (stopFlag = 1){
	initialize_count()
	stopFlag = 0
	return
}

move_to_thigh()


write_temp_thigh()


if( push_thigh_data() = 0){
	MsgBox, "ROI was not moved!"
	return
}


write_statistics()


initialize_count()

initialize_progress()

close_reviews()

return 




















r:: ; move ROI to right with Left leg

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , 5, 0, ,R

return 

















t:: ; Switch Left leg to Right leg

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













x:: ; 5. if can, press each leg macro key in Right side

sideOfleg = right

write_temp_femur()


push_femur_data()


if (stopFlag = 1){
	initialize_count()
	stopFlag = 0
	return
}

move_to_thigh() ; right leg


write_temp_thigh()

if( push_thigh_data() = 0){
	MsgBox, "ROI was not moved!"
	return
}


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





Create_ROI(sideOfleg){
	
	anglePosition = 580

	if(sideOfleg = "right"){
		anglePosition -= 5
	}else{
		anglePosition += 5
	}

	MouseClick, left, 19, 687
	MouseClick, left, 61, 635

	MouseClickDrag, L, 545, 580, 553, 617
	Sleep, 200
	MouseClickDrag, L, 552, 580, 599, %anglePosition%

}




open_statistics(){

	Mousemove, 19, 646
	MouseClick
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
	Sleep, 50


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

	Sleep, 50
	Mousemove, 157, 819 ; Std Dev
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}

	Sleep, 50
	Mousemove, 157, 848 ; Area
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}

	Sleep, 50
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
	Sleep, 50


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

	Sleep, 50
	Mousemove, 157, 819 ; Std Dev
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}

	Sleep, 50
	Mousemove, 157, 848 ; Area
	MouseClick, ,,, 2 ; double click
	Send ^c
	Mousemove, 1750, 77 ; temp Window
	MouseClick
	Send ^v
	Send, {Space}
	Send !{Tab}

	Sleep, 50
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick
}




move_to_thigh(){

	global sideOfLeg

	if(sideOfleg = "right"){
		moveTo = 85
	}else{
		moveTo= -85
	}

	Sleep, 50

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , %moveTo%, 0, ,R

}








little_move_to_right(){

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , 5, 0, ,R

}


little_move_to_left(){

	MouseClick
	Sleep, 200
	MouseClickDrag, L, , , -5, 0, ,R

}







push_femur_data(){

	global xpos
	global ypos

	;MouseGetPos, xpos, ypos ; store position of mouse for move femur to thigh 

	Mousemove, 1482, 189 ; femur push button
	MouseClick

	Sleep, 50

	Mousemove, -900, 175
	MouseClick

	Sleep, 50
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick

}

push_thigh_data(){

	global sideOfLeg

	global reIndicateFlag

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

	Sleep, 50

	Mousemove, -900, 175
	MouseClick

	Sleep, 50
	Mousemove, %xpos%, %ypos% ; return to origin of mouse position
	MouseClick

	Sleep, 50

	; Check if it moved well 
	; MsgBox, 4, , % "femur_total_cnt: " femur_total_cnt ", thigh_total_cnt: " thigh_total_cnt, 3
	
	if (femur_total_cnt = thigh_total_cnt){
		MsgBox, 4, , "same", 3
		return 0
	}else {
		if (femur_area != thigh_area){
	
			if (reIndicateFlag = 2){
				reIndicateFlag = 0
				return
			}

			; MsgBox, 4, , % "femur_area: " femur_area ", thigh_area: " thigh_area, 3

			if(sideOfleg = "right"){
				little_move_to_left()
			}else{
				little_move_to_right()
			}

			Sleep, 50	
			write_temp_thigh()
			Sleep, 50
			push_thigh_data()
			
			reIndicateFlag ++

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

initialize_progress(){
	
	global idIdx
	global idArrayLen

	Mousemove, 1750, 100 ; temp Window
	MouseClick
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

	Mousemove, 262, 125
	MouseClick
	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send,{TAB}
	Sleep, 200
	Send, {Enter}

	Sleep, 100

	Mousemove, -740, 50
	MouseClick
}



^ESC::
	ExitApp

GuiClose:
	ExitApp
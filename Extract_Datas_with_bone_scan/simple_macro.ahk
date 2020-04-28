^1:: ; Start Settings with Left leg

X = -4
Y = 0
Zoom = 4
Bright = 28

StartSetting(X, Y, Zoom, Bright)

return



^2:: ; Create ROI

Create_ROI()

return 



^3:: ; Start Settings with Left leg

X = 4
Y = 0
Zoom = 4
Bright = 28

StartSetting(X, Y, Zoom, Bright)

return 



^5:: ; delete Cards
cardNum = 9
closeCards(cardNum)
return




StartSetting(X, Y, Zoom, Bright){

	Mousemove, 18, 780
	Sleep, 200
	MouseClick
	Sleep, 200

	Mousemove, 46, 662
	Sleep, 200
	MouseClick
	Sleep, 200
	Send, %X%
	Sleep, 200
	Send, {Enter}
	Sleep, 200
	
	Mousemove, 49, 775
	Sleep, 200
	MouseClick
	Sleep, 200
	Send, {BS}%Zoom%
	Sleep, 200
	Send, {Enter}
	Sleep, 200

	Mousemove, 248, 618
	Sleep, 200
	MouseClick, ,,, 2 ; double click
	Sleep, 200
	Send, %Bright%
	Sleep, 200
	Send, {Enter}
}


Create_ROI(){

MouseClick, left, 19, 687
Sleep, 200
MouseClick, left, 61, 635
Sleep, 300

MouseClickDrag, L, 545, 580, 553, 617
Sleep, 200
MouseClickDrag, L, 552, 580, 599, 589

}




closeCards(cardNum){

	Mousemove, 263, 123
	Sleep, 300


	while(cardNum != 0){
		MouseClick
		Sleep, 300
		Send,{TAB}
		Sleep, 300
		Send,{TAB}
		Sleep, 300
		Send, {Enter}
		Sleep, 300
		cardNum --
	}
}



^ESC::
ExitApp
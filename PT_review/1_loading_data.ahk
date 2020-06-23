; 1. gui À§Ä¡ Á¶Á¤
; 2. ¸î°³°¡Á®¿Ã Áö ÆÄ¾Ç ¤¡¤¡
; 3. 

Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text, x21 y34 w100 h19 ,Progress 
Gui, Add, Edit, x75 y30 w300 h19 vProcess,
Gui, Add, Text, x21 y74 w100 h19 ,Contents 
Gui, Add, Edit, x75 y70 w400 h300 vMyEdit, ºóÄ­
Gui, Add, Button, x220 y400 w100 h30 gPush_thigh, load reviews
Gui, Show, x1000 y800



Push_thigh:


Gui,Submit,nohide

return



^ESC::
	ExitApp

GuiClose:
	ExitApp
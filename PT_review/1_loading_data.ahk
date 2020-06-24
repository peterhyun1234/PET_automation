; 1. gui À§Ä¡ Á¶Á¤
; 2. ¸î°³°¡Á®¿Ã Áö ÆÄ¾Ç ¤¡¤¡
; 3. 





currentId := ""
currentApprover := ""
currentCI := ""
currentAllComments := ""






Gui, New, hwndhGui AlwaysOnTop
Gui, Add, Text, x21 y34 w100 h19 ,Progress 
Gui, Add, Edit, x75 y30 w300 h19 vProcess,
Gui, Add, Text, x21 y74 w100 h19 ,Contents 
Gui, Add, Edit, x75 y70 w400 h300 vMyEdit, ºóÄ­
Gui, Add, Button, x220 y400 w100 h30 gLoad_data, load reviews
Gui, Show, x1000 y800


Load_data:

Gui,Submit,nohide

parsed_temp_window := MyEdit

MsgBox, 4, , % parsed_temp_window , 3

return



^ESC::
	ExitApp

GuiClose:
	ExitApp
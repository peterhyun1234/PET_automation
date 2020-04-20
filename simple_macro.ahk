F1::
Send, D1    ;
Send, {TAB}  
return

F2::
Send, D2    ; 
Send, {TAB} 
return


F3::
Send, M   ; 
Send, {TAB}
return

F4::
Send, F ; 
Send, {TAB}
return

^x::
MouseClick
Sleep, 100
Send, ^v ; 붙여넣고
Sleep, 100
Send, {Enter}
Sleep, 200
Send, !{TAB}
return


^Esc::
ExitApp















































































































































































































































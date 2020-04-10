^1::

modality = NM
requestingName = bone
studyTerm = 1 ; count for scroll: 기간이 늘수록 데이터가 커지므로 비례하게 스크롤링

FormatTime, currentDate,,yyyyMMdd

currentDate += -1,days  ; += 밖에 안되네

FormatTime, yesterday,%currentDate%,yyyyMMdd

MsgBox, 4, , 의무기록지를 다운로드 받으시겠습니까?`n (이 명령창은 5초 후 자동 종료 됩니다.), 5
IfMsgBox, No
    Return  ; "No" button.
IfMsgBox, Timeout
    Return ; i.e. Assume "No" if it timed out.
; Otherwise, continue:



InputBox, studyDate, 의무기록지 날짜 입력, 다운로드 받으실 의무 기록지의 날짜를 입력하세요. `n ex) 20200409 `n ex) 20200409~20200411, , , , , , , , %yesterday%
if ErrorLevel <> 0
    MsgBox, 0, Exit ready, CANCEL was pressed. , 2
else{

    ; Parsing input string to start date and end date
    ; 파싱하는 부분에서 기간 없는 거 제거해주기
    StringSplit, word_array, studyDate, ~, .  ; Omits periods.

    startDate = %word_array1%
    endDate = %word_array2%

    FormatTime,TargetTime,,%startDate%
    studyTerm := endDate
    TimePrev := studyTerm

    studyTerm -= %TargetTime%, days
    
    MsgBox, 4, , %studyTerm%일: "%studyDate%" 날짜의 데이터 다운로드를 시작하겠습니다. , 5
    IfMsgBox, No
        Return  ; "No" button.
    ; Otherwise, continue:
    

    
    ClickAndInput(studyDate, 625, 180) ; position of INFINIT Date input form
    ClickAndInput(modality, 850, 180) ; position of INFINIT Date input form
    ClickAndInput(requestingName, 1190, 180) ; position of INFINIT Date input form
    
    Mousemove, 1512, 150 ; search
    Sleep, 200
    MouseClick
    Sleep, 200

    Mousemove, 1060, 660 ; for scroll
    Sleep, 200
    MouseClick
    Sleep, 200

    RefreshData(studyTerm) ; data refreshing

    MouseClick 
    Sleep, 1000 ; 마우스 클릭 후 로드될 때까지 1초 대기
    MouseClick, right ; 마우스 오른쪽 버튼 클릭 후
    MouseClick, left, 59, 774, , , , R ; relative로 mouse이동 ㄱㄱ
    ; 

}
return


; 클릭 함수
ClickAndInput(inputText, X_pos, Y_pos)
{
    Mousemove, X_pos, Y_pos ; position of INFINIT Date input form
    Sleep, 200
    MouseClick
    Sleep, 200
    Send, %inputText%
    Sleep, 200
    return
}

; 데이터 수동 새로고침
RefreshData(studyTerm)
{
    while(studyTerm != 0){
        MouseClick, WheelDown, , , 50
        Sleep, 2000
        studyTerm --
    }
    return
}



^Esc::
ExitApp


















































































































































































































































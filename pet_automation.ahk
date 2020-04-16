^1::

modality = NM
requestingName = bone
studyTerm = 1 ; count for scroll: 기간이 늘수록 데이터가 커지므로 비례하게 스크롤링

destPath = D:\shs\ ; 다운로드할 파일이 저장되는 주소 20200415.csv

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

    endDate := startDate

    
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

    StoringData(destPath, endDate)

    
    SourcePattern = {destPath}{endDate}.csv
    DestinationFolder = {destPath}additional_results.csv

    ErrorCount := CopyFilesAndFolders(SourcePattern, DestinationFolder, 1)  ; 아직 제대로 동작안함
    if ErrorCount <> 0
        MsgBox %ErrorCount% files/folders could not be copied.


    ; ExecuteVBA() ; alt + f11, f5 ; 202004 정리본.xlsm 열어서 코드 컴파일 + 실행

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

StoringData(destPath, endDate){
    MouseClick 
    Sleep, 1000 ; 마우스 클릭 후 로드될 때까지 1초 대기
    MouseClick, right ; 마우스 오른쪽 버튼 클릭 후
    MouseClick, left, 59, 774, , , , R ; relative로 mouse이동 ㄱㄱ

    Mousemove, 600, 1000
    Sleep, 200
    MouseClick
    Send, {destPath}{endDate}.csv
    Sleep, 200
    MouseMove, 300, 130
    Sleep, 200
    MouseClick
}


CopyFilesAndFolders(SourcePattern, DestinationFolder, DoOverwrite = false)
{
    FileCopy, %SourcePattern%, %DestinationFolder%, %DoOverwrite%
    ErrorCount := ErrorLevel
    ; Now copy all the folders:
    Loop, %SourcePattern%, 2  ; 2 means "retrieve folders only".
    {
        FileCopyDir, %A_LoopFileFullPath%, %DestinationFolder%\%A_LoopFileName%, %DoOverwrite%
        ErrorCount += ErrorLevel
        if ErrorLevel  ; Report each problem folder by name.
            MsgBox Could not copy %A_LoopFileFullPath% into %DestinationFolder%.
    }
    return ErrorCount
}


/::

Send, !{TAB}    ;   alt + tab
return

*::
Send, ^f    ;   ctrl + f
return


-::
Send, ^{Right}   ;   Ctrl + >
Sleep, 200
Send, {Right}
return



return


^Esc::
ExitApp


















































































































































































































































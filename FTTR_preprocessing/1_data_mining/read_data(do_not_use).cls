Sub fttr_data_preprocessing()


Call read_fttr_data


End Sub


Function read_fttr_data()
    
    On Error GoTo exmsg
    
    Dim path
    path = "F:\Nuclear Medicine\FTTR 정리 완료\FTTR 50개씩 전처리\data\201907-202007 기계별 정리본 with charts.xlsm"
    
    Workbooks(CopyFrom).Activate
    Sheets("Sheet1").Select

    FinalRow_F = Workbooks(CopyFrom).Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
    Range_Sorting_1 = "Q2:Q" & FinalRow_F
    Range_Sorting_2 = "A2:X" & FinalRow_F

    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(Range_Sorting_1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range(Range_Sorting_2)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' MsgBox Mach(c + 1) & " Part has finished (" & c + 1 & "/4)" & vbNewLine & "Now we have " & FinalRow - 2 & " Cases!"
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function


Function findIndexOf(ID)
    'Debug.Print "function called with "; ID; vbNewLine;
    Dim findedIdx
    
    Dim searchID As String
    
    Row = 0
    
    Do
    DoEvents
    
        Row = Row + 1
        searchID = ActiveSheet.UsedRange.Range("A" & Row)
        If InStr(searchID, ID) > 0 Then
            findedIdx = Row
        End If
        
    Loop Until searchID = ""
    
    findIndexOf = findedIdx

End Function


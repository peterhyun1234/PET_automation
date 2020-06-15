Sub inspectionOfFemurToThighRatio()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    
    
    Dim lastRow As Long
    
    'Find the last non-blank cell in column A(1)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim idx
    
    idx = 2

    Do

    Debug.Print idx & ": lastRow is " & lastRow & vbNewLine
    
    idx = idx + 1
    
    Loop While idx <= lastRow

    
    
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function


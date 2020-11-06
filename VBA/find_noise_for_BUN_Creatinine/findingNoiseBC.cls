Sub findingNoiseBC()


Call findNoises


End Sub



Function findNoises()
    
    On Error GoTo exmsg
    
    Dim degreeOfNoises
    degreeOfNoises = 0
    
    Dim lastRow As Long
    
    'Find the last non-blank cell in column A(1)
    ' lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    lastRow = 252
    
    Dim idx
    
    idx = 3

    Do
    'Debug.Print idx & ": lastRow is " & lastRow & vbNewLine
    Dim age
    Dim bun
    Dim creatinine
    age = Cells(idx, 4).Value
    sex = Cells(idx, 3).Value
    bun = Cells(idx, 22).Value
    creatinine = Cells(idx, 23).Value

    Debug.Print age & " " & bun & " " & creatinine & vbNewLine
    
    If bun = "" Then
        Cells(idx, 22).Interior.Color = RGB(100, 100, 100)
    
    End If

    If creatinine = "" Then
        Cells(idx, 23).Interior.Color = RGB(100, 100, 100)
    
    End If



    If sex = "M" Then
        ' 여기에서 나이 분류
    ELSE 

    End If










    idx = idx + 1
    
    Loop While idx <= lastRow

    ' MsgBox ("Finded noises: " & degreeOfNoises)
    
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function



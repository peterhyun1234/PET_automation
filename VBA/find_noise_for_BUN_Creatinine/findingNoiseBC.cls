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
    Dim id
    Dim sex
    Dim age
    Dim bun
    Dim creatinine
    id = Cells(idx, 1).Value
    sex = Cells(idx, 3).Value
    age = Cells(idx, 4).Value
    bun = Cells(idx, 22).Value
    creatinine = Cells(idx, 23).Value

    
    
    If bun = "" Then
        Cells(idx, 22).Interior.Color = RGB(100, 100, 100)
    
    End If

    If creatinine = "" Then
        Cells(idx, 23).Interior.Color = RGB(100, 100, 100)
    
    End If

    ' BUN noises check
    IF age >= 60 Then
        ' Debug.Print id & ": " & age & vbNewLine
        ' 나이가 60대 이상이면 8.0 - 23.0

    ELSE
        ' 나이가 60대 이하이면 6.0 - 20.0

    End IF

    ' Creatinine noises check
    If sex = "M" Then
        ' Debug.Print id & ": " & sex & vbNewLine
        ' 성별이 남자면 C: 0.7 - 1.2
    ELSE 
        ' 성별이 여자면 C: 0.5 - 0.9

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



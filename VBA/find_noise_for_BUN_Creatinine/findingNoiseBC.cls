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

    
    
    ' BUN noises check
    If bun = "" Then
        Cells(idx, 22).Interior.Color = RGB(100, 100, 100)
    ELSE
        IF age >= 60 Then

            ' 나이가 60대 이상이면 8.0 - 23.0

            IF bun <= 8 Then
                Cells(idx, 22).Interior.Color = RGB(255, 155, 155)
                degreeOfNoises = degreeOfNoises + 1
            ELSEIF bun >= 23 Then
                Cells(idx, 22).Interior.Color = RGB(155, 155, 255)
                degreeOfNoises = degreeOfNoises + 1
            End IF

        ELSE

            ' 나이가 60대 이하이면 6.0 - 20.0

            IF bun <= 6 Then
                Cells(idx, 22).Interior.Color = RGB(255, 155, 155)
                degreeOfNoises = degreeOfNoises + 1
            ELSEIF bun >= 20 Then
                Cells(idx, 22).Interior.Color = RGB(155, 155, 255)
                degreeOfNoises = degreeOfNoises + 1
            End IF

        End IF
    End If




    ' Creatinine noises check
    If creatinine = "" Then
        Cells(idx, 23).Interior.Color = RGB(100, 100, 100)
    ELSE 
        If sex = "M" Then

            ' 성별이 남자면 C: 0.7 - 1.2

            IF creatinine <= 0.7 Then
                Cells(idx, 23).Interior.Color = RGB(255, 155, 155)
                degreeOfNoises = degreeOfNoises + 1
            ELSEIF creatinine >= 1.2 Then
                Cells(idx, 23).Interior.Color = RGB(155, 155, 255)
                degreeOfNoises = degreeOfNoises + 1
            End IF
        ELSE 

            ' 성별이 여자면 C: 0.5 - 0.9

            IF creatinine <= 0.5 Then
                Cells(idx, 23).Interior.Color = RGB(255, 155, 155)
                degreeOfNoises = degreeOfNoises + 1
            ELSEIF creatinine >= 0.9 Then
                Cells(idx, 23).Interior.Color = RGB(155, 155, 255)
                degreeOfNoises = degreeOfNoises + 1
            End IF

        End If
    End If

    


    idx = idx + 1
    
    Loop While idx <= lastRow

    MsgBox ("Finded noises: " & degreeOfNoises)
    
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function



Sub FTTR_for_study_inspection()


Call findNoises


End Sub



Function findNoises()
    
    On Error GoTo exmsg
    
    Dim degreeOfNoises
    degreeOfNoises = 0
    
    Dim lastRow As Long
    
    'Find the last non-blank cell in column A(1)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim idx
    
    idx = 2

    Do
    'Debug.Print idx & ": lastRow is " & lastRow & vbNewLine
    Dim comments As String
    comments = Cells(idx, 16).Value
    
    If Not comments = "" Then
        Cells(idx, 16).Interior.Color = RGB(255, 255, 0)
        'Debug.Print idx & comments & vbNewLine
    Else
        Dim FTC
        FTC = Cells(idx, 8).Value
        Dim FSD
        FSD = Cells(idx, 10).Value
        Dim FAR
        FAR = Cells(idx, 11).Value
        
        Dim TTC
        TTC = Cells(idx, 12).Value
        Dim TSD
        TSD = Cells(idx, 14).Value
        Dim TAR
        TAR = Cells(idx, 15).Value
        
        If Not (FTC > 300 And FTC < 7000) Then
            Debug.Print idx & "FTC: " & FTC & vbNewLine
            Cells(idx, 8).Interior.Color = RGB(255, 0, 0)
            degreeOfNoises = degreeOfNoises + 1
        End If
        
        If Not (FSD > 0 And FSD < 20) Then
            Debug.Print idx & "FSD: " & FSD & vbNewLine
            Cells(idx, 10).Interior.Color = RGB(255, 0, 0)
            degreeOfNoises = degreeOfNoises + 1
        End If
        
        If Not (FAR > 95 And FAR < 110) Then
            Debug.Print idx & "FAR: " & FAR & vbNewLine
            Cells(idx, 11).Interior.Color = RGB(255, 0, 0)
            degreeOfNoises = degreeOfNoises + 1
        End If
        
        If Not (TTC > 100 And TTC < 5000) Then
            Debug.Print idx & "TTC: " & TTC & vbNewLine
            Cells(idx, 12).Interior.Color = RGB(255, 0, 0)
            degreeOfNoises = degreeOfNoises + 1
        End If
        
        If Not (TSD > 0 And TSD < 10) Then
            Debug.Print idx & "TSD: " & TSD & vbNewLine
            Cells(idx, 14).Interior.Color = RGB(255, 0, 0)
            degreeOfNoises = degreeOfNoises + 1
        End If
        
        If Not (TAR > 95 And TAR <= 110) Then
            Debug.Print idx & "TAR: " & TAR & vbNewLine
            Cells(idx, 15).Interior.Color = RGB(255, 0, 0)
            degreeOfNoises = degreeOfNoises + 1
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



Sub inspectionOfParsingReviews()

Call deleteNoises

End Sub



Function deleteNoises()
    
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
    Dim AETitle As String
    AETitle = Cells(idx, 8).Value
    
    If AETitle = "dst1" Or AETitle = "dst2" Then
        Dim approver
        approver = Cells(idx, 11).Value

        If approver = "" Then
            Cells(idx, 11).Interior.Color = RGB(255, 0, 0)
        Else
            Dim comments
            comments = Cells(idx, 13).Value
            If InStr(comments, "*Clinical") Or InStr(comments, "*clinical") Then 'comments에서 특정 CI 찾기
                Cells(idx, 12).Value = "change needed"
                Cells(idx, 11).Interior.Color = RGB(0, 255, 0)
                Cells(idx, 12).Interior.Color = RGB(0, 255, 0)
                Cells(idx, 13).Interior.Color = RGB(0, 255, 0)
            Else
                Dim CI
                Dim i As Long
                CI = Cells(idx, 12).Value
                If CI = "" Or CI = " " Then
                    Dim parsedCommentsTemp() As String
                    parsedCommentsTemp() = Split(comments, "(clinical information: ")
                    If UBound(parsedCommentsTemp) > 0 Then
                        Dim parsedComments() As String
                        parsedComments() = Split(parsedCommentsTemp(1), ")")
                        'Debug.Print "idx: "; idx
                        'Debug.Print "CI: "; parsedComments(0)
                        Cells(idx, 12).Value = parsedComments(0)
                    End If
                End If
            End If
        End If


    Else
        Rows(idx).Delete
        idx = idx - 1
        lastRow = lastRow - 1
        degreeOfNoises = degreeOfNoises + 1
    End If

    idx = idx + 1
    
    Loop While idx <= lastRow

    MsgBox ("Finded noises: " & degreeOfNoises & vbNewLine)
    
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function

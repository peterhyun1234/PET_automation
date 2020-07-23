Sub orderingByApprover()

Call orderingByApproverWithDividedSheet

End Sub



Function orderingByApproverWithDividedSheet()
    
    On Error GoTo exmsg
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Dim error_cnt
    error_cnt = 0

    Dim sj_cnt
    sj_cnt = 2
    Dim ys_cnt
    ys_cnt = 2
    Dim jk_cnt
    jk_cnt = 2

    Dim idx
    idx = 2

    Do
        'Debug.Print idx & ": lastRow is " & lastRow & vbNewLine
        Dim approver
        approver = Cells(idx, 11).Value
        Dim currentRange
        currentRange = "A" & idx & ":M" & idx

        If approver = "이수진" Then
            Worksheets("Sheet1").Range(currentRange).Copy Worksheets("이수진").Range("A" & sj_cnt)
            sj_cnt = sj_cnt + 1
        Elseif approver = "안영실" Then
            Worksheets("Sheet1").Range(currentRange).Copy Worksheets("안영실").Range("A" & ys_cnt)
            ys_cnt = ys_cnt + 1
        Elseif approver = "윤준기" Then
            Worksheets("Sheet1").Range(currentRange).Copy Worksheets("윤준기").Range("A" & jk_cnt)
            jk_cnt = jk_cnt + 1
        Else
            Cells(idx, 11).Interior.Color = RGB(0, 255, 0)
            error_cnt = error_cnt + 1
        End If
        

        idx = idx + 1
    
    Loop While idx <= lastRow

    MsgBox ("sj_cnt: " & sj_cnt & vbNewLine)
    MsgBox ("ys_cnt: " & ys_cnt & vbNewLine)
    MsgBox ("jk_cnt: " & jk_cnt & vbNewLine)
    
    Dim total_cnt
    total_cnt = sj_cnt + ys_cnt + jk_cnt
    MsgBox ("lastRow: " & lastRow & vbNewLine & "total_cnt: " & total_cnt & vbNewLine & "error_cnt: " & error_cnt & vbNewLine)
    
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function



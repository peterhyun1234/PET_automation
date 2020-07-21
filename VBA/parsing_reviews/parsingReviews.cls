Sub parsingReviews()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    Dim path
    path = "C:\판독문 정리\process_1\2018\20_reviews_1.txt"
    
    Dim userNum
    userNum = 1

    Dim lineNum
    lineNum = 0
    
    Dim lineValue As String
    
    Dim ID
    Dim startIdx

    startIdx = 0;
    
    Dim Approver
    Dim CI
    
    Dim comments
    

    Open path For Input As #1
        Do While Not EOF(1)

            Do While (True)
                Line Input #1, lineValue

                If (StrComp(lineValue, "///", vbTextCompare) = 0) Then
                    'Debug.Print "lineNum: "; lineNum; ", userNum"; userNum; vbNewLine;
                    'Debug.Print "all comments: "; vbNewLine; comments; vbNewLine; vbNewLine;


                    Dim idxOfID
                    
                    idxOfID = findIndexOf(ID, startIdx)
                    'Debug.Print "idxOfID: "; idxOfID; ", ID: "; ID; vbNewLine;

                    Cells(idxOfID, 11).Value = Approver
                    Cells(idxOfID, 12).Value = CI
                    Cells(idxOfID, 13).Value = comments

                    ID = ""
                    Approver = ""
                    CI = ""
                    comments = ""
                    lineNum = 0
                    Exit Do
                End If
                

                If (lineNum = 0) Then ' ID parsing
                    ID = Right(lineValue, Len(lineValue) - 4)
                    ID = Left(ID, 7)
                    'Debug.Print "lineNum: "; lineNum; ", ID: "; ID; vbNewLine;

                ElseIf (lineNum = 1) Then ' Approver parsing
                    Approver = Right(lineValue, Len(lineValue) - 10)
                    'Debug.Print "lineNum: "; lineNum; ", Approver: "; Approver; vbNewLine;

                ElseIf (lineNum = 2) Then ' CI parsing
                    CI = Right(lineValue, Len(lineValue) - 4)
                    'Debug.Print "lineNum: "; lineNum; ", CI: "; CI; vbNewLine;

                ElseIf (lineNum = 3) Then ' skip

                Else ' All comments parsing
                    comments = comments & lineValue
                End If

                lineNum = lineNum + 1

            Loop

            userNum = userNum + 1
        Loop
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function


Function findIndexOf(ID, startIdx)
    'Debug.Print "function called with "; ID; vbNewLine;
    Dim findedIdx
    
    Dim searchID As String
    
    Row = startIdx
    
    Do
    DoEvents
    
        Row = Row + 1
        searchID = ActiveSheet.UsedRange.Range("A" & Row)
        If InStr(searchID, ID) > 0 And Row > 10000 Then
            findedIdx = Row
            Exit Do
        End If
        
    Loop Until searchID = ""
    
    findIndexOf = findedIdx

End Function




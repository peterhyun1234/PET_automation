Sub parsingReviews()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    Dim path
    path = "F:\Nuclear Medicine\판독문 정리\process_1\202005_reviews_1.txt"
    
    Dim userNum
    userNum = 1

    Dim lineNum
    lineNum = 0
    
    Dim lineValue As String
    
    Dim ID
    
    Dim Approver
    Dim CI
    
    Dim comments
    

    Open path For Input As #1
        Do While Not EOF(1)

            Do While (True)
                Line Input #1, lineValue

                If (StrComp(lineValue, "///", vbTextCompare) = 0) Then
                    'Debug.Print "lineNum: "; lineNum; ", userNum"; userNum; vbNewLine;
                    Debug.Print "all comments: "; vbNewLine; comments; vbNewLine; vbNewLine;
                    ID = ""
                    Approver = ""
                    CI = ""
                    comments = ""
                    lineNum = 0
                    Exit Do
                End If
                

                If (lineNum = 0) Then ' ID parsing
                    ID = Right(lineValue, Len(lineValue) - 4)
                    Debug.Print "lineNum: "; lineNum; ", ID: "; ID; vbNewLine;

                Elseif (lineNum = 1) Then ' Approver parsing
                    Approver = Right(lineValue, Len(lineValue) - 10)
                    Debug.Print "lineNum: "; lineNum; ", Approver: "; Approver; vbNewLine;

                Elseif (lineNum = 2) Then ' CI parsing
                    CI = Right(lineValue, Len(lineValue) - 4)
                    Debug.Print "lineNum: "; lineNum; ", CI: "; CI; vbNewLine;

                Elseif (lineNum = 3) Then ' skip

                Else ' All comments parsing
                    comments = comments & lineValue
                End If

                lineNum = lineNum + 1

            Loop
            
            'ID = parsedLine(0)
            
            'Dim idxOfID
            
            'idxOfID = findIndexOf(ID)
                    
            
            userNum = userNum + 1
        Loop
    Close #1
    
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

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

            Do While (i<5)
                Line Input #1, lineValue

                If (StrComp(lineValue, "///", vbTextCompare) = 0) Then
                    Debug.Print "lineNum: "; lineNum; ", userNum"; userNum; vbNewLine;
                    lineNum = 0
                    Exit Do
                End If
                

                If (lineNum = 0) Then ' ID parsing
                    Debug.Print "lineNum: "; lineNum; ", lineValue"; lineValue; vbNewLine;

                    'parsedLine = Split(lineValue)

                Elseif (lineNum = 1) Then ' Approver parsing
                    Debug.Print "lineNum: "; lineNum; ", lineValue"; lineValue; vbNewLine;

                    'parsedLine = Split(lineValue)

                Elseif (lineNum = 2) Then ' CI parsing
                    Debug.Print "lineNum: "; lineNum; ", lineValue"; lineValue; vbNewLine;

                    'parsedLine = Split(lineValue)

                Else ' All comments parsing

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


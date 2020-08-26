Sub measurementOfFemurToThighRatio()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    Dim path
    path = "F:\Nuclear Medicine\FTTR 정리 완료\202007\202007_result.txt"
    
    Dim lineNum
    lineNum = 1
    
    Dim lineValue As String
    
    Dim ID
    
    Dim femurTotalCnt
    Dim femurSD
    Dim femurArea
    
    Dim thighTotalCnt
    Dim thighSD
    Dim thighArea
    
    Dim comments
    

    Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineValue
            ' Debug.Print "Line value"; lineNum; ": "; lineValue; vbNewLine;
            
            parsedLine = Split(lineValue)
            
            ID = parsedLine(0)
            
            Dim arrLen
            
            arrLen = UBound(parsedLine) - LBound(parsedLine) + 1
            
            ID = parsedLine(0)
            
            Dim idxOfID
            
            idxOfID = findIndexOf(ID)
            
            
            If arrLen < 3 Then
                ' Debug.Print idxOfID; "."; parsedLine(0); ": "; parsedLine(1); vbNewLine;
                comments = parsedLine(1)
                
                Cells(idxOfID, 16).Value = comments
                
            Else

                femurTotalCnt = parsedLine(1)
                femurSD = parsedLine(2)
                femurArea = parsedLine(3)
                
                thighTotalCnt = parsedLine(4)
                thighSD = parsedLine(5)
                thighArea = parsedLine(6)
                
                Cells(idxOfID, 8).Value = femurTotalCnt
                Cells(idxOfID, 10).Value = femurSD
                Cells(idxOfID, 11).Value = femurArea
                
                Cells(idxOfID, 12).Value = thighTotalCnt
                Cells(idxOfID, 14).Value = thighSD
                Cells(idxOfID, 15).Value = thighArea
            
            
            End If
            
            
            lineNum = lineNum + 1
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


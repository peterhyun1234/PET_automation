Sub FTTR_for_study()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    Dim path
    path = "C:\Users\ajounm\Desktop\ajounm\for_study\202008\202008_result.txt"
    
    Dim lineNum
    lineNum = 1
    
    Dim lineValue As String
    
    Dim ID
    
    Dim firstFemurTotalCnt
    Dim firstFemurSD
    Dim firstFemurArea
    
    Dim firstThighTotalCnt
    Dim firstThighSD
    Dim firstThighArea

    Dim secondFemurTotalCnt
    Dim secondFemurSD
    Dim secondFemurArea
    
    Dim secondThighTotalCnt
    Dim secondThighSD
    Dim secondThighArea

    Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineValue
            ' Debug.Print "Line value"; lineNum; ": "; lineValue; vbNewLine;
            
            parsedLine = Split(lineValue)
            
            Dim arrLen
            
            arrLen = UBound(parsedLine) - LBound(parsedLine) + 1
            
            ID = parsedLine(0)
            
            Dim idxOfID
            
            idxOfID = findIndexOf(ID)
            
            
            If arrLen < 3 Then
                Debug.Print idxOfID; "."; parsedLine(0); ": "; parsedLine(1); vbNewLine;
                ' comments = parsedLine(1)
                
                ' Cells(idxOfID, 16).Value = comments
                
            Else

                firstFemurTotalCnt = parsedLine(1)
                firstFemurSD = parsedLine(2)
                firstFemurArea = parsedLine(3)
                
                firstThighTotalCnt = parsedLine(4)
                firstThighSD = parsedLine(5)
                firstThighArea = parsedLine(6)

                secondFemurTotalCnt = parsedLine(7)
                secondFemurSD = parsedLine(8)
                secondFemurArea = parsedLine(9)
                
                secondThighTotalCnt = parsedLine(10)
                secondThighSD = parsedLine(11)
                secondThighArea = parsedLine(12)

                ' Debug.Print idxOfID; "."; secondThighSD; ": "; secondThighArea; vbNewLine;
                
                Cells(idxOfID, 8).Value = firstFemurTotalCnt
                Cells(idxOfID, 10).Value = firstFemurSD
                Cells(idxOfID, 11).Value = firstFemurArea
                
                Cells(idxOfID, 12).Value = firstThighTotalCnt
                Cells(idxOfID, 14).Value = firstThighSD
                Cells(idxOfID, 15).Value = firstThighArea
            
                Cells(idxOfID, 17).Value = secondFemurTotalCnt
                Cells(idxOfID, 19).Value = secondFemurSD
                Cells(idxOfID, 20).Value = secondFemurArea
                
                Cells(idxOfID, 21).Value = secondThighTotalCnt
                Cells(idxOfID, 23).Value = secondThighSD
                Cells(idxOfID, 24).Value = secondThighArea

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


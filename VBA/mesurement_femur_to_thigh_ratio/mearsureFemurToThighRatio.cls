Sub measurementOfFemurToThighRatio()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    Dim path
    path = "C:\Users\Ajou NM\Desktop\shs\March_data\march_result.txt"
    
    Dim lineNum
    lineNum = 1
    
    Dim lineValue As String

    Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineValue
            ' Debug.Print "Line value"; lineNum; ": "; lineValue; vbNewLine;
            
            parsedLine = Split(lineValue)
            
            Dim ID
            
            Dim femurTotalCnt
            Dim femurSD
            Dim femurArea
            
            Dim thighTotalCnt
            Dim thighSD
            Dim thighArea
            
            Dim comments
            
            ID = parsedLine(0)
            
            If Len(parsedLine) < 3 Then
            Debug.Print "Line value"; parsedLine; vbNewLine;

            End If
            lineNum = lineNum + 1
        Loop
    Close #1
    
    Exit Function
    
exmsg:
    Close #1
    MsgBox ("Error" & Err.Number & " : " & Err.Description)

End Function


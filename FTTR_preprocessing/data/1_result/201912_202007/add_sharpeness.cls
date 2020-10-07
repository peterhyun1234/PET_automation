Sub add_sharpeness()


Call readTextFile


End Sub



Function readTextFile()
    
    On Error GoTo exmsg
    
    Dim path
    path = "C:\Users\ajoumc\인턴 전현빈\workspace\PET_automation\FTTR_preprocessing\1_data_mining\result\201912_202007\201912_202007_result.txt"
    
    Dim lineNum
    lineNum = 1
    
    Dim lineValue As String
    
    Dim ID
    
    Dim sharpeness
    
    Dim pos As Integer
    Dim sheet_name As String

    Dim arrLen


    sheet_name = "Infi"
    Worksheets(sheet_name).Activate


    Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineValue

            If InStr(lineValue, "#") Then

                arrLen = Len(lineValue)

                ' Debug.Print "arrLen: "; arrLen; vbNewLine;

                pos = InStr(lineValue, "#")

                sheet_name = Mid(lineValue, pos + 1, arrLen - 1)
                
                ' change sheet(select)

                Worksheets(sheet_name).Activate

                MsgBox "쉬트: " & ActiveSheet.Name
            
            Else
                parsedLine = Split(lineValue)
                
                ID = parsedLine(0)
                sharpeness = parsedLine(1)
                
                Dim idxOfID
                idxOfID = findIndexOf(ID)

                Cells(idxOfID, 18).Value = sharpeness

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

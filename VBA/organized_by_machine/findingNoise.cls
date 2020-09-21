Sub findingNoises()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                      '
CopyTo = "temp.xlsm"'
                                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Total As Long
Total = 0

'''To find a number of Row which is the last for each Machine'''

Dim cellarr(0 To 4) As Long
Dim Mach() As Variant

Mach = [{"Infi","NM830","NMCT670","Sym"}]

' For i = 1 To 4
' With Sheets("Sheet1")
'     cellarr(i) = .Range("Q:Q").Find(what:=Mach(i), after:=.Range("Q2"), searchdirection:=xlPrevious).Row
' End With
' Next i


' cellarr(0) = 1


For c = 0 To 3

    Workbooks(CopyTo).Activate
    Workbooks(CopyTo).Sheets(Mach(c + 1)).Select
    finalRow = Sheets(Mach(c + 1)).Cells(Rows.Count, 1).End(xlUp).Row
            
    ' find noise and data formatting
    ' 처음부터 끝까지 데이터 없는 거 검사하고 No_Data 기입
            
    Dim startTime
    Dim endTime
    
    Dim searchID As String
    
    curRow = 3
    
    Do
    DoEvents
    
        startTime = ActiveSheet.UsedRange.Range("V" & curRow)
        endTime = ActiveSheet.UsedRange.Range("W" & curRow)

        If isNull(startTime) Or isNull(endTime) Then
            searchID = ActiveSheet.UsedRange.Range("A" & curRow)
            ' time_inteval <- "No_data"
            Debug.Print "searchID: "; searchID; vbNewLine;
        End If
        
        curRow = curRow + 1

    Loop Until Row <= finalRow

       
    ' print with msgbox for each session ended status        
    MsgBox Mach(c + 1) & " Part has finished (" & c + 1 & "/4)" & vbNewLine & "Now we have " & finalRow - 2 & " Cases!"

    Total = Total + finalRow - 2

Next c


''''For closing without saving changed in "CopyFrom"'''
Workbooks(CopyFrom).Close Savechanges:= False


MsgBox "Complete" & vbNewLine & "The Number of Cases is " & Total & " in total."

End Sub

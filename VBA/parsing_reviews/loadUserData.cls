Sub change()

Sheets("Sheet1").Select
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

Columns("A:Z").Hidden = False


''''''''''''''''''''''''''''''''''''''
Origin = "20_review_result.xlsm" '
                                     '
WBS = "20_reviews_2"             '
WB = WBS + ".csv"                    '
''''''''''''''''''''''''''''''''''''''

Workbooks.Open Filename:="F:\Nuclear Medicine\판독문 정리\process_2\" & WB

Workbooks(WB).Activate
Sheets(WBS).Select
lastRow = Workbooks(WB).Sheets(WBS).Cells(Rows.Count, 1).End(xlUp).Row

LastCell_1 = "J" & lastRow

Range_1 = "A2:" & LastCell_1

Dest_1 = "A" & FinalRow + 1

Workbooks(WB).Sheets(WBS).Range(Range_1).Copy Destination:=Workbooks(Origin).Sheets("sheet1").Range(Dest_1)


Workbooks(Origin).Activate
Sheets("Sheet1").Select
StartRow = FinalRow + 1
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
Range_Sorting_1 = "F" & StartRow & ":F" & FinalRow
Range_Sorting_2 = "A" & StartRow & ":X" & FinalRow

i = 2

Cells(i, 1).NumberFormat = "@"
Dim num0 As String
num0 = Cells(i, 1)
Cells(i, 1).Value = Right("0000000" & num0, 7)

        
    
    For i = StartRow To FinalRow
    
        Cells(i, 1).NumberFormat = "@"
        Dim num As String
        num = Cells(i, 1)
        Cells(i, 1).Value = Right("0000000" & num, 7)
        
            
        If Right(Cells(i, 4).Value, 1) = "Y" Then Cells(i, 4).Value = Left(Cells(i, 4).Value, 3)
        If Right(Cells(i, 4).Value, 1) = "M" Then Cells(i, 4).Value = Left(Cells(i, 4).Value, 3) + (Mid(Cells(i, 4).Value, 5, 3) / 12)
        
        
        'If Cells(i, 8).Value = "XE22" Then Cells(i, 10).Value = "Infi"
        'If Cells(i, 8).Value = "XE42" Then Cells(i, 10).Value = "NMCT670"
        'If Cells(i, 8).Value = "SYMBIAE1429" Then Cells(i, 10).Value = "Sym"
        'If Cells(i, 8).Value = "XE12" Then Cells(i, 10).Value = "NM830"
        
        If (Cells(i, 8).Value = "READOUT_T" Or Cells(i, 8).Value = "PIVIEW") Then
            Rows(i).Delete
            i = i - 1
        End If
        
         
    Next i


ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(Range_Sorting_1), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Sheet1").Sort
    .SetRange Range(Range_Sorting_2)
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

MsgBox "Data Merging & Reforming Complete"

End Sub




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Data_Merging()


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                      '
CopyFrom = "201911 ������.xlsm"                       '
                                                      '
CopyTo = "201907-201911 ��躰 ������ with charts"    '
                                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Workbooks(CopyTo).Activate
Columns("A:Z").Hidden = False

Workbooks.Open Filename:=ThisWorkbook.Path & "\" & CopyFrom

Dim Total As Long
Total = 0

''''For rearanging by Machine'''

Workbooks(CopyFrom).Activate
Sheets("Sheet1").Select

FinalRow_F = Workbooks(CopyFrom).Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
Range_Sorting_1 = "Q2:Q" & FinalRow_F
Range_Sorting_2 = "A2:X" & FinalRow_F

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


'''To find a number of Row which is the last for each Machine'''

Dim cellarr(0 To 4) As Long
Dim Mach() As Variant

Mach = [{"Infi","NM830","NMCT670","Sym"}]

For i = 1 To 4
With Sheets("Sheet1")
    cellarr(i) = .Range("Q:Q").Find(what:=Mach(i), after:=.Range("Q2"), searchdirection:=xlPrevious).Row
End With
Next i


'''Copy the data From "CopyFrom" To "CopyTo"'''

cellarr(0) = 1




For c = 0 To 3

    Workbooks(CopyTo).Activate
    Workbooks(CopyTo).Sheets(Mach(c + 1)).Select
    FinalRow_ini = Sheets(Mach(c + 1)).Cells(Rows.Count, 1).End(xlUp).Row

    First = "A" & cellarr(c) + 1
    Last = "X" & cellarr(c + 1)

    CopyRange = First & ":" & Last

    Dest = "A" & FinalRow_ini - 1

    Workbooks(CopyFrom).Activate

    With ActiveWorkbook.Sheets("Sheet1")
         Workbooks(CopyFrom).Sheets("Sheet1").Range(CopyRange).Copy
         Workbooks(CopyTo).Sheets(Mach(c + 1)).Range(Dest).Insert Shift:=xlDown
    End With


'''For rearranging by Study Date'''

    Workbooks(CopyTo).Activate
    Workbooks(CopyTo).Sheets(Mach(c + 1)).Select
    FinalRow = Sheets(Mach(c + 1)).Cells(Rows.Count, 1).End(xlUp).Row

    Range_Sorting_1 = "F3:F" & FinalRow
    Range_Sorting_2 = "A3:X" & FinalRow

    ActiveWorkbook.Worksheets(Mach(c + 1)).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(Mach(c + 1)).Sort.SortFields.Add Key:=Range(Range_Sorting_1), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(Mach(c + 1)).Sort
        .SetRange Range(Range_Sorting_2)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(FinalRow_ini + 8, 25).Copy
    Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(FinalRow + 9, 25).Select
    ActiveSheet.Paste
    Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(FinalRow_ini + 8, 25).Delete
    Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(FinalRow + 8, 25).NumberFormat = "0.000%"
    'Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(FinalRow + 8, 25).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1

    Dim ComCell(1 To 4) As Integer
    
    ComCell(1) = 13
    ComCell(2) = 54
    ComCell(3) = 27
    ComCell(4) = 41
    
    Workbooks(CopyTo).Sheets("Comparing").Cells(ComCell(c + 1), 13) = "=" & Mach(c + 1) & "!Y" & FinalRow + 8
   
    
'''For filtering non-Data IDs'''

    Row = 3
    
        Do While Row <= FinalRow
            Workbooks(CopyTo).Sheets(Mach(c + 1)).Select
            
            If Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(Row, 8).Text = none Then
                FinalRow_Residue = Sheets("Residue").Cells(Rows.Count, 1).End(xlUp).Row
                Sheets(Mach(c + 1)).Rows(Row).Copy Worksheets("Residue").Range("A" & FinalRow_Residue + 2)
                Sheets(Mach(c + 1)).Rows(Row).Delete
                FinalRow = FinalRow - 1
                Row = Row - 1
            End If
            
            If Workbooks(CopyTo).Sheets(Mach(c + 1)).Cells(Row, 16).Text = none Then
                Sheets(Mach(c + 1)).Cells(Row - 1, 16).Copy
                Sheets(Mach(c + 1)).Cells(Row, 16).Select
                ActiveSheet.Paste
            End If
        
            Row = Row + 1
            
        Loop

    '''To draw Outlines'''
    
        Sheets(Mach(c + 1)).Range("B1:G" & FinalRow).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1
        Sheets(Mach(c + 1)).Range("L1:O" & FinalRow).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1
        Sheets(Mach(c + 1)).Range("Q1:Q" & FinalRow).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1
        Sheets(Mach(c + 1)).Range("V1:W" & FinalRow).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1
        Sheets(Mach(c + 1)).Range("X1:X" & FinalRow).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1
            
     MsgBox Mach(c + 1) & " Part has finished (" & c + 1 & "/4)" & vbNewLine & "Now we have " & FinalRow - 2 & " Cases!"

     Total = Total + FinalRow - 2

Next c


''''For closing without saving changed in "CopyFrom"'''
Workbooks(CopyFrom).Close Savechanges:=False


MsgBox "Merging Data Complete" & vbNewLine & "The Number of Cases is " & Total & " in total."

End Sub




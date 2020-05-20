VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ex()



For i = 1 To 420

Cells(i + 2, 4).Value = Left(Cells(i + 2, 4).Value, 3)

Next i


End Sub


Sub STD_range()

Dim matA(1, 660) As Double
Dim matB(1, 660) As Double
Dim matC(1, 660) As Double
Dim matD(1, 660) As Double
Dim matE(1, 660) As Double
Dim matF(1, 660) As Double
Dim matG(1, 660) As Double
Dim matH(1, 660) As Double
Dim matI(1, 660) As Double


For j = 1 To 9

For i = 1 To 660

If Cells(i + 1, 4).Value >= (j - 1) * 10 And Cells(i + 1, 4).Value < j * 10 Then
matA(1, i + 1) = Cells(i + 1, 9)

Next i

Next j


Cells(714, 6).Value = matA


End Sub



Sub AET()


For i = 1 To 664

For j = 1 To 700

If Cells(j + 2, 1).Value = Cells(i + 2, 25) Then Cells(j + 2, 21).Value = Cells(i + 2, 32).Value

Next j

Next i

End Sub

Sub change()


For i = 1 To 700

If Cells(i, 17).Value = "XE22" Then Cells(i, 17).Value = "Infi"
If Cells(i, 17).Value = "XE42" Then Cells(i, 17).Value = "NMCT670"
If Cells(i, 17).Value = "SYMBIAE1429" Then Cells(i, 17).Value = "Sym"

Next i



End Sub


Sub DATA()

Sht = ["NMCT670"]

For i = 1 To 220

For j = 1 To 500

If Sheets(Sht).Cells(j + 2, 1).Value = Sheets("Sym").Cells(i + 2, 29) And Sheets(Sht).Cells(j + 2, 6).Value = Sheets("Sym").Cells(i + 2, 31) Then Sheets(Sht).Cells(j + 2, 22).Value = Sheets("Sym").Cells(i + 2, 33).Value
If Sheets(Sht).Cells(j + 2, 1).Value = Sheets("Sym").Cells(i + 2, 29) And Sheets(Sht).Cells(j + 2, 6).Value = Sheets("Sym").Cells(i + 2, 31) Then Sheets(Sht).Cells(j + 2, 23).Value = Sheets("Sym").Cells(i + 2, 34).Value

Next j

Next i


End Sub


Sub Total()

Sheets("NM830").Select
FinalRow = Cells(Rows.Count, 22).End(xlUp).Row
         
         
    For i = 3 To FinalRow
    
    Cells(i, 1).NumberFormat = "@"
    Dim num As String
    num = Cells(i, 1)
    Cells(i, 1).Value = Right("0000000" & num, 7)
    
        
    If Right(Cells(i, 4).Value, 1) = "Y" Then Cells(i, 4).Value = Left(Cells(i, 4).Value, 3)
    If Right(Cells(i, 4).Value, 1) = "M" Then Cells(i, 4).Value = Left(Cells(i, 4).Value, 3) + (Mid(Cells(i, 4).Value, 5, 3) / 12)
    
    
    If Cells(i, 17).Value = "XE22" Then Cells(i, 17).Value = "Infi"
    If Cells(i, 17).Value = "XE42" Then Cells(i, 17).Value = "NMCT670"
    If Cells(i, 17).Value = "SYMBIAE1429" Then Cells(i, 17).Value = "Sym"
    If Cells(i, 17).Value = "XE12" Then Cells(i, 17).Value = "NM830"
     
    Next i

MsgBox "Complete"

End Sub

 


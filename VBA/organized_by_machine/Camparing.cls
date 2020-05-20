VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ex()


For i = 1 To 50

For j = 1 To 50

If Cells(j, i).Value = 0# Then Cells(j, i).Value = ""

Next j

Next i



End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockmaster()
Dim volume As Double
Dim opener As Double
Dim closer As Double

Dim lastrow As Long
Dim a As Integer
a = 2
' lastrow = Cells(Rows.Count, 1).End(xlUp).Row

volume = 0
For i = 2 To 80000
opener = Cells(i, 3).Value
If Cells(i + 1, 1) <> Cells(i, 1) Then
closer = Cells(i, 6).Value

Cells(a, 10).Value = closer - opener
Cells(a, 9).Value = Cells(i, 1).Value
End If

Next i
End Sub

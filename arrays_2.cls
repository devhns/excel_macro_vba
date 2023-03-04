VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub diziler2()
Dim dizim2(1 To 20) As Variant
Dim i As Integer

For i = 1 To 20
dizim2(i) = Cells(i, 1)
Next i
'MsgBox dizim2(11)
'MsgBox LBound(dizim2) - min index degeri
'MsgBox UBound(dizim2) - max index degeri
MsgBox "dizinin en buyuk degeri " & Excel.WorksheetFunction.Max(dizim2)
MsgBox "dizinin en buyuk degeri " & Excel.WorksheetFunction.Min(dizim2)
MsgBox "dizinin en buyuk degeri " & Excel.WorksheetFunction.Average(dizim2)
MsgBox "dizinin en buyuk degeri " & Excel.WorksheetFunction.Sum(dizim2)
End Sub


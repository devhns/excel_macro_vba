VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Range("C2:D100").Clear
Dim ws As Worksheet
Set ws = Sheets(1)
tarih = Range("A2")

sonsatir = ws.Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To sonsatir
sonsatir2 = Cells(Rows.Count, "c").End(xlUp).Row + 1
    If ws.Cells(i, 2) = tarih Then
    Cells(sonsatir2, 3) = ws.Cells(i, 1)
    Cells(sonsatir2, 4) = ws.Cells(i, 3)
    End If
Next i

End Sub

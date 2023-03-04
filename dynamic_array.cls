VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub dinamikdizi()
Dim sayi_dizisi()  As Integer
sonsatir = Cells(Rows.Count, "a").End(xlUp).Row

ReDim sayi_dizisi(1 To sonsatir)
For i = 1 To sonsatir
sayi_dizisi(i) = Range("A" & i)
Next i

End Sub

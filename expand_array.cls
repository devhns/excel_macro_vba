VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub dizi_genislet()
Dim sayi_dizisi()  As Integer
sonsatir = Cells(Rows.Count, "a").End(xlUp).Row

ReDim sayi_dizisi(1 To sonsatir)
For i = 1 To sonsatir
sayi_dizisi(i) = Cells(i, 1)
Next i

'sum < 500 ise eleman sayisina 1 arttir / deger olarak cell'e 0 atar
dizi_toplam = Excel.WorksheetFunction.Sum(sayi_dizisi)
If dizi_toplam < 500 Then
ReDim Preserve sayi_dizisi(1 To sonsatir + 1)
End If
End Sub

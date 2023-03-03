Attribute VB_Name = "Module1"
'excel fonksiyonlar�n�n makro ici kullanimi
Sub sumifs()

Dim aranan As Range
Dim alan As Range
Dim kriter As Range

Set aranan = Range("F7")
Set toplanan = Range("C:C")
Set kriter = Range("D:D")

sonuc = Excel.WorksheetFunction.sumifs(toplanan, kriter, aranan)
'WorksheetFunction. ile t�m kullan�labilir fonksiyonlara erisilebilir
Range("H7") = sonuc

End Sub

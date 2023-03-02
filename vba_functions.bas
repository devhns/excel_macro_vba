Attribute VB_Name = "Module1"
'VBA fonskiyonlarý
Public Sub vba_functions()
Range("A2") = VBA.Date 'guncel tarih
Range("A3") = VBA.Time 'guncel saat
Range("A4") = VBA.Now 'guncel longtime formatý
Range("A5") = Format(Range("C5"), "#;") 'istenilen formatý uygulama
Range("A6") = Len(Range("C5")) 'degerin uzunlugunu alma
Range("A7") = VBA.StrConv(Range("C5"), vbLowerCase) 'String bicimini degistirme
Range("A8") = StrReverse(Range("C5")) 'hücredeki stringin tersini alma
'Cint / Cstr /Cdate -- Hucre deðerinin ham formatýný Integer,String,Date tipine çevirmeyi saðlar.
End Sub

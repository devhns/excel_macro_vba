Attribute VB_Name = "Module1"
'VBA fonskiyonlar�
Public Sub vba_functions()
Range("A2") = VBA.Date 'guncel tarih
Range("A3") = VBA.Time 'guncel saat
Range("A4") = VBA.Now 'guncel longtime format�
Range("A5") = Format(Range("C5"), "#;") 'istenilen format� uygulama
Range("A6") = Len(Range("C5")) 'degerin uzunlugunu alma
Range("A7") = VBA.StrConv(Range("C5"), vbLowerCase) 'String bicimini degistirme
Range("A8") = StrReverse(Range("C5")) 'h�credeki stringin tersini alma
'Cint / Cstr /Cdate -- Hucre de�erinin ham format�n� Integer,String,Date tipine �evirmeyi sa�lar.
End Sub

Attribute VB_Name = "Module1"
'sablon word dosyasinda yer imleri kullanilmali
Private Sub CommandButton1_Click()
Dim doc As Word.Document
Set wordapp = CreateObject("word.application")
sablon = "C:\Users\Havva\OneDrive\Masaüstü\sablon.docx"

For i = 2 To 11

Set doc = wordapp.documents.Open(sablon)
doc.bookmarks("isim").Range.InsertAfter Cells(i, 1)
doc.bookmarks("bolge").Range.InsertAfter Cells(i, 2)
doc.bookmarks("siralama").Range.InsertAfter Cells(i, 4)
doc.bookmarks("satis").Range.InsertAfter Cells(i, 3)

doc.SaveAs2 "C:\Users\Havva\OneDrive\Masaüstü\" & Cells(i, 1).Text
doc.Close
Next i
End Sub

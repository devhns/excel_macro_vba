VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'sablon word dosyasinda yer imleri kullanilmali
Private Sub CommandButton1_Click()
Dim doc As Word.Document
Set wordapp = CreateObject("word.application")
sablon = "C:\Users\Nadide\Desktop\sablon.docx"

For i = 2 To 11

Set doc = wordapp.documents.Open(sablon)
doc.bookmarks("isim").Range.InsertAfter Cells(i, 1)
doc.bookmarks("bolge").Range.InsertAfter Cells(i, 2)
doc.bookmarks("siralama").Range.InsertAfter Cells(i, 4)
doc.bookmarks("satis").Range.InsertAfter Cells(i, 3)

doc.SaveAs2 "C:\Users\Nadide\Desktop\" & Cells(i, 1).Text
doc.Close
Next i
End Sub

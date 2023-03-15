VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'button atamasi olmadan modul run edince sorun cikarabiliyor
Dim verial As Variant
Dim kitap As Workbook
Application.ScreenUpdating = False
'Kod bitimi sorun cikarmamasi icin True cekmeyi unutma

verial = Application.GetOpenFilename(Title:="Dosya Seç", FileFilter:="Excel Files(*.xlsx*),*xlsx*")
If verial <> False Then
Set kitap = Workbooks.Open(verial)
kitap.Sheets("ham").Range("A1:C15").Copy
Range("A1:c15").PasteSpecial xlPasteValues
kitap.Close

End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton2_Click()

'secilen klasordeki dosyalar arasinda dongusel islem yapma
'ya da secilen bir dosyanin aktarilmasi icin kullanilir

'ornegin birden cok dosyayý tek dosyada birden fazla sayfaya toplamak
'bulundugumuz exceli farklý sekilde kaydetmek icin msoFileDialogSaveAd kullanabiliriz.

Dim dosya As String
Dim klasor_liste As String
Dim kitap As Workbook

With Application.FileDialog(msoFileDialogFolderPicker)
.Title = "Birleþtirilecek dosyalarýn olduðu klasörü seçin"
.ButtonName = "Dosya Seç"

If .Show = 0 Then
Exit Sub
Else
dosya = .SelectedItems(1) & "\"
End If

klasor_liste = Dir(dosya & "*.xls*")
Do Until klasor_liste = ""
Set kitap = Workbooks.Open(dosya & klasor_liste)
kitap.Sheets(1).Copy after:=ThisWorkbook.Sheets(1)
kitap.Close
klasor_liste = Dir
Loop


End With

End Sub

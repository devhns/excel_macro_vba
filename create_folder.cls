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
Dim path As String
Dim folder As String
path = "C:\Users\Havva\OneDrive\Masaüstü\denemee"
folder = Dir(path, vbDirectory)

If folder = vbNullString Then
cevap = MsgBox("Yeni klasör oluþturulsun mu?", vbYesNo)
      If cevap = vbYes Then
      MkDir (path)
      Else
      Exit Sub
      End If
End If
End Sub

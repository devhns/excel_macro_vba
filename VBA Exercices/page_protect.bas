Attribute VB_Name = "Module1"
Sub auto_close()
sifre1 = InputBox("�ifrenizi Girin", "�ifre")
sifre2 = InputBox("�ifrenizi Tekrarlay�n", "Sifre")

If sifre1 <> sifre2 Then Exit Sub

Dim sekme As Worksheet
For Each sekme In Worksheets
sekme.Protect Password:=sifre1
'sekme.Unprotect Password:=sifre1
Next sekme
End Sub

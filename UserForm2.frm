VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Veri Giriþ Formu"
   ClientHeight    =   5376
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7860
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

'Bos birakilmama zorunlulugu icin
'If me.TextBox1 = vbNullString Then
'MsgBox ...

Dim ws As Worksheet
Dim egitim As String
Dim medeni As String

Set ws = Sheets("Rapor")
sonsatir = Cells(Rows.Count, "a").End(xlUp).Row + 1

If Me.OptionButton1 = True Then
egitim = "Doktora"
ElseIf Me.OptionButton2 = True Then
egitim = "Master"
ElseIf Me.OptionButton3 = True Then
egitim = "Üniversite"
ElseIf Me.OptionButton4 = True Then
egitim = "Lise"
ElseIf Me.OptionButton5 = True Then
egitim = "Ortaöðretim"
End If

If Me.OptionButton6 = True Then
medeni = "Evli"
ElseIf Me.OptionButton7 = True Then
medeni = "Bekar"
End If

With ws
.Cells(sonsatir, 1) = Me.TextBox1.Text
.Cells(sonsatir, 2) = Me.TextBox2.Text
.Cells(sonsatir, 3) = Me.TextBox3.Text
.Cells(sonsatir, 4) = egitim
.Cells(sonsatir, 5) = medeni
.Cells(sonsatir, 6) = Me.ComboBox1
.Cells(sonsatir, 7) = Me.TextBox4.Text
End With

'kayýt alýndýktan sonra formu temizle
Me.TextBox1 = vbNullString
Me.TextBox2 = vbNullString
Me.TextBox3 = vbNullString
Me.TextBox4 = vbNullString
Me.ComboBox1.ListIndex = -1


End Sub

Private Sub ScrollBar1_Change()
Me.TextBox3.Text = Me.ScrollBar1.Value
End Sub

Private Sub UserForm_Initialize()
Me.ComboBox1.List = [iller!A1:A14].Value
End Sub


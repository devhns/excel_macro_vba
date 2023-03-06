VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} textbox_check 
   Caption         =   "UserForm1"
   ClientHeight    =   3348
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5244
   OleObjectBlob   =   "textbox_check.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "textbox_check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()
If Not IsNumeric(Me.TextBox1.Value) Then
MsgBox "sadece sayý girebilirsiniz"
End If
End Sub

Private Sub TextBox2_Change()
If IsNumeric(Me.TextBox2.Value) Then
MsgBox "sadece metin girebilirsiniz"
End If
End Sub

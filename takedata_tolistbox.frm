VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5064
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7488
   OleObjectBlob   =   "takedata_tolistbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
With Me.ListBox1
.ColumnCount = 3
.ColumnHeads = True
.RowSource = "Sayfa1!A2:C6"
.ColumnWidths = "50;50;50"
End With
End Sub

Private Sub ListBox1_Click()
For i = 0 To Me.ListBox1.ListCount - 1
      If Me.ListBox1.Selected(i) = True Then
      Me.TextBox1 = Me.ListBox1.List(i, 1)
      End If
Next i
End Sub

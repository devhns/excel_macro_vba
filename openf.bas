Attribute VB_Name = "Module1"
Sub openf()

Dim file_name As String
file_name = VBA.FileSystem.Dir("C:\Users\Havva\OneDrive\Masaüstü\deneme.xlsx")
If file_name = vbNullString Then
MsgBox "file not found"
Else
Workbooks.Open ("C:\Users\Havva\OneDrive\Masaüstü\deneme.xlsx")
End If

End Sub

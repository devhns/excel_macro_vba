Attribute VB_Name = "Module1"
Sub tdim_array()
Dim tdimarray(1 To 14, 1 To 2) As Variant

For i = 1 To 14
      For y = 1 To 2
      tdimarray(i, y) = Cells(i, y).Value
      Next y
Next i

End Sub


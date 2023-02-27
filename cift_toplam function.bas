Attribute VB_Name = "Module1"
Function cift_toplam(st As Range, finish As Range)
ct = 0

For i = st To finish
      If i Mod 2 = 0 Then
      ct = ct + 1
      End If
Next i
End Function

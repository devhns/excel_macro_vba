VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub diziler4()
Dim isimler As Variant
Dim filtre As Variant
isimler = Array("nur", "cem", "baris", "hilay", "mert")
filtre = Filter(isimler, "a", True, vbBinaryCompare)
'vbBinaryCompare - case sensitive
'vbTextCompare - not cs
End Sub


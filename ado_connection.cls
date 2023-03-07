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
Dim baglanti As New ADODB.Connection
Dim rs As New ADODB.Recordset

yol = "C:\Users\Nadide\Desktop\ado ham verisi.xlsx"

baglanti.Open "provider=microsoft.ace.oledb.12.0;data source=" & _
yol & ";extended properties=""Excel 12.0;hdr=yes"""

sorgu = "select * from [ham$]"
rs.Open sorgu, baglanti, adOpenKeyset, adLockPessimistic

a = 1
For Each baslik In rs.Fields
Cells(1, a) = baslik.Name
a = a + 1
Next baslik
Range("A2").CopyFromRecordset rs

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'eger ki dizi indexim 0dan degil 1den baslasin istersem
Option Base 1
Sub diziler()
'dizim(10) as Variant 11 elemanlýk bir dizi (index 0'dan basladigi icin)
'dizim (1 to 10) as Variant - 10 elemanlýk dizi, index 1'den baslasin dedik
'dizim (2,5) as string - 2 boyutlu dizi / matris
'dizim (2,5,3) as Integer - 3 boyutlu dizi
dizim = Array("elma", "armut", "kiraz", "karpuz")
'pyda oldugu gibi eleman indexi 0 dan baslar
End Sub

Attribute VB_Name = "Carteles"
'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit
'Carteles
Public cartel As Boolean
Public Leyenda As String
Public GrhCartel As Integer



Sub InitCartel(Ley As String, grh As Integer)

If Not cartel Then
    Leyenda = Ley
    GrhCartel = grh
    cartel = True
Else
    Exit Sub
End If
End Sub


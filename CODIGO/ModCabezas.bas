Attribute VB_Name = "ModCabezas"
'RevolucionAo 1.0
'Pablo Mercavides
Option Explicit

Public MiCabeza As Integer

Private Sub DrawGrafico(grh As grh, ByVal x As Byte, ByVal y As Byte)
If grh.GrhIndex <= 0 Then Exit Sub
'Call Draw_Grh_Picture(grh.GrhIndex, frmCrearPersonaje.PlayerView, -6, -13, False, 0, 1)

End Sub

Sub DibujarCPJ(ByVal MyHead As Long, Optional ByVal Heading As Byte = 3)

CPHead = MyHead
Dim grh As grh
grh = HeadData(MyHead).Head(Heading)
'Call DrawGrafico(grh, 0, 0)
End Sub

Sub DameOpciones()

Dim i As Integer
If frmCrearPersonaje.lstGenero.ListIndex < 0 Or frmCrearPersonaje.lstRaza.ListIndex < 0 Then
frmCrearPersonaje.Cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.ListIndex <> -1 And frmCrearPersonaje.lstRaza.ListIndex <> -1 Then
frmCrearPersonaje.Cabeza.Enabled = True
End If

frmCrearPersonaje.Cabeza.Clear

Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
    Case "Hombre"
Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
    Case "Humano"
        Call DibujarCPJ(1)
        MiCabeza = 1
        For i = 1 To 41
            frmCrearPersonaje.Cabeza.AddItem i
        Next i

    Case "Elfo"
        Call DibujarCPJ(101)
        MiCabeza = 101
        For i = 101 To 132
        frmCrearPersonaje.Cabeza.AddItem i
        Next i
    
    Case "Elfo Drow"
        Call DibujarCPJ(200)
        MiCabeza = 200
        For i = 200 To 229
        frmCrearPersonaje.Cabeza.AddItem i
        Next i
    
    Case "Enano"
        Call DibujarCPJ(300)
        MiCabeza = 300
        For i = 300 To 329
        frmCrearPersonaje.Cabeza.AddItem i
        Next i
    
    Case "Gnomo"
        Call DibujarCPJ(400)
        MiCabeza = 400
        For i = 400 To 429
        frmCrearPersonaje.Cabeza.AddItem i
        Next i
    Case "Orco"
        Call DibujarCPJ(500)
        MiCabeza = 500
        For i = 500 To 529
        frmCrearPersonaje.Cabeza.AddItem i
        Next i
    
    Case Else
        Call DibujarCPJ(1)
        MiCabeza = 1
        UserHead = 1
    
    End Select
    
    Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Call DibujarCPJ(50)
                MiCabeza = 50
                For i = 50 To 80
                frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Elfo"
                Call DibujarCPJ(150)
                MiCabeza = 150
                For i = 150 To 179
                frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Elfo Drow"
                Call DibujarCPJ(250)
                MiCabeza = 250
                For i = 250 To 279
                frmCrearPersonaje.Cabeza.AddItem i
                Next i
                
            Case "Enano"
                Call DibujarCPJ(350)
                MiCabeza = 350
                For i = 350 To 379
                frmCrearPersonaje.Cabeza.AddItem i
                                Next i
            Case "Gnomo"
                Call DibujarCPJ(450)
                MiCabeza = 450
                For i = 450 To 479
                frmCrearPersonaje.Cabeza.AddItem i

                Next i
            Case "Orco"
                Call DibujarCPJ(550)
                MiCabeza = 550
                For i = 550 To 579
                frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case Else
                MiCabeza = 50
                Call DibujarCPJ(50)
                frmCrearPersonaje.Cabeza.AddItem "50"
                
            End Select
    End Select



Rem frmCrearPersonaje.PlayerView.Cls

End Sub


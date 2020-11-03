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
frmCrearPersonaje.cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.ListIndex <> -1 And frmCrearPersonaje.lstRaza.ListIndex <> -1 Then
frmCrearPersonaje.cabeza.Enabled = True
End If

frmCrearPersonaje.cabeza.Clear

Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
    Case "Hombre"
Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
    Case "Humano"
        MiCabeza = RandomNumber(1, 41)
        Call DibujarCPJ(MiCabeza)
        For i = 1 To 41
            frmCrearPersonaje.cabeza.AddItem i
        Next i

    Case "Elfo"
        MiCabeza = RandomNumber(101, 132)
        Call DibujarCPJ(MiCabeza)
        For i = 101 To 132
        frmCrearPersonaje.cabeza.AddItem i
        Next i
    
    Case "Elfo Drow"
        MiCabeza = RandomNumber(200, 229)
        Call DibujarCPJ(MiCabeza)
        For i = 200 To 229
        frmCrearPersonaje.cabeza.AddItem i
        Next i
    
    Case "Enano"
        MiCabeza = RandomNumber(300, 329)
        Call DibujarCPJ(MiCabeza)
        For i = 300 To 329
        frmCrearPersonaje.cabeza.AddItem i
        Next i
    
    Case "Gnomo"
        MiCabeza = RandomNumber(400, 429)
        Call DibujarCPJ(MiCabeza)
        For i = 400 To 429
        frmCrearPersonaje.cabeza.AddItem i
        Next i
    Case "Orco"
        MiCabeza = RandomNumber(500, 529)
        Call DibujarCPJ(MiCabeza)
        For i = 500 To 529
        frmCrearPersonaje.cabeza.AddItem i
        Next i
    
    Case Else
        MiCabeza = 1
        Call DibujarCPJ(MiCabeza)
        UserHead = 1
    
    End Select
    
    Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                MiCabeza = RandomNumber(50, 80)
                Call DibujarCPJ(MiCabeza)
                For i = 50 To 80
                frmCrearPersonaje.cabeza.AddItem i
                Next i
            Case "Elfo"
                MiCabeza = RandomNumber(150, 179)
                Call DibujarCPJ(MiCabeza)
                For i = 150 To 179
                frmCrearPersonaje.cabeza.AddItem i
                Next i
            Case "Elfo Drow"
                MiCabeza = RandomNumber(250, 279)
                Call DibujarCPJ(MiCabeza)
                For i = 250 To 279
                frmCrearPersonaje.cabeza.AddItem i
                Next i
                
            Case "Enano"
                MiCabeza = RandomNumber(350, 379)
                Call DibujarCPJ(MiCabeza)
                For i = 350 To 379
                frmCrearPersonaje.cabeza.AddItem i
                                Next i
            Case "Gnomo"
                MiCabeza = RandomNumber(450, 479)
                Call DibujarCPJ(MiCabeza)
                For i = 450 To 479
                frmCrearPersonaje.cabeza.AddItem i

                Next i
            Case "Orco"
                MiCabeza = RandomNumber(550, 579)
                Call DibujarCPJ(MiCabeza)
                For i = 550 To 579
                frmCrearPersonaje.cabeza.AddItem i
                Next i
                
            Case Else
                MiCabeza = 50
                Call DibujarCPJ(50)
                frmCrearPersonaje.cabeza.AddItem "50"
                
            End Select
    End Select



Rem frmCrearPersonaje.PlayerView.Cls

End Sub


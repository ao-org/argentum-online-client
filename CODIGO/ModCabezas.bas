Attribute VB_Name = "ModCabezas"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Public MiCabeza As Integer

Private Sub DrawGrafico(grh As grh, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo DrawGrafico_Err
    

    If grh.GrhIndex <= 0 Then Exit Sub
    'Call Draw_Grh_Picture(grh.GrhIndex, frmCrearPersonaje.PlayerView, -6, -13, False, 0, 1)

    
    Exit Sub

DrawGrafico_Err:
    Call RegistrarError(Err.number, Err.Description, "ModCabezas.DrawGrafico", Erl)
    Resume Next
    
End Sub

Sub DibujarCPJ(ByVal MyHead As Long, Optional ByVal Heading As Byte = 3)
    
    On Error GoTo DibujarCPJ_Err
    

    CPHead = MyHead

    Dim grh As grh

    grh = HeadData(MyHead).Head(Heading)

    'Call DrawGrafico(grh, 0, 0)
    
    Exit Sub

DibujarCPJ_Err:
    Call RegistrarError(Err.number, Err.Description, "ModCabezas.DibujarCPJ", Erl)
    Resume Next
    
End Sub

Sub DameOpciones()
    
    On Error GoTo DameOpciones_Err
    

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
                    MiCabeza = RandomNumber(1, 41)
                    Call DibujarCPJ(MiCabeza)

                    For i = 1 To 41
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i

                Case "Elfo"
                    MiCabeza = RandomNumber(101, 132)
                    Call DibujarCPJ(MiCabeza)

                    For i = 101 To 132
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i
    
                Case "Elfo Oscuro"
                    MiCabeza = RandomNumber(200, 229)
                    Call DibujarCPJ(MiCabeza)

                    For i = 200 To 229
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i
    
                Case "Enano"
                    MiCabeza = RandomNumber(300, 329)
                    Call DibujarCPJ(MiCabeza)

                    For i = 300 To 329
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i
    
                Case "Gnomo"
                    MiCabeza = RandomNumber(400, 429)
                    Call DibujarCPJ(MiCabeza)

                    For i = 400 To 429
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i

                Case "Orco"
                    MiCabeza = RandomNumber(500, 529)
                    Call DibujarCPJ(MiCabeza)

                    For i = 500 To 529
                        frmCrearPersonaje.Cabeza.AddItem i
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
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i

                Case "Elfo"
                    MiCabeza = RandomNumber(150, 179)
                    Call DibujarCPJ(MiCabeza)

                    For i = 150 To 179
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i

                Case "Elfo Oscuro"
                    MiCabeza = RandomNumber(250, 279)
                    Call DibujarCPJ(MiCabeza)

                    For i = 250 To 279
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i
                
                Case "Enano"
                    MiCabeza = RandomNumber(350, 379)
                    Call DibujarCPJ(MiCabeza)

                    For i = 350 To 379
                        frmCrearPersonaje.Cabeza.AddItem i
                    Next i

                Case "Gnomo"
                    MiCabeza = RandomNumber(450, 479)
                    Call DibujarCPJ(MiCabeza)

                    For i = 450 To 479
                        frmCrearPersonaje.Cabeza.AddItem i

                    Next i

                Case "Orco"
                    MiCabeza = RandomNumber(550, 579)
                    Call DibujarCPJ(MiCabeza)

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

    
    Exit Sub

DameOpciones_Err:
    Call RegistrarError(Err.number, Err.Description, "ModCabezas.DameOpciones", Erl)
    Resume Next
    
End Sub


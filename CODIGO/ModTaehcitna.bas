Attribute VB_Name = "ModTaehcitna"
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
Private Const MAX_COMPROBACIONES As Byte = 4
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private ContadorMacroClicks(1 To MAX_COMPROBACIONES) As Position
' Historial de los ultimos tiles absolutos clickeados (coordenada del mapa).
Private ContadorMacroTilesAbsolutos(1 To MAX_COMPROBACIONES) As Position
' Historial de tiles relativos al usuario (tile objetivo - posicion del personaje).
Private ContadorMacroTilesRelativos(1 To MAX_COMPROBACIONES) As Position
Public LastSentPosX                                  As Integer
Public LastSentPosY                                  As Integer

Public Function ComprobarPosibleMacro(ByVal mouseX As Integer, ByVal mouseY As Integer, Optional ByVal tileX As Integer = -1, Optional ByVal tileY As Integer = -1) As Boolean
    Call CopyMemory(ContadorMacroClicks(2), ContadorMacroClicks(1), Len(ContadorMacroClicks(1)) * (MAX_COMPROBACIONES - 1))
    ContadorMacroClicks(1).x = mouseX
    ContadorMacroClicks(1).y = mouseY

    ' Si recibimos tile valido, tambien registramos el objetivo en coordenadas de mapa
    ' y en coordenadas relativas al personaje para detectar patron repetitivo aunque se mueva el mouse.
    If tileX >= 0 And tileY >= 0 Then
        Call CopyMemory(ContadorMacroTilesAbsolutos(2), ContadorMacroTilesAbsolutos(1), Len(ContadorMacroTilesAbsolutos(1)) * (MAX_COMPROBACIONES - 1))
        ContadorMacroTilesAbsolutos(1).x = tileX
        ContadorMacroTilesAbsolutos(1).y = tileY

        Call CopyMemory(ContadorMacroTilesRelativos(2), ContadorMacroTilesRelativos(1), Len(ContadorMacroTilesRelativos(1)) * (MAX_COMPROBACIONES - 1))
        ContadorMacroTilesRelativos(1).x = tileX - UserPos.x
        ContadorMacroTilesRelativos(1).y = tileY - UserPos.y
    End If

    Dim i As Byte

    ' Deteccion existente: 4 clicks seguidos en el mismo pixel de mouse.
    For i = 1 To MAX_COMPROBACIONES
        If ContadorMacroClicks(i).x <> mouseX Or ContadorMacroClicks(i).y <> mouseY Then
            Exit For
        End If
    Next i
    If i > MAX_COMPROBACIONES Then
        ComprobarPosibleMacro = True
        Call generarLogMacrero
        Exit Function
    End If

    ' Nuevo control adicional: 4 clicks seguidos al mismo tile absoluto o al mismo tile relativo al usuario.
    If tileX >= 0 And tileY >= 0 Then
        For i = 1 To MAX_COMPROBACIONES
            If ContadorMacroTilesAbsolutos(i).x <> tileX Or ContadorMacroTilesAbsolutos(i).y <> tileY Then
                Exit For
            End If
        Next i
        If i > MAX_COMPROBACIONES Then
            ComprobarPosibleMacro = True
            Call generarLogMacrero
            Exit Function
        End If

        For i = 1 To MAX_COMPROBACIONES
            If ContadorMacroTilesRelativos(i).x <> (tileX - UserPos.x) Or ContadorMacroTilesRelativos(i).y <> (tileY - UserPos.y) Then
                Exit For
            End If
        Next i
        If i > MAX_COMPROBACIONES Then
            ComprobarPosibleMacro = True
            Call generarLogMacrero
            Exit Function
        End If
    End If

    ComprobarPosibleMacro = False
End Function

Private Sub generarLogMacrero()
    ' Mantenemos el mismo tipo de log para no cambiar el protocolo ni la logica del servidor/panel GM.
    Call WriteLogMacroClickHechizo(tMacro.inasistidoPosFija)
End Sub

Public Sub CountPacketIterations(ByRef packetControl As t_packetControl, ByVal expectedAverage As Double)
    Dim delta       As Long
    Dim actualcount As Long
    actualcount = GetTickCount()
    delta = actualcount - packetControl.last_count
    If delta < 40 Then Exit Sub
    packetControl.last_count = actualcount
    Call alterIndex(packetControl)
    packetControl.iterations(10) = delta
    Dim percentageDiff As Double, average As Double
    percentageDiff = getPercentageDiff(packetControl)
    average = getAverage(packetControl)
    ' frmdebug.add_text_tracebox "Delta: " & delta & " Average: " & average
    If percentageDiff < 5 Then
        'frmdebug.add_text_tracebox "DIFF: " & getPercentageDiff(packetControl)
        'Call AddtoRichTextBox(frmMain.RecTxt, "DIFF: " & getPercentageDiff(packetControl), 255, 200, 0, True)
        Call WriteRepeatMacro
        'frmdebug.add_text_tracebox "DIFF: " & getPercentageDiff(packetControl)
    End If
    If average > 20 And average < expectedAverage Then
        Call WriteRepeatMacro
    End If
End Sub

Private Function getPercentageDiff(ByRef packetControl As t_packetControl) As Double
    Dim i As Long, min As Long, max As Long
    min = packetControl.iterations(1)
    max = packetControl.iterations(1)
    Dim count As Long
    For i = 1 To 10
        If packetControl.iterations(i) < min Then
            min = packetControl.iterations(i)
        End If
        If packetControl.iterations(i) > max Then
            max = packetControl.iterations(i)
        End If
    Next i
    getPercentageDiff = 100 - ((min * 100) / max)
End Function

Private Function getAverage(ByRef packetControl As t_packetControl) As Double
    Dim i As Long, suma As Long
    For i = 1 To UBound(packetControl.iterations)
        suma = suma + packetControl.iterations(i)
    Next i
    getAverage = suma / UBound(packetControl.iterations)
End Function

Private Sub alterIndex(ByRef packetControl As t_packetControl)
    Dim i As Long
    For i = 1 To 10 ' packetControl.cant_iterations
        If i < 10 Then 'packetControl.cant_iterations Then
            packetControl.iterations(i) = packetControl.iterations(i + 1)
        End If
    Next i
End Sub

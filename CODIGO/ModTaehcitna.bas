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
' Obtiene una muestra independiente del cursor desde Windows para no confiar en coordenadas cacheadas.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Convierte la muestra global de Windows a coordenadas locales del control de render.
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private ContadorMacroClicks(1 To MAX_COMPROBACIONES) As Position
Public LastSentPosX                                  As Integer
Public LastSentPosY                                  As Integer

Public Function ComprobarPosibleMacro(ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    Call CopyMemory(ContadorMacroClicks(2), ContadorMacroClicks(1), Len(ContadorMacroClicks(1)) * (MAX_COMPROBACIONES - 1))
    ContadorMacroClicks(1).x = mouseX
    ContadorMacroClicks(1).y = mouseY
    Dim i As Byte
    For i = 1 To MAX_COMPROBACIONES
        If ContadorMacroClicks(i).x <> mouseX Or ContadorMacroClicks(i).y <> mouseY Then
            ComprobarPosibleMacro = False
            Exit Function
        End If
    Next i
    ComprobarPosibleMacro = True
    Call generarLogMacrero
End Function

Public Function CoincideObjetivoHechizoConMouse(ByVal objetivoX As Byte, ByVal objetivoY As Byte) As Boolean
    ' Toma la posicion real del cursor desde Windows para no confiar en valores modificables del cliente.
    Dim cursorReal As POINTAPI
    Dim mouseTileX As Byte
    Dim mouseTileY As Byte

    ' Rechaza el envio si Windows no puede informar o convertir la posicion actual del cursor.
    If GetCursorPos(cursorReal) = 0 Then Exit Function
    If ScreenToClient(frmMain.renderer.hWnd, cursorReal) = 0 Then Exit Function

    ' Rechaza acciones dirigidas cuando el cursor real no se encuentra dentro del render.
    If cursorReal.x < 0 Or cursorReal.x > frmMain.renderer.ScaleWidth Then Exit Function
    If cursorReal.y < 0 Or cursorReal.y > frmMain.renderer.ScaleHeight Then Exit Function

    ' Convierte exclusivamente la muestra independiente del cursor al tile real señalado.
    Call ConvertCPtoTP(cursorReal.x, cursorReal.y, mouseTileX, mouseTileY)

    ' Solo permite enviar el hechizo cuando su objetivo coincide con el tile real del cursor.
    CoincideObjetivoHechizoConMouse = (mouseTileX = objetivoX And mouseTileY = objetivoY)

    ' Informa al servidor la diferencia de coordenadas para que pueda registrar el intento sospechoso.
    If Not CoincideObjetivoHechizoConMouse Then Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
End Function

Private Sub generarLogMacrero()
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

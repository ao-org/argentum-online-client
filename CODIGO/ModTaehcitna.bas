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
Private Const MAX_DESVIOS_COORDENADAS_CONSECUTIVOS As Byte = 2
Private Const TOLERANCIA_PIXELES_COORDENADAS As Integer = 6
Private Const MAX_CLICKS_RAPIDOS_CONSECUTIVOS As Byte = 3
Private Const TOLERANCIA_CLICKS_RAPIDOS_MS As Long = 15
' Obtiene una muestra independiente del cursor desde Windows para no confiar en coordenadas cacheadas.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Convierte la muestra global de Windows a coordenadas locales del control de render.
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private ContadorMacroClicks(1 To MAX_COMPROBACIONES) As Position
Private DesviosCoordenadasConsecutivos              As Byte
Private ClicksRapidosCoordenadasConsecutivos        As Byte
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

Public Function CoincideObjetivoHechizoConMouse(ByVal objetivoX As Byte, ByVal objetivoY As Byte, Optional ByVal clickX As Integer = -1, Optional ByVal clickY As Integer = -1) As Boolean
    ' Toma la posicion real del cursor desde Windows para no confiar en valores modificables del cliente.
    Dim cursorReal As POINTAPI
    Dim mouseTileX As Byte
    Dim mouseTileY As Byte
    Dim clickTileX As Byte
    Dim clickTileY As Byte
    Dim clickValido As Boolean

    ' Si recibimos la muestra del evento original, la usamos como respaldo contra falsos positivos
    ' producidos por mover el mouse inmediatamente despues del click.
    If clickX >= 0 And clickY >= 0 Then
        If clickX >= 0 And clickX <= frmMain.renderer.ScaleWidth And clickY >= 0 And clickY <= frmMain.renderer.ScaleHeight Then
            Call ConvertCPtoTP(clickX, clickY, clickTileX, clickTileY)
            clickValido = (clickTileX = objetivoX And clickTileY = objetivoY)
        End If
    End If

    ' Rechaza el envio si Windows no puede informar o convertir la posicion actual del cursor,
    ' salvo que la muestra exacta del evento original coincida con el objetivo.
    If GetCursorPos(cursorReal) = 0 Then
        If clickValido Then CoincideObjetivoHechizoConMouse = True
        Exit Function
    End If
    If ScreenToClient(frmMain.renderer.hWnd, cursorReal) = 0 Then
        If clickValido Then CoincideObjetivoHechizoConMouse = True
        Exit Function
    End If

    ' Rechaza acciones dirigidas cuando el cursor real no se encuentra dentro del render,
    ' salvo que el click original haya sido valido. Esto evita sancionar movimientos posteriores al click.
    If cursorReal.x < 0 Or cursorReal.x > frmMain.renderer.ScaleWidth Then
        If clickValido Then CoincideObjetivoHechizoConMouse = True
        Exit Function
    End If
    If cursorReal.y < 0 Or cursorReal.y > frmMain.renderer.ScaleHeight Then
        If clickValido Then CoincideObjetivoHechizoConMouse = True
        Exit Function
    End If

    ' Convierte exclusivamente la muestra independiente del cursor al tile real señalado.
    Call ConvertCPtoTP(cursorReal.x, cursorReal.y, mouseTileX, mouseTileY)

    ' Permite pequeñas diferencias de pocos pixeles dentro del mismo tile objetivo para evitar falsos
    ' positivos por bordes, DPI o redondeos entre el evento de click y la lectura de Windows.
    If clickValido Or (mouseTileX = objetivoX And mouseTileY = objetivoY) Or EsDesvioMenorDeCoordenadas(cursorReal.x, cursorReal.y, objetivoX, objetivoY) Then
        DesviosCoordenadasConsecutivos = 0
        CoincideObjetivoHechizoConMouse = True
        Exit Function
    End If

    ' No reporta por una unica muestra aislada: puede ser un movimiento legitimo entre el click y la
    ' comprobacion. El log se envia recien ante desvios consecutivos.
    DesviosCoordenadasConsecutivos = DesviosCoordenadasConsecutivos + 1
    If DesviosCoordenadasConsecutivos >= MAX_DESVIOS_COORDENADAS_CONSECUTIVOS Then
        DesviosCoordenadasConsecutivos = 0
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
End Function

Private Function EsDesvioMenorDeCoordenadas(ByVal cursorX As Long, ByVal cursorY As Long, ByVal objetivoX As Byte, ByVal objetivoY As Byte) As Boolean
    Dim objetivoPixelX As Long
    Dim objetivoPixelY As Long

    objetivoPixelX = (objetivoX - UserPos.x + frmMain.renderer.ScaleWidth \ 64) * 32
    objetivoPixelY = (objetivoY - UserPos.y + frmMain.renderer.ScaleHeight \ 64) * 32

    EsDesvioMenorDeCoordenadas = cursorX >= objetivoPixelX - TOLERANCIA_PIXELES_COORDENADAS And _
                                   cursorX < objetivoPixelX + 32 + TOLERANCIA_PIXELES_COORDENADAS And _
                                   cursorY >= objetivoPixelY - TOLERANCIA_PIXELES_COORDENADAS And _
                                   cursorY < objetivoPixelY + 32 + TOLERANCIA_PIXELES_COORDENADAS
End Function

Public Sub RegistrarPosibleMacroCoordenadasPorClickRapido(ByVal diferenciaTicks As Long, ByVal ultimoBoton As Long, ByVal botonActual As Long, ByVal intervaloMinimo As Long)
    Dim umbralReporte As Long

    If diferenciaTicks < 0 Then
        ClicksRapidosCoordenadasConsecutivos = 0
        Exit Sub
    End If

    ' Si intervaloMinimo es 50 ms, no reportamos todo lo menor a 50: restamos 15 ms de tolerancia.
    ' Asi se ignoran falsos positivos de 35 a 49 ms que pueden aparecer por hotkeys, polling o scheduling,
    ' y solo se consideran sospechosas rachas repetidas por debajo de 35 ms.
    umbralReporte = intervaloMinimo - TOLERANCIA_CLICKS_RAPIDOS_MS
    If umbralReporte < 1 Then umbralReporte = 1

    If diferenciaTicks < umbralReporte And ultimoBoton <> botonActual Then
        ClicksRapidosCoordenadasConsecutivos = ClicksRapidosCoordenadasConsecutivos + 1
        If ClicksRapidosCoordenadasConsecutivos >= MAX_CLICKS_RAPIDOS_CONSECUTIVOS Then
            ClicksRapidosCoordenadasConsecutivos = 0
            Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
        End If
    Else
        ClicksRapidosCoordenadasConsecutivos = 0
    End If
End Sub

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
